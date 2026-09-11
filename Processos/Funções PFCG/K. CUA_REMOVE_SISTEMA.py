# -*- coding: utf-8 -*-
"""
Processo: Remoção de Atribuição de Sistema no SAP CUA (SU01)
=============================================================================
Objetivo:
  Percorrer utilizador a utilizador no SAP CUA (sistema SPA, cliente 001) e
  remover a atribuição de um sistema recetor (ex.: S4DCLNT100) na aba 'Sistemas'.

Regra fundamental do negócio:
  Se o sistema alvo NÃO existir no utilizador, NÃO É ERRO.
  Regista-se como 'NÃO EXISTIA' / 'JÁ REMOVIDO' e avança para o próximo.
=============================================================================
"""

from __future__ import annotations

import argparse
import os
import sys
import threading
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Forçar UTF-8 na consola Windows
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def conectar_sap_cua(sistema_alvo: str = "SPA"):
    """
    Liga-se à sessão ativa do SAP GUI Scripting no sistema CUA (SPA).
    Assegura a ligação à WindowStation WinSta0\\default para compatibilidade COM.
    """
    if sys.platform.startswith("win"):
        try:
            import win32service
            hwinsta = win32service.OpenWindowStation("WinSta0", True, 0x037F)
            hwinsta.SetProcessWindowStation()
            hdesk = win32service.OpenDesktop("default", 0, True, 0x01FF)
            hdesk.SetThreadDesktop()
        except Exception:
            pass

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()

    rot = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
    sap_gui_auto = rot.GetROTEntry("SAPGUI")
    if not sap_gui_auto:
        raise RuntimeError("SAP GUI não encontrado no ROT. Certifique-se de que o SAP GUI está aberto e logado.")

    engine = sap_gui_auto.GetScriptingEngine
    if not engine or engine.Children.Count == 0:
        raise RuntimeError("Nenhuma conexão SAP GUI encontrada.")

    for i in range(engine.Children.Count):
        conn = engine.Children(i)
        for j in range(conn.Children.Count):
            sess = conn.Children(j)
            try:
                sys_name = str(getattr(sess.Info, "SystemName", "") or "").strip().upper()
                if sys_name == sistema_alvo.upper():
                    return sess
            except Exception:
                continue

    # Fallback: primeira sessão se só houver uma
    if engine.Children.Count == 1 and engine.Children(0).Children.Count >= 1:
        sess = engine.Children(0).Children(0)
        return sess

    raise RuntimeError(f"Sessão para o sistema {sistema_alvo} não encontrada.")


def remover_sistema_do_utilizador(
    session,
    bname: str,
    sistema_a_remover: str = "S4DCLNT100",
    dry_run: bool = False,
    delay_s: float = 0.3,
) -> Tuple[str, str, List[str]]:
    """
    Acessa SU01 em modo de alteração, seleciona a aba 'Sistemas',
    localiza e remove a linha do sistema alvo.

    Retorna: (status, mensagem, sistemas_encontrados)
      Status possíveis:
        - 'CONCLUIDO': Sistema foi localizado e removido com sucesso
        - 'NAO_EXISTIA': Sistema já não estava atribuído ao utilizador (não é erro!)
        - 'SIMULADO': Modo dry-run, remoção não efetuada
        - 'BLOQUEADO': Utilizador bloqueado por outro operador
        - 'NAO_ENCONTRADO': Utilizador não existe no sistema CUA
        - 'ERRO': Falha técnica durante a execução
    """
    inicio = time.time()
    bname_limpo = str(bname).strip().upper()
    sistema_alvo = str(sistema_a_remover).strip().upper()

    try:
        # 1. Entrar na SU01
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(delay_s)

        # 2. Informar Utilizador
        session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = bname_limpo

        # 3. Entrar em modo alteração (btn[18] = F6 / Alterar)
        session.findById("wnd[0]/tbar[1]/btn[18]").press()
        time.sleep(delay_s + 0.1)

        # Verificar mensagens na barra de status inicial
        sbar = session.findById("wnd[0]/sbar")
        sbar_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
        sbar_text = str(getattr(sbar, "Text", "") or "").strip()

        if sbar_type in ("E", "A"):
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            if "lock" in sbar_text.lower() or "bloque" in sbar_text.lower():
                return "BLOQUEADO", sbar_text, []
            return "NAO_ENCONTRADO", sbar_text, []

        # 4. Selecionar aba Sistemas (tabpSYSTEMS)
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpSYSTEMS").select()
        time.sleep(delay_s)

        # 5. Acessar ALV Grid de Sistemas
        grid = session.findById(
            "wnd[0]/usr/tabsTABSTRIP1/tabpSYSTEMS/"
            "ssubMAINAREA:SAPLSUID_MAINTENANCE:1210/"
            "cntlG_CUA_SYSTEMS_CONTAINER/shellcont/shell"
        )

        row_count = int(getattr(grid, "RowCount", 0))
        linha_alvo = None
        sistemas_atuais = []

        for r in range(row_count):
            sys_val = str(grid.GetCellValue(r, "SUBSYSTEM") or "").strip().upper()
            if sys_val:
                sistemas_atuais.append(sys_val)
                if sys_val == sistema_alvo:
                    linha_alvo = r

        # 6. Se o sistema alvo não existe: NÃO É ERRO!
        if linha_alvo is None:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            msg = f"Sistema {sistema_alvo} não existia no utilizador (sistemas atuais: {', '.join(sistemas_atuais) or 'nenhum'})"
            return "NAO_EXISTIA", msg, sistemas_atuais

        # Modo de simulação (dry-run)
        if dry_run:
            session.findById("wnd[0]/tbar[0]/btn[12]").press()  # Cancelar
            time.sleep(0.2)
            try:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except Exception:
                pass
            return "SIMULADO", f"Sistema {sistema_alvo} detetado na linha {linha_alvo} (remoção simulada)", sistemas_atuais

        # 7. Selecionar linha e premir DEL_LINE
        grid.setCurrentCell(linha_alvo, "SUBSYSTEM")
        grid.selectedRows = str(linha_alvo)
        grid.pressToolbarButton("DEL_LINE")
        time.sleep(delay_s)

        # 8. Tratar popup de confirmação ('Removes a recipient system')
        try:
            wnd1 = session.findById("wnd[1]")
            wnd1.findById("tbar[0]/btn[0]").press()  # Botão Continue (Enter)
            time.sleep(delay_s)
        except Exception:
            pass

        # 9. Gravar alterações (Ctrl+S / VKey 11)
        session.findById("wnd[0]").sendVKey(11)
        time.sleep(delay_s + 0.2)

        # 10. Validar status bar após gravação
        sbar = session.findById("wnd[0]/sbar")
        sbar_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
        sbar_text = str(getattr(sbar, "Text", "") or "").strip()

        # Limpar sessão para /n
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        duracao = time.time() - inicio
        sistemas_restantes = [s for s in sistemas_atuais if s != sistema_alvo]

        if sbar_type in ("E", "A"):
            return "ERRO", f"Erro ao gravar ({sbar_text}) [{duracao:.1f}s]", sistemas_atuais

        msg_ok = sbar_text or f"Sistema {sistema_alvo} removido com sucesso"
        return "CONCLUIDO", f"{msg_ok} (restantes: {', '.join(sistemas_restantes)}) [{duracao:.1f}s]", sistemas_restantes

    except Exception as exc:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except Exception:
            pass
        return "ERRO", f"Exceção técnica: {exc}", []


def processar_lista_utilizadores(
    utilizadores: List[str],
    sistema_a_remover: str = "S4DCLNT100",
    dry_run: bool = False,
    delay_s: float = 0.3,
) -> List[Dict[str, Any]]:
    """
    Executa o procedimento para a lista de utilizadores indicada.
    """
    print("=" * 80)
    print("🚀 INICIANDO REMOÇÃO DE SISTEMA NO CUA (SPA / SU01)")
    print("=" * 80)
    print(f"🎯 Sistema Alvo a Remover: {sistema_a_remover}")
    print(f"👥 Total Utilizadores a Processar: {len(utilizadores)}")
    print(f"🧪 Modo Simulação (Dry-run): {'SIM' if dry_run else 'NÃO (EXECUÇÃO REAL)'}")
    print("=" * 80)

    session = conectar_sap_cua("SPA")
    info = session.Info
    print(f"✅ Conectado ao SAP: {info.SystemName} | Cliente: {info.Client} | Utilizador: {info.User}\n")

    resultados = []
    total = len(utilizadores)

    for idx, bname in enumerate(utilizadores, 1):
        print(f"[{idx:03d}/{total:03d}] A analisar utilizador: {bname.strip().upper()} ...", end=" ", flush=True)
        st, msg, sistemas = remover_sistema_do_utilizador(
            session=session,
            bname=bname,
            sistema_a_remover=sistema_a_remover,
            dry_run=dry_run,
            delay_s=delay_s,
        )

        icone = "✅" if st == "CONCLUIDO" else "ℹ️" if st == "NAO_EXISTIA" else "🧪" if st == "SIMULADO" else "❌"
        print(f"{icone} {st}: {msg}")

        resultados.append({
            "utilizador": bname.strip().upper(),
            "status": st,
            "mensagem": msg,
            "sistemas": sistemas,
            "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        })

    # Resumo Final
    print("\n" + "=" * 80)
    print("📊 RESUMO DA EXECUÇÃO")
    print("=" * 80)
    contagem: Dict[str, int] = {}
    for r in resultados:
        contagem[r["status"]] = contagem.get(r["status"], 0) + 1

    print(f"Total Processados: {len(resultados)}")
    for status, count in sorted(contagem.items()):
        icone = "✅" if status == "CONCLUIDO" else "ℹ️" if status == "NAO_EXISTIA" else "🧪" if status == "SIMULADO" else "❌"
        print(f"  {icone} {status}: {count}")
    print("=" * 80)

    return resultados


def obter_utilizadores_excel(
    caminho_excel: Optional[str] = None,
    departamento: Optional[str] = None,
) -> List[str]:
    """Carrega os utilizadores a partir da folha Proposta Ativa."""
    import importlib.util
    raiz = Path(__file__).resolve().parent.parent.parent
    spec = importlib.util.spec_from_file_location("projeto_perfil", raiz / "Projeto Perfil.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    if not caminho_excel:
        caminho_excel = mod.encontrar_excel_padrao()

    import pandas as pd
    fonte = mod.abrir_excel_seguro(caminho_excel)
    df_pa = pd.read_excel(fonte, sheet_name="Proposta Ativa", header=None)

    dep_filtro = mod.normalizar_texto(departamento) if departamento else ""
    utilizadores = []

    for idx in range(1, len(df_pa)):
        u = str(df_pa.iloc[idx, 0]).strip().upper()
        if not u or u in ("NAN", "NONE", ""):
            continue

        if dep_filtro:
            dep_val = str(df_pa.iloc[idx, 7]).strip() if pd.notna(df_pa.iloc[idx, 7]) else ""
            if mod.normalizar_texto(dep_val) == dep_filtro or dep_filtro in mod.normalizar_texto(dep_val):
                utilizadores.append(u)
        else:
            utilizadores.append(u)

    return sorted(list(dict.fromkeys(utilizadores)))


def main():
    parser = argparse.ArgumentParser(description="Remover sistema recetor no SAP CUA (SU01) utilizador a utilizador")
    parser.add_argument("--users", "-u", nargs="+", help="Lista de utilizadores específicos (ex: S170 S270 S419)")
    parser.add_argument("--departamento", "-d", default=None, help="Nome do departamento no Excel (se omitido, deteta automaticamente o próximo pendente da sheet CONTROLO com STATUS e TIMESTAMP vazios)")
    parser.add_argument("--todos", action="store_true", help="Processar todos os utilizadores da folha Proposta Ativa")
    parser.add_argument("--sistema", "-s", default="S4DCLNT100", help="Sistema a remover (default: S4DCLNT100)")
    parser.add_argument("--dry-run", action="store_true", help="Apenas simular sem gravar")
    parser.add_argument("--delay", type=float, default=0.3, help="Delay entre passos no SAP GUI (segundos)")

    args = parser.parse_args()

    def run_worker():
        if args.users:
            users_alvo = [u.strip().upper() for u in args.users]
        elif args.todos:
            users_alvo = obter_utilizadores_excel(departamento=None)
        else:
            dep = args.departamento
            if not dep:
                import importlib.util
                raiz = Path(__file__).resolve().parent.parent.parent
                spec = importlib.util.spec_from_file_location("projeto_perfil", raiz / "Projeto Perfil.py")
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                dep = mod.obter_proximo_departamento_controlo()
                if dep:
                    print(f"🎯 Próximo departamento detetado automaticamente da sheet CONTROLO: '{dep}' (STATUS e TIMESTAMP vazios)")
                else:
                    dep = "Client Services"
            users_alvo = obter_utilizadores_excel(departamento=dep)

        if not users_alvo:
            print(f"❌ Nenhum utilizador encontrado para os critérios indicados.")
            return

        processar_lista_utilizadores(
            utilizadores=users_alvo,
            sistema_a_remover=args.sistema,
            dry_run=args.dry_run,
            delay_s=args.delay,
        )

    # Executar numa thread isolada para garantir inicialização de WinSta0 limpa
    t = threading.Thread(target=run_worker)
    t.start()
    t.join()


if __name__ == "__main__":
    main()
