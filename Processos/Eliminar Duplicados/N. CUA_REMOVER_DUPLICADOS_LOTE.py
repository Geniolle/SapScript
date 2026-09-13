# -*- coding: utf-8 -*-
"""
N. CUA_REMOVER_DUPLICADOS_LOTE.py
======================================================================
FASE 2: remoção em lote de TODAS as ocorrências diretas duplicadas em
AGR_USERS identificadas por "L. CUA_DUPLICADOS.py", uma sessão SU01/CUA
por utilizador (sistema central CUA, sessão GUI "SPA"), com validação
RFC (SAP PRD, onde a atribuição é efetiva) antes e depois de cada
utilizador.

Nunca localiza uma linha do grid por posição: cada remoção é confirmada
comparando SUBSYSTEM + AGR_NAME + UPDATE_FROM_DAT + UPDATE_TO_DAT antes
de premir DEL_LINE. Se, para qualquer função, não for possível
identificar de forma inequívoca a linha a remover e a linha a manter,
o lote inteiro pára imediatamente sem tocar nesse utilizador.
======================================================================
"""

import os
import sys
import time
import importlib.util
from pathlib import Path
from typing import Any, Dict, List, Optional

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

SUBSYSTEM = "S4PCLNT100"
SISTEMA_GUI = "SPA"


def _importar_modulo(caminho: Path, nome_modulo: str):
    spec = importlib.util.spec_from_file_location(nome_modulo, caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _data_br_para_aaaammdd(valor: str) -> Optional[str]:
    v = str(valor or "").strip()
    partes = v.split(".")
    if len(partes) != 3:
        return None
    dia, mes, ano = partes
    if not (dia.isdigit() and mes.isdigit() and ano.isdigit()):
        return None
    return f"{ano.zfill(4)}{mes.zfill(2)}{dia.zfill(2)}"


# =====================================================================
# LEITURA RFC (SOMENTE LEITURA) — TODAS AS LINHAS DE UM UTILIZADOR
# =====================================================================

def ler_registos_uname_rfc(uname: str) -> List[Dict[str, Any]]:
    os.environ["SAP_TARGET_ENV"] = "PRD"
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, read_table,
    )
    from pyrfc import Connection

    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    conn = Connection(**params)
    guard = make_read_only_guard(("AGR_USERS",))
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_USERS",
            fields=["UNAME", "AGR_NAME", "FROM_DAT", "TO_DAT", "COL_FLAG"],
            options=[{"TEXT": f"UNAME = '{uname}'"}],
            rowcount=0,
        )
    finally:
        conn.close()

    return [
        {
            "uname": u.strip().upper(),
            "agr_name": a.strip().upper(),
            "from_dat": f.strip(),
            "to_dat": t.strip(),
            "col_flag": c.strip(),
        }
        for u, a, f, t, c in rows
    ]


def diretas_por_agr(registos: List[Dict[str, Any]], agr_name: str) -> List[Dict[str, Any]]:
    return [r for r in registos if r["agr_name"] == agr_name and not r["col_flag"]]


# =====================================================================
# CONSTRUÇÃO DA LISTA DE AÇÕES A PARTIR DA ANÁLISE (L. CUA_DUPLICADOS)
# =====================================================================

def obter_acoes_pendentes(l_dup) -> List[Dict[str, Any]]:
    escopo = l_dup.obter_utilizadores_escopo()
    leitura = l_dup.ler_agr_users_prd(escopo["utilizadores"])
    if not leitura.get("ok"):
        raise RuntimeError(f"Falha na leitura RFC: {leitura.get('erro')}")
    analise = l_dup.analisar_duplicados(leitura["registos"])

    acoes: List[Dict[str, Any]] = []
    for g in analise["grupos"]:
        if g["tipo_grupo"] != "DIRETAS_RESOLVIDAS" or g["ambiguo"]:
            continue
        manter = [r for r, acao, _ in g["classificacoes"] if acao == "MANTER"][0]
        remover_list = [r for r, acao, _ in g["classificacoes"] if acao == "REMOVER"]
        for rem in remover_list:
            acoes.append({
                "uname": g["uname"],
                "agr_name": g["agr_name"],
                "manter_from": manter["from_dat"],
                "manter_to": manter["to_dat"],
                "remover_from": rem["from_dat"],
                "remover_to": rem["to_dat"],
            })
    return acoes


def agrupar_por_uname(acoes: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    por_uname: Dict[str, List[Dict[str, Any]]] = {}
    for a in acoes:
        por_uname.setdefault(a["uname"], []).append(a)
    return dict(sorted(por_uname.items()))


# =====================================================================
# PROCESSAMENTO DE UM UTILIZADOR NA SESSÃO GUI (SU01 -> tab ACTG)
# =====================================================================

def processar_usuario(crw, session, uname: str, acoes_usuario: List[Dict[str, Any]], log) -> List[Dict[str, Any]]:
    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/nSU01", "campo de comando")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar SU01")
    time.sleep(0.3)
    crw.setar_texto_debug(session, "wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME", uname, "campo utilizador")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar utilizador")
    time.sleep(0.4)

    tipo_sbar, texto_sbar = crw.obter_status_bar(session)
    if tipo_sbar in ("E", "A"):
        raise RuntimeError(f"Erro ao abrir utilizador '{uname}': {texto_sbar}")

    crw.pressionar_botao_debug(session, "wnd[0]/tbar[1]/btn[18]", "botão alterar")
    time.sleep(0.3)
    crw.selecionar_tab_debug(session, "wnd[0]/usr/tabsTABSTRIP1/tabpACTG", "tab ACTG")
    time.sleep(0.3)
    shell = crw.obter_grid_roles(session)

    removidos: List[Dict[str, Any]] = []

    for acao in acoes_usuario:
        agr_name = acao["agr_name"]
        candidatos = []
        total_linhas = crw.obter_row_count_grid(shell)
        for row in range(total_linhas):
            agr_grid = crw.obter_valor_celula_grid(shell, row, "AGR_NAME")
            if crw.normalizar_coluna(agr_grid) != agr_name:
                continue
            sub_grid = crw.obter_valor_celula_grid(shell, row, "SUBSYSTEM")
            if crw.normalizar_coluna(sub_grid) != SUBSYSTEM:
                continue
            from_grid = _data_br_para_aaaammdd(crw.obter_valor_celula_grid(shell, row, "UPDATE_FROM_DAT"))
            to_grid = _data_br_para_aaaammdd(crw.obter_valor_celula_grid(shell, row, "UPDATE_TO_DAT"))
            candidatos.append({"row": row, "from_dat": from_grid, "to_dat": to_grid})

        linhas_alvo = [c for c in candidatos if c["from_dat"] == acao["remover_from"] and c["to_dat"] == acao["remover_to"]]
        linhas_manter = [c for c in candidatos if c["from_dat"] == acao["manter_from"] and c["to_dat"] == acao["manter_to"]]

        if len(linhas_alvo) != 1 or len(linhas_manter) != 1:
            raise RuntimeError(
                f"[{uname} / {agr_name}] linha ambígua ou não encontrada no grid — abortado sem gravar. "
                f"candidatos={candidatos} | esperado remover(FROM={acao['remover_from']},TO={acao['remover_to']}) "
                f"manter(FROM={acao['manter_from']},TO={acao['manter_to']})"
            )

        row_alvo = linhas_alvo[0]["row"]
        shell.setCurrentCell(row_alvo, "AGR_NAME")
        shell.selectedRows = str(row_alvo)
        shell.pressToolbarButton("DEL_LINE")
        time.sleep(0.25)
        removidos.append(acao)
        log(f"    removida: {agr_name} | FROM={acao['remover_from']} TO={acao['remover_to']} "
            f"(mantida FROM={acao['manter_from']} TO={acao['manter_to']})")

    crw.enviar_vkey_debug(session, "wnd[0]", 11, "gravar alterações")
    time.sleep(0.6)
    tipo_sbar, texto_sbar = crw.obter_status_bar(session)

    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/n", "sair da transação")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saída")
    time.sleep(0.3)

    if tipo_sbar in ("E", "A"):
        raise RuntimeError(f"Erro ao gravar utilizador '{uname}': {texto_sbar}")

    log(f"    status bar após gravar: '{texto_sbar}'")
    return removidos


# =====================================================================
# EXECUÇÃO PRINCIPAL
# =====================================================================

def executar_lote() -> Dict[str, Any]:
    print("=" * 75)
    print("  FASE 2: REMOÇÃO EM LOTE DOS DUPLICADOS DIRETOS (AGR_USERS)")
    print("=" * 75)

    l_dup = _importar_modulo(FOLDER / "L. CUA_DUPLICADOS.py", "l_cua_duplicados")
    crw = _importar_modulo(PROJECT_ROOT / "Processos" / "Funções PFCG" / "CUA_REMOVE_WEB.py", "cua_remove_web")

    print("\n[PREP] A reconfirmar por RFC a lista atual de duplicados diretos (SAP PRD)...")
    acoes = obter_acoes_pendentes(l_dup)
    por_uname = agrupar_por_uname(acoes)
    total_acoes = len(acoes)
    print(f"       {total_acoes} remoções pendentes em {len(por_uname)} utilizador(es).")

    session = crw.conectar_sap(SISTEMA_GUI)
    if not session:
        print(f"[ABORTADO] Sessão SAP GUI '{SISTEMA_GUI}' não encontrada.")
        return {"ok": False, "message": "sessão GUI não encontrada"}

    try:
        from sap_session import apply_window_mode
        apply_window_mode(session)
    except Exception:
        pass

    auditoria: List[Dict[str, Any]] = []
    utilizadores_ok = 0
    utilizadores_falha = 0
    acoes_confirmadas = 0

    for uname, acoes_usuario in por_uname.items():
        print(f"\n[{uname}] {len(acoes_usuario)} remoção(ões) previstas...")

        antes = ler_registos_uname_rfc(uname)
        estado_ok = True
        for acao in acoes_usuario:
            diretas = diretas_por_agr(antes, acao["agr_name"])
            tem_remover = any(r["from_dat"] == acao["remover_from"] and r["to_dat"] == acao["remover_to"] for r in diretas)
            tem_manter = any(r["from_dat"] == acao["manter_from"] and r["to_dat"] == acao["manter_to"] for r in diretas)
            if not (tem_remover and tem_manter):
                estado_ok = False
                print(f"    [AVISO] estado RFC já não confere para {acao['agr_name']} — a saltar utilizador.")
                break

        if not estado_ok:
            auditoria.append({"uname": uname, "status": "SALTADO", "motivo": "estado RFC divergente antes da execução"})
            continue

        try:
            removidos = processar_usuario(crw, session, uname, acoes_usuario, print)
        except Exception as err:
            print(f"[ABORTADO] Falha inesperada no utilizador '{uname}': {err}")
            auditoria.append({"uname": uname, "status": "ERRO", "motivo": str(err)})
            utilizadores_falha += 1
            print("\n" + "=" * 75)
            print("  LOTE INTERROMPIDO — revisão humana necessária antes de continuar")
            print("=" * 75)
            _imprimir_resumo_final(auditoria, total_acoes, acoes_confirmadas)
            return {"ok": False, "auditoria": auditoria, "interrompido_em": uname, "erro": str(err)}

        depois = ler_registos_uname_rfc(uname)
        falhas_validacao = []
        for acao in acoes_usuario:
            diretas = diretas_por_agr(depois, acao["agr_name"])
            removido_sumiu = not any(r["from_dat"] == acao["remover_from"] and r["to_dat"] == acao["remover_to"] for r in diretas)
            mantido_presente = any(r["from_dat"] == acao["manter_from"] and r["to_dat"] == acao["manter_to"] for r in diretas)
            if not (removido_sumiu and mantido_presente and len(diretas) >= 1):
                falhas_validacao.append(acao["agr_name"])

        if falhas_validacao:
            print(f"[ABORTADO] Validação RFC pós-gravação falhou para: {falhas_validacao}")
            auditoria.append({
                "uname": uname, "status": "ERRO_VALIDACAO",
                "motivo": f"funções com validação RFC divergente: {falhas_validacao}",
                "removidos": removidos,
            })
            utilizadores_falha += 1
            print("\n" + "=" * 75)
            print("  LOTE INTERROMPIDO — revisão humana necessária antes de continuar")
            print("=" * 75)
            _imprimir_resumo_final(auditoria, total_acoes, acoes_confirmadas)
            return {"ok": False, "auditoria": auditoria, "interrompido_em": uname}

        acoes_confirmadas += len(removidos)
        utilizadores_ok += 1
        auditoria.append({"uname": uname, "status": "OK", "removidos": removidos})
        print(f"    [OK] {uname}: {len(removidos)} remoção(ões) confirmada(s) por RFC.")

    print("\n" + "=" * 75)
    print("  LOTE CONCLUÍDO")
    print("=" * 75)
    _imprimir_resumo_final(auditoria, total_acoes, acoes_confirmadas)

    return {
        "ok": utilizadores_falha == 0,
        "auditoria": auditoria,
        "utilizadores_ok": utilizadores_ok,
        "utilizadores_falha": utilizadores_falha,
        "acoes_confirmadas": acoes_confirmadas,
        "total_acoes_previstas": total_acoes,
    }


def _imprimir_resumo_final(auditoria: List[Dict[str, Any]], total_acoes: int, acoes_confirmadas: int) -> None:
    print(f"\nUtilizadores processados: {len(auditoria)}")
    print(f"Remoções previstas: {total_acoes} | Remoções confirmadas por RFC: {acoes_confirmadas}")
    for a in auditoria:
        if a["status"] == "OK":
            print(f"  OK       | {a['uname']} | {len(a['removidos'])} removida(s)")
        elif a["status"] == "SALTADO":
            print(f"  SALTADO  | {a['uname']} | {a['motivo']}")
        else:
            print(f"  {a['status']:<9}| {a['uname']} | {a['motivo']}")


if __name__ == "__main__":
    executar_lote()
