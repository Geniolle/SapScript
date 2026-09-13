# -*- coding: utf-8 -*-
"""
B. SYNC_EXECUTAR.py
======================================================================
FASE 2: sincroniza as funções (AGR_USERS, ocorrências diretas) de um
utilizador entre PRD e QAS, através do sistema central CUA (sessão GUI
"SPA"):

  1) Lê por RFC (SAP PRD) quantas/quais funções diretas o utilizador
     tem em produção.
  2) No CUA (SU01, tab "Funções"), localiza e remove TODAS as linhas
     do utilizador cujo SUBSYSTEM seja o ambiente QAS (S4QCLNT100) —
     nunca por posição, sempre lendo a grelha atual antes de cada
     remoção.
  3) No CUA (SU10, mesma sessão), adiciona as funções lidas de PRD ao
     mesmo utilizador com SUBSYSTEM = S4QCLNT100.
  4) Volta a ler a grelha do SU01 (ACTG, SUBSYSTEM=S4QCLNT100) para
     confirmar que a lista final em QAS é exatamente igual à de PRD.

Nota: a ligação RFC direta ao QAS (SAP_QAD_*) não está acessível a
partir desta rede (timeout de rede confirmado) — por isso a leitura/
validação do lado QAS é sempre feita pela grelha do SAP GUI (CUA),
nunca por RFC. O lado PRD é sempre lido por RFC, conforme pedido.
======================================================================
"""

import os
import sys
import time
import importlib.util
from pathlib import Path
from typing import Any, Dict, List

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

SUBSYSTEM_QAS = "S4QCLNT100"
SISTEMA_GUI = "SPA"


def _importar_modulo(caminho: Path, nome_modulo: str):
    spec = importlib.util.spec_from_file_location(nome_modulo, caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# =====================================================================
# 1) LEITURA RFC — SOMENTE PRD
# =====================================================================

def ler_funcoes_diretas_prd(uname: str) -> List[str]:
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
            options=[{"TEXT": f"UNAME = '{uname.strip().upper()}'"}],
            rowcount=0,
        )
    finally:
        conn.close()

    return sorted({a.strip().upper() for _, a, _, _, c in rows if not c.strip()})


# =====================================================================
# 2) REMOÇÃO NO CUA (SU01, tab ACTG) — TODAS AS LINHAS DO SUBSYSTEM QAS
# =====================================================================

def remover_funcoes_qas(crw, session, uname: str, log) -> List[str]:
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

    removidas: List[str] = []
    while True:
        total_linhas = crw.obter_row_count_grid(shell)
        row_alvo = None
        agr_alvo = None
        for row in range(total_linhas):
            sub_grid = crw.obter_valor_celula_grid(shell, row, "SUBSYSTEM")
            if crw.normalizar_coluna(sub_grid) == SUBSYSTEM_QAS:
                row_alvo = row
                agr_alvo = crw.obter_valor_celula_grid(shell, row, "AGR_NAME")
                break

        if row_alvo is None:
            break

        shell.setCurrentCell(row_alvo, "AGR_NAME")
        shell.selectedRows = str(row_alvo)
        shell.pressToolbarButton("DEL_LINE")
        time.sleep(0.25)
        removidas.append(crw.normalizar_coluna(agr_alvo))
        log(f"    removida de QAS: {agr_alvo}")

    if removidas:
        crw.enviar_vkey_debug(session, "wnd[0]", 11, "gravar alterações")
        time.sleep(0.8)
        tipo_sbar, texto_sbar = crw.obter_status_bar(session)
        if tipo_sbar in ("E", "A"):
            raise RuntimeError(f"Erro ao gravar remoção QAS de '{uname}': {texto_sbar}")
        log(f"    status bar após gravar remoção: '{texto_sbar}'")

        # Não confiar apenas na status bar: reabrir e reler a grelha para
        # confirmar que já não restam linhas QAS antes de prosseguir.
        crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/nSU01", "campo de comando")
        crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar SU01")
        time.sleep(0.3)
        crw.setar_texto_debug(session, "wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME", uname, "campo utilizador")
        crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar utilizador")
        time.sleep(0.4)
        crw.pressionar_botao_debug(session, "wnd[0]/tbar[1]/btn[18]", "botão alterar")
        time.sleep(0.3)
        crw.selecionar_tab_debug(session, "wnd[0]/usr/tabsTABSTRIP1/tabpACTG", "tab ACTG")
        time.sleep(0.3)
        shell_verif = crw.obter_grid_roles(session)
        restantes = [
            crw.obter_valor_celula_grid(shell_verif, r, "AGR_NAME")
            for r in range(crw.obter_row_count_grid(shell_verif))
            if crw.normalizar_coluna(crw.obter_valor_celula_grid(shell_verif, r, "SUBSYSTEM")) == SUBSYSTEM_QAS
        ]
        if restantes:
            raise RuntimeError(
                f"Gravação da remoção não persistiu: ainda restam {len(restantes)} linha(s) QAS: {restantes}"
            )
        log("    confirmado: 0 linhas QAS restantes após gravar.")

    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/n", "sair da transação")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saída")
    time.sleep(0.3)

    return removidas


# =====================================================================
# 3) ADIÇÃO NO CUA (SU10, tab ACTG) — MESMAS FUNÇÕES DE PRD
# =====================================================================

def adicionar_funcoes_qas(cad, session, uname: str, funcoes: List[str], log) -> Dict[str, Any]:
    import pandas as pd

    df = pd.DataFrame([
        {
            "ID": str(i + 1),
            "UTILIZADOR": uname,
            "SISTEMA": SUBSYSTEM_QAS,
            "AGR_NAME": agr_name,
            "STATUS": "",
            "MSG": "",
            "TIMESTEMP": "",
        }
        for i, agr_name in enumerate(funcoes)
    ])

    df_proc = cad.atribuir_funcao_usuario(df, session, SISTEMA_GUI, pular_confirmacao=True)
    linha = df_proc.iloc[0]
    status = str(linha.get("STATUS") or "").strip()
    msg = str(linha.get("MSG") or "").strip()
    log(f"    resultado SU10: status='{status}' | msg='{msg}'")
    return {"status": status, "msg": msg}


# =====================================================================
# 4) VALIDAÇÃO FINAL — RELÊ A GRELHA DO SU01 (ACTG, SUBSYSTEM=QAS)
# =====================================================================

def ler_funcoes_qas_via_gui(crw, session, uname: str) -> List[str]:
    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/nSU01", "campo de comando")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar SU01")
    time.sleep(0.3)
    crw.setar_texto_debug(session, "wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME", uname, "campo utilizador")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar utilizador")
    time.sleep(0.4)
    crw.pressionar_botao_debug(session, "wnd[0]/tbar[1]/btn[18]", "botão alterar")
    time.sleep(0.3)
    crw.selecionar_tab_debug(session, "wnd[0]/usr/tabsTABSTRIP1/tabpACTG", "tab ACTG")
    time.sleep(0.3)
    shell = crw.obter_grid_roles(session)

    funcoes = []
    for row in range(crw.obter_row_count_grid(shell)):
        sub_grid = crw.obter_valor_celula_grid(shell, row, "SUBSYSTEM")
        if crw.normalizar_coluna(sub_grid) == SUBSYSTEM_QAS:
            funcoes.append(crw.normalizar_coluna(crw.obter_valor_celula_grid(shell, row, "AGR_NAME")))

    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/n", "sair da transação")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saída")
    time.sleep(0.3)

    return sorted(set(funcoes))


# =====================================================================
# EXECUÇÃO PRINCIPAL
# =====================================================================

def sincronizar_utilizador(uname: str) -> Dict[str, Any]:
    uname = uname.strip().upper()

    print("=" * 75)
    print("  SINCRONIZAR AMBIENTES: PRD -> QAS (via CUA)")
    print("=" * 75)
    print(f"\nUtilizador: {uname}")

    print("\n[1/4] A ler funções diretas em SAP PRD (RFC)...")
    funcoes_prd = ler_funcoes_diretas_prd(uname)
    print(f"       PRD: {len(funcoes_prd)} função(ões) direta(s): {funcoes_prd}")

    crw = _importar_modulo(PROJECT_ROOT / "Processos" / "Funções PFCG" / "CUA_REMOVE_WEB.py", "cua_remove_web")
    cad = _importar_modulo(PROJECT_ROOT / "Processos" / "Funções PFCG" / "CUA_ADICIONAR_WEB.py", "cua_adicionar_web")

    session = crw.conectar_sap(SISTEMA_GUI)
    if not session:
        msg = f"Sessão SAP GUI '{SISTEMA_GUI}' não encontrada."
        print(f"[ABORTADO] {msg}")
        return {"ok": False, "message": msg}

    try:
        from sap_session import apply_window_mode
        apply_window_mode(session)
    except Exception:
        pass

    print(f"\n[2/4] A remover em CUA todas as funções atuais de QAS ({SUBSYSTEM_QAS})...")
    try:
        removidas = remover_funcoes_qas(crw, session, uname, print)
    except Exception as err:
        print(f"[ABORTADO] Falha ao remover funções QAS de '{uname}': {err}")
        return {"ok": False, "message": str(err), "etapa": "remocao"}
    print(f"       {len(removidas)} função(ões) removida(s) de QAS: {removidas}")

    print(f"\n[3/4] A adicionar em CUA as {len(funcoes_prd)} função(ões) de PRD em QAS...")
    if funcoes_prd:
        try:
            resultado_add = adicionar_funcoes_qas(cad, session, uname, funcoes_prd, print)
        except Exception as err:
            print(f"[ABORTADO] Falha ao adicionar funções em QAS para '{uname}': {err}")
            return {"ok": False, "message": str(err), "etapa": "adicao", "removidas": removidas}
        # Nota: a deteção de sucesso/erro pela status bar em
        # CUA_ADICIONAR_WEB.atribuir_funcao_usuario já se mostrou pouco
        # fiável (falsos negativos confirmados). Por isso não abortamos
        # aqui com base em `resultado_add['status']` — o passo [4/4]
        # relê a grelha real do SU01 e é essa a validação que decide
        # sucesso/falha.
        print(f"       (info, não decisivo) status SU10 reportado: {resultado_add}")
    else:
        print("       Nenhuma função em PRD — nada a adicionar.")

    print(f"\n[4/4] A validar o resultado final em QAS (via grelha do SU01)...")
    funcoes_qas_final = ler_funcoes_qas_via_gui(crw, session, uname)
    print(f"       QAS final: {len(funcoes_qas_final)} função(ões): {funcoes_qas_final}")

    igual = funcoes_qas_final == sorted(funcoes_prd)

    print("\n" + "-" * 75)
    print("RESUMO")
    print("-" * 75)
    print(f"Utilizador: {uname}")
    print(f"Funções em PRD: {len(funcoes_prd)}")
    print(f"Funções removidas de QAS: {len(removidas)}")
    print(f"Funções adicionadas em QAS: {len(funcoes_prd)}")
    print(f"Funções em QAS após sincronização: {len(funcoes_qas_final)}")
    print(f"QAS igual a PRD: {'SIM' if igual else 'NÃO'}")
    if not igual:
        print(f"  Só em PRD (faltam em QAS): {sorted(set(funcoes_prd) - set(funcoes_qas_final))}")
        print(f"  Só em QAS (não deviam estar): {sorted(set(funcoes_qas_final) - set(funcoes_prd))}")

    return {
        "ok": igual,
        "uname": uname,
        "funcoes_prd": funcoes_prd,
        "removidas_qas": removidas,
        "funcoes_qas_final": funcoes_qas_final,
        "igual": igual,
    }


if __name__ == "__main__":
    alvo = sys.argv[1] if len(sys.argv) > 1 else "S4244"
    sincronizar_utilizador(alvo)
