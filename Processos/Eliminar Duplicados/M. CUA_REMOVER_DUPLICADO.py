# -*- coding: utf-8 -*-
"""
M. CUA_REMOVER_DUPLICADO.py
======================================================================
Remocao segura de UMA ocorrencia direta duplicada em AGR_USERS, via SU01
na sessao GUI do sistema central CUA (SPA), com validacao RFC antes e
depois (leitura em SAP PRD, onde a atribuicao e efetiva).

Ao contrario de CUA_REMOVE_WEB.py, NAO usa o dialogo de filtro do grid
(que compara SUBSYSTEM pelo valor curto 'S4P' vindo de MAPA_SISTEMA,
mas o grid real usa 'S4PCLNT100'). Em vez disso, percorre todas as
linhas do grid e localiza a linha exata comparando SUBSYSTEM + AGR_NAME
+ UPDATE_FROM_DAT + UPDATE_TO_DAT, nunca assumindo a primeira linha.
======================================================================
"""

import os
import sys
import time
import importlib.util
from pathlib import Path
from typing import Any, Dict, Optional

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _importar_cua_remove_web():
    caminho = PROJECT_ROOT / "Processos" / "Funções PFCG" / "CUA_REMOVE_WEB.py"
    spec = importlib.util.spec_from_file_location("cua_remove_web", caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _data_br_para_aaaammdd(valor: str) -> Optional[str]:
    """Converte 'DD.MM.AAAA' (formato do grid SAP) para 'AAAAMMDD'."""
    v = str(valor or "").strip()
    partes = v.split(".")
    if len(partes) != 3:
        return None
    dia, mes, ano = partes
    if not (dia.isdigit() and mes.isdigit() and ano.isdigit()):
        return None
    return f"{ano.zfill(4)}{mes.zfill(2)}{dia.zfill(2)}"


def _ler_grupo_rfc(uname: str, agr_name: str) -> Dict[str, Any]:
    os.environ["SAP_TARGET_ENV"] = "PRD"
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, read_table, make_option_eq
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
            options=[{"TEXT": f"UNAME = '{uname}'"}, {"TEXT": f"AND AGR_NAME = '{agr_name}'"}],
            rowcount=0,
        )
    finally:
        conn.close()

    registos = [
        {
            "uname": u.strip().upper(),
            "agr_name": a.strip().upper(),
            "from_dat": f.strip(),
            "to_dat": t.strip(),
            "col_flag": c.strip(),
        }
        for u, a, f, t, c in rows
    ]
    return {"ok": True, "registos": registos}


def remover_duplicado_direto(
    *,
    uname: str,
    agr_name: str,
    subsystem: str,
    from_dat_remover: str,
    to_dat_remover: str,
    from_dat_manter: str,
    to_dat_manter: str,
    sistema_gui: str = "SPA",
) -> Dict[str, Any]:
    uname = uname.strip().upper()
    agr_name = agr_name.strip().upper()
    subsystem = subsystem.strip().upper()

    resultado: Dict[str, Any] = {"ok": False, "uname": uname, "agr_name": agr_name}

    # 1) Confirmar novamente por RFC que o grupo continua duplicado (estado ANTES).
    print("[1/6] A confirmar por RFC que o duplicado ainda existe (SAP PRD)...")
    antes = _ler_grupo_rfc(uname, agr_name)
    diretas_antes = [r for r in antes["registos"] if not r["col_flag"]]
    tem_remover = any(r["from_dat"] == from_dat_remover and r["to_dat"] == to_dat_remover for r in diretas_antes)
    tem_manter = any(r["from_dat"] == from_dat_manter and r["to_dat"] == to_dat_manter for r in diretas_antes)
    if not (tem_remover and tem_manter and len(diretas_antes) >= 2):
        resultado["message"] = (
            f"Estado RFC nao confere com o esperado antes da remocao. "
            f"Diretas encontradas: {diretas_antes}"
        )
        print(f"[ABORTADO] {resultado['message']}")
        return resultado

    print(f"      OK: diretas antes = {diretas_antes}")

    # 2) Ligar a sessao GUI do sistema central CUA e navegar ate ao grid.
    print(f"[2/6] A ligar a sessao SAP GUI '{sistema_gui}' e a abrir SU01 para '{uname}'...")
    crw = _importar_cua_remove_web()
    session = crw.conectar_sap(sistema_gui)
    if not session:
        resultado["message"] = f"Sessao SAP GUI '{sistema_gui}' nao encontrada."
        return resultado

    try:
        from sap_session import apply_window_mode
        apply_window_mode(session)
    except Exception:
        pass

    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/nSU01", "campo de comando")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar SU01")
    time.sleep(0.3)
    crw.setar_texto_debug(session, "wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME", uname, "campo utilizador")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar utilizador")
    time.sleep(0.4)

    tipo_sbar, texto_sbar = crw.obter_status_bar(session)
    if tipo_sbar in ("E", "A"):
        resultado["message"] = f"Erro ao abrir utilizador '{uname}': {texto_sbar}"
        return resultado

    crw.pressionar_botao_debug(session, "wnd[0]/tbar[1]/btn[18]", "botao alterar")
    time.sleep(0.3)
    crw.selecionar_tab_debug(session, "wnd[0]/usr/tabsTABSTRIP1/tabpACTG", "tab ACTG")
    time.sleep(0.3)
    shell = crw.obter_grid_roles(session)

    # 3) Localizar a linha exata (SUBSYSTEM + AGR_NAME + FROM + TO), nunca a primeira linha por omissao.
    print("[3/6] A localizar a linha exata no grid (SUBSYSTEM + AGR_NAME + FROM_DAT + TO_DAT)...")
    candidatos = []
    total_linhas = crw.obter_row_count_grid(shell)
    for row in range(total_linhas):
        agr_grid = crw.obter_valor_celula_grid(shell, row, "AGR_NAME")
        if crw.normalizar_coluna(agr_grid) != agr_name:
            continue
        sub_grid = crw.obter_valor_celula_grid(shell, row, "SUBSYSTEM")
        if crw.normalizar_coluna(sub_grid) != subsystem:
            continue
        from_grid = _data_br_para_aaaammdd(crw.obter_valor_celula_grid(shell, row, "UPDATE_FROM_DAT"))
        to_grid = _data_br_para_aaaammdd(crw.obter_valor_celula_grid(shell, row, "UPDATE_TO_DAT"))
        candidatos.append({"row": row, "from_dat": from_grid, "to_dat": to_grid})

    linhas_alvo = [c for c in candidatos if c["from_dat"] == from_dat_remover and c["to_dat"] == to_dat_remover]
    linhas_manter = [c for c in candidatos if c["from_dat"] == from_dat_manter and c["to_dat"] == to_dat_manter]

    print(f"      Candidatos SUBSYSTEM+AGR_NAME no grid: {candidatos}")

    if len(linhas_alvo) != 1 or len(linhas_manter) != 1:
        resultado["message"] = (
            "Nao foi possivel identificar de forma inequivoca a linha exata no grid "
            f"(alvo={linhas_alvo}, manter={linhas_manter}). Abortado sem alteracoes."
        )
        print(f"[ABORTADO] {resultado['message']}")
        crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/n", "sair sem gravar")
        crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saida")
        return resultado

    linha_alvo = linhas_alvo[0]["row"]

    # 4) Confirmacao final imediatamente antes de remover.
    print(f"[4/6] Linha exata confirmada: row={linha_alvo}. A remover e a gravar...")
    shell.setCurrentCell(linha_alvo, "AGR_NAME")
    shell.selectedRows = str(linha_alvo)
    shell.pressToolbarButton("DEL_LINE")
    time.sleep(0.3)

    crw.enviar_vkey_debug(session, "wnd[0]", 11, "gravar alteracao")
    time.sleep(0.6)

    tipo_sbar, texto_sbar = crw.obter_status_bar(session)
    print(f"      STATUS BAR apos gravar: tipo='{tipo_sbar}' | texto='{texto_sbar}'")

    crw.setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/n", "sair da transacao")
    crw.enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saida")

    # 5) Validar novamente por RFC (estado DEPOIS).
    print("[5/6] A validar novamente por RFC (SAP PRD)...")
    depois = _ler_grupo_rfc(uname, agr_name)
    diretas_depois = [r for r in depois["registos"] if not r["col_flag"]]
    print(f"      Diretas depois = {diretas_depois}")

    removido_sumiu = not any(
        r["from_dat"] == from_dat_remover and r["to_dat"] == to_dat_remover for r in diretas_depois
    )
    mantido_presente = any(
        r["from_dat"] == from_dat_manter and r["to_dat"] == to_dat_manter for r in diretas_depois
    )
    funcao_valida_restante = len(depois["registos"]) >= 1

    sucesso = removido_sumiu and mantido_presente and funcao_valida_restante

    resultado.update({
        "ok": sucesso,
        "sap_status_tipo": tipo_sbar,
        "sap_status_texto": texto_sbar,
        "diretas_antes": diretas_antes,
        "diretas_depois": diretas_depois,
        "removido_confirmado": removido_sumiu,
        "mantido_confirmado": mantido_presente,
        "funcao_valida_restante": funcao_valida_restante,
    })
    print(f"[6/6] Resultado final: {'SUCESSO' if sucesso else 'FALHA/REVISAO NECESSARIA'}")
    return resultado


def imprimir_resultado_compacto(r: Dict[str, Any]) -> None:
    def sim_nao(valor: bool) -> str:
        return "SIM" if valor else "NAO"

    print("\nTESTE REAL")
    print(f"User: {r.get('uname')}")
    print(f"Role: {r.get('agr_name')}")
    manter = r.get("diretas_depois", [{}])[0] if r.get("diretas_depois") else {}
    print(f"Mantido: FROM={manter.get('from_dat', '?')} | TO={manter.get('to_dat', '?')}")
    removido = next(
        (x for x in r.get("diretas_antes", []) if not any(
            x["from_dat"] == y["from_dat"] and x["to_dat"] == y["to_dat"] for y in r.get("diretas_depois", [])
        )),
        None,
    )
    if removido:
        print(f"Removido: FROM={removido['from_dat']} | TO={removido['to_dat']}")
    print(f"Validação RFC: {'OK' if r.get('ok') else 'FALHOU'}")
    print(f"Função válida restante: {sim_nao(r.get('funcao_valida_restante', False))}")
    if not r.get("ok"):
        print(f"Mensagem: {r.get('message') or r.get('sap_status_texto')}")


if __name__ == "__main__":
    r = remover_duplicado_direto(
        uname="S105",
        agr_name="ZORG_BP_FLVN01_LOGISTICS_VENDO",
        subsystem="S4PCLNT100",
        from_dat_remover="20260709",
        to_dat_remover="99991231",
        from_dat_manter="20260912",
        to_dat_manter="99991231",
    )
    imprimir_resultado_compacto(r)
