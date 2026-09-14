# -*- coding: utf-8 -*-

###################################################################################
# M. PFCG_ADD_FUNCAO_COMPOSTA.py
# PFCG - Atribuir a função individual de uma transação à função composta de um
# utilizador, via RFC (sem GUI scripting).
#
# Fluxo (validado manualmente em 2026-09-14 para KS03 / S965 em PRD):
#   1. Sheet "Proposta" do Excel: localizar o bloco de função individual (Z_...)
#      que contém a transação indicada (coluna A do bloco).
#   2. Sheet "Proposta Ativa" do Excel: localizar a "Composite Role" do
#      utilizador indicado (coluna "Usuário").
#   3. RFC (somente leitura, AGR_AGRS): se a função individual já for membro da
#      composta, não faz nada (skip).
#   4. RFC PRGN_RFC_ADD_AGRS_TO_COLL_AGR: adiciona a função individual como
#      membro da composta.
#   5. RFC PRGN_GEN_PROFILES_FOR_ROLES (IV_USERCOMPARE='X'): gera o perfil da
#      composta e propaga a autorização aos utilizadores já atribuídos a ela.
#   6. RFC (ligação independente, somente leitura, AGR_AGRS): confirma que a
#      função individual passou a ser membro da composta.
#   7. Excel, sheet "PFCG_COMPOSTA": regista uma nova linha (TEXT da composta
#      obtido via RFC, tabela AGR_TEXTS).
###################################################################################

import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

CAMINHO_EXCEL_PADRAO = PROJECT_ROOT / "sap_script_uploads" / "S4H_Perfis de autorização.xlsx"
NOME_SHEET_PROPOSTA = "Proposta"
NOME_SHEET_PROPOSTA_ATIVA = "Proposta Ativa"
NOME_SHEET_PFCG_COMPOSTA = "PFCG_COMPOSTA"

ALLOWED_TABLES = ("AGR_AGRS", "AGR_TEXTS", "AGR_DEFINE")
ALLOWED_FUNCTIONS = (
    "RFC_READ_TABLE",
    "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
    "PRGN_GEN_PROFILES_FOR_ROLES",
)


def localizar_funcao_individual_por_tcode(tcode: str, caminho_excel: Optional[str] = None) -> Optional[str]:
    """Sheet 'Proposta': encontra a função individual (Z_...) cujo bloco contém a tcode."""
    import openpyxl

    tcode = tcode.strip().upper()
    caminho = Path(caminho_excel) if caminho_excel else CAMINHO_EXCEL_PADRAO
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb[NOME_SHEET_PROPOSTA]

    role_rows = [r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=2).value]
    role_rows.append(ws.max_row + 1)

    for i in range(len(role_rows) - 1):
        inicio, fim = role_rows[i], role_rows[i + 1]
        role_name = ws.cell(row=inicio, column=1).value
        tcodes_bloco = {
            str(ws.cell(row=r, column=1).value).strip().upper()
            for r in range(inicio + 1, fim)
            if ws.cell(row=r, column=1).value
        }
        if tcode in tcodes_bloco:
            return str(role_name).strip().upper()
    return None


def localizar_composta_do_utilizador(usuario: str, caminho_excel: Optional[str] = None) -> Optional[str]:
    """Sheet 'Proposta Ativa': encontra a 'Composite Role' do utilizador indicado."""
    import openpyxl

    usuario = usuario.strip().upper()
    caminho = Path(caminho_excel) if caminho_excel else CAMINHO_EXCEL_PADRAO
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb[NOME_SHEET_PROPOSTA_ATIVA]

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col_usuario = next((i + 1 for i, h in enumerate(header) if str(h or "").strip().upper() == "USUÁRIO"), 1)
    col_composite = next((i + 1 for i, h in enumerate(header) if "COMPOSITE" in str(h or "").strip().upper()), None)
    if not col_composite:
        raise RuntimeError(f"Coluna 'Composite Role' não encontrada na sheet '{NOME_SHEET_PROPOSTA_ATIVA}'.")

    for r in range(2, ws.max_row + 1):
        valor_usuario = ws.cell(row=r, column=col_usuario).value
        if str(valor_usuario or "").strip().upper() == usuario:
            composta = ws.cell(row=r, column=col_composite).value
            return str(composta).strip().upper() if composta else None
    return None


def _conn_e_guard(ambiente: str, allow_write: bool):
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, make_write_guard,
    )
    from pyrfc import Connection

    load_project_env(find_project_root())
    params = build_connection_params_for(ambiente)
    guard = (
        make_write_guard(ALLOWED_FUNCTIONS, ALLOWED_TABLES)
        if allow_write
        else make_read_only_guard(ALLOWED_TABLES)
    )
    return Connection(**params), guard


def obter_texto_role(ambiente: str, role_name: str) -> str:
    from sap_rfc._rfc_common import make_option_eq, read_table

    conn, guard = _conn_e_guard(ambiente, allow_write=False)
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_TEXTS",
            fields=["AGR_NAME", "TEXT", "SPRAS"],
            options=make_option_eq("AGR_NAME", role_name),
            rowcount=0,
        )
    finally:
        conn.close()
    for _agr, text, spras in rows:
        if spras.strip().upper() == "P":
            return text.strip()
    return rows[0][1].strip() if rows else role_name


def verificar_membro_composta(ambiente: str, composta: str, funcao_individual: str) -> bool:
    from sap_rfc._rfc_common import fetch_composite_members

    conn, guard = _conn_e_guard(ambiente, allow_write=False)
    try:
        membros = fetch_composite_members(conn, guard, composta)
    finally:
        conn.close()
    return funcao_individual.strip().upper() in membros


def adicionar_funcao_a_composta(ambiente: str, composta: str, funcao_individual: str, texto_funcao: str) -> Dict[str, Any]:
    conn, guard = _conn_e_guard(ambiente, allow_write=True)
    try:
        guard.assert_function_allowed("PRGN_RFC_ADD_AGRS_TO_COLL_AGR")
        resultado = conn.call(
            "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
            ACTIVITY_GROUP=composta,
            ACTIVITY_GROUPS=[{"AGR_NAME": funcao_individual, "TEXT": texto_funcao}],
            NO_DIALOG="X",
        )
        mensagens = resultado.get("RETURN", []) or []
        erros = [m for m in mensagens if str(m.get("TYPE", "")).upper() in ("E", "A")]
        if erros:
            return {"ok": False, "message": "; ".join(m.get("MESSAGE", "") for m in erros)}

        guard.assert_function_allowed("PRGN_GEN_PROFILES_FOR_ROLES")
        conn.call(
            "PRGN_GEN_PROFILES_FOR_ROLES",
            IT_ROLES=[{"AGR_NAME": composta}],
            IV_USERCOMPARE="X",
        )
        return {"ok": True, "message": "Função individual adicionada à composta e perfil gerado (user compare incluído)."}
    finally:
        conn.close()


def registar_pfcg_composta(
    composta: str,
    texto_composta: str,
    funcao_individual: str,
    ambiente: str,
    caminho_excel: Optional[str] = None,
) -> None:
    from openpyxl import load_workbook

    caminho = Path(caminho_excel) if caminho_excel else CAMINHO_EXCEL_PADRAO
    wb = load_workbook(caminho)
    ws = wb[NOME_SHEET_PFCG_COMPOSTA]

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col_id = next((i + 1 for i, h in enumerate(header) if str(h or "").strip().upper() == "ID"), 1)
    novo_id = max(
        (ws.cell(row=r, column=col_id).value or 0 for r in range(2, ws.max_row + 1)),
        default=0,
    ) + 1

    ws.append([
        novo_id,
        composta,
        texto_composta,
        funcao_individual,
        "Criado",
        f"Atribuído em SAP {ambiente} (via RFC PRGN_RFC_ADD_AGRS_TO_COLL_AGR)",
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Validado",
    ])
    wb.save(caminho)


def executar_individual(ambiente: str, tcode: str, usuario: str, caminho_excel: Optional[str] = None) -> Dict[str, Any]:
    tcode = tcode.strip().upper()
    usuario = usuario.strip().upper()

    funcao_individual = localizar_funcao_individual_por_tcode(tcode, caminho_excel)
    if not funcao_individual:
        return {"ok": False, "skip": False, "message": f"Nenhuma função individual encontrada para a transação '{tcode}' na sheet '{NOME_SHEET_PROPOSTA}'."}
    print(f"Transação {tcode} -> função individual: {funcao_individual}")

    composta = localizar_composta_do_utilizador(usuario, caminho_excel)
    if not composta:
        return {"ok": False, "skip": False, "message": f"Nenhuma Composite Role encontrada para o utilizador '{usuario}' na sheet '{NOME_SHEET_PROPOSTA_ATIVA}'."}
    print(f"Utilizador {usuario} -> função composta: {composta}")

    print(f"A verificar via RFC (somente leitura) se '{funcao_individual}' já é membro de '{composta}'...")
    if verificar_membro_composta(ambiente, composta, funcao_individual):
        return {
            "ok": True, "skip": True,
            "tcode": tcode, "usuario": usuario, "funcao_individual": funcao_individual, "composta": composta,
            "message": f"'{funcao_individual}' já é membro de '{composta}' — nada a fazer.",
        }

    texto_funcao = obter_texto_role(ambiente, funcao_individual)
    texto_composta = obter_texto_role(ambiente, composta)

    print(f"A adicionar '{funcao_individual}' à composta '{composta}' via RFC ({ambiente})...")
    r = adicionar_funcao_a_composta(ambiente, composta, funcao_individual, texto_funcao)
    if not r["ok"]:
        return {"ok": False, "skip": False, "tcode": tcode, "usuario": usuario, "funcao_individual": funcao_individual, "composta": composta, "message": r["message"]}

    print("A verificar via RFC (ligação independente, somente leitura) o resultado...")
    if not verificar_membro_composta(ambiente, composta, funcao_individual):
        return {
            "ok": False, "skip": False,
            "tcode": tcode, "usuario": usuario, "funcao_individual": funcao_individual, "composta": composta,
            "message": "RFC de escrita não reportou erro, mas a verificação independente não confirmou a nova associação.",
        }

    registar_pfcg_composta(composta, texto_composta, funcao_individual, ambiente, caminho_excel)

    return {
        "ok": True, "skip": False,
        "tcode": tcode, "usuario": usuario, "funcao_individual": funcao_individual, "composta": composta,
        "message": f"'{funcao_individual}' adicionada a '{composta}' em {ambiente}, verificado via RFC e registado no Excel.",
    }


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"], required=True)
    parser.add_argument("--tcode", required=True, help="Transação (ex.: KS03).")
    parser.add_argument("--usuario", required=True, help="Utilizador SAP (ex.: S965).")
    args = parser.parse_args()

    resultado = executar_individual(args.ambiente, args.tcode, args.usuario)

    print("\n" + "=" * 75)
    print("RESULTADO FINAL")
    print("=" * 75)
    print(resultado)
