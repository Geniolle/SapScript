# -*- coding: utf-8 -*-
"""
G. PFCG_ORGFIELDS_ANALISE.py
======================================================================
Somente leitura, via RFC: para as funções (AGR_NAME únicos) da sheet
'PFCG_CREATE' do ficheiro Excel oficial, verifica em SAP PRD quais
delas têm, em algum objeto de autorização, os campos organizacionais
PRCTR, WERKS, EKORG ou EKGRP preenchidos (AGR_1251.FIELD), expandindo
funções compostas (Sammelrolle) para as funções-membro.

Mesma lógica de 'E. PFCG_ACTVT_ANALISE.py', mas para estes 4 campos
em vez de ACTVT.

Não altera o SAP. Apenas lê (RFC_READ_TABLE) em PRD.
======================================================================
"""

import sys
from pathlib import Path
from typing import Any, Dict, List

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

CAMINHO_EXCEL_PADRAO = PROJECT_ROOT / "sap_script_uploads" / "S4H_Perfis de autorização.xlsx"
NOME_SHEET = "PFCG_CREATE"
CAMPOS_ORG = ["PRCTR", "WERKS", "EKORG", "EKGRP"]


def obter_roles_pfcg_create(caminho_excel: str = None) -> List[str]:
    from openpyxl import load_workbook

    caminho = Path(caminho_excel) if caminho_excel else CAMINHO_EXCEL_PADRAO
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb[NOME_SHEET]
        linhas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not linhas:
        return []

    cabecalho = linhas[0]
    idx_agr = cabecalho.index("AGR_NAME")

    roles: List[str] = []
    vistos = set()
    for linha in linhas[1:]:
        valor = linha[idx_agr]
        if valor is None:
            continue
        nome = str(valor).strip().upper()
        if nome and nome not in vistos:
            vistos.add(nome)
            roles.append(nome)
    return roles


def analisar_campos_org(roles: List[str], campos: List[str] = None) -> Dict[str, Any]:
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, make_option_in, read_table,
    )
    from pyrfc import Connection

    campos = [c.strip().upper() for c in (campos or CAMPOS_ORG)]

    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(("AGR_DEFINE", "AGR_1251", "AGR_AGRS"))
    conn = Connection(**params)
    try:
        # Expande funções compostas (Sammelrolle) para as funções-membro.
        agrs_rows = read_table(
            conn, guard,
            table_name="AGR_AGRS",
            fields=["AGR_NAME", "CHILD_AGR"],
            options=make_option_in("AGR_NAME", roles),
            rowcount=0,
        )
        membros_por_role: Dict[str, List[str]] = {}
        for agr_name, child_agr in agrs_rows:
            membros_por_role.setdefault(agr_name.strip().upper(), []).append(child_agr.strip().upper())

        todas_as_roles_a_ler = set(roles)
        for membros in membros_por_role.values():
            todas_as_roles_a_ler.update(membros)

        auth_rows = read_table(
            conn, guard,
            table_name="AGR_1251",
            fields=["AGR_NAME", "OBJECT", "FIELD", "LOW", "DELETED"],
            options=make_option_in("AGR_NAME", sorted(todas_as_roles_a_ler)),
            rowcount=0,
        )
    finally:
        conn.close()

    # Para cada role (ou função-membro, se composta), lista os objetos/campos organizacionais.
    campos_por_role: Dict[str, Dict[str, set]] = {}
    for agr_name, obj, campo, low, deleted in auth_rows:
        if str(deleted).strip().upper() == "X":
            continue
        campo_norm = campo.strip().upper()
        if campo_norm not in campos:
            continue
        agr_name = agr_name.strip().upper()
        campos_por_role.setdefault(agr_name, {}).setdefault(campo_norm, set()).add(obj.strip())

    resultado: List[Dict[str, Any]] = []
    for role in roles:
        membros = membros_por_role.get(role, [])
        roles_a_verificar = [role] + membros

        por_campo: Dict[str, List[str]] = {c: [] for c in campos}
        for r in roles_a_verificar:
            info = campos_por_role.get(r, {})
            for campo, objetos in info.items():
                por_campo[campo] = sorted(set(por_campo[campo]) | objetos)

        tem_algum = any(por_campo[c] for c in campos)
        resultado.append({
            "AGR_NAME": role,
            "is_composite": bool(membros),
            "composite_members": membros,
            "tem_algum_campo_org": tem_algum,
            "objetos_por_campo": por_campo,
        })

    return {
        "system": "PRD",
        "client": params["client"],
        "total_roles": len(roles),
        "campos_analisados": campos,
        "resultado": resultado,
    }


def main(caminho_excel: str = None):
    print("=" * 75)
    print("  ANALISE (RFC, SOMENTE LEITURA) — CAMPOS ORGANIZACIONAIS NAS FUNÇÕES DA SHEET PFCG_CREATE")
    print(f"  Campos: {', '.join(CAMPOS_ORG)}")
    print("=" * 75)

    roles = obter_roles_pfcg_create(caminho_excel)
    print(f"\nFunções únicas encontradas em '{NOME_SHEET}': {len(roles)}")

    print("\nA consultar SAP PRD (AGR_1251, AGR_AGRS) via RFC...")
    analise = analisar_campos_org(roles)

    com_campo = [r for r in analise["resultado"] if r["tem_algum_campo_org"]]
    sem_campo = [r for r in analise["resultado"] if not r["tem_algum_campo_org"]]

    print("\n" + "-" * 75)
    print(f"RESUMO (Sistema: {analise['system']} | Client: {analise['client']})")
    print("-" * 75)
    print(f"Total de funções analisadas: {analise['total_roles']}")
    print(f"Com pelo menos um dos campos {CAMPOS_ORG}: {len(com_campo)}")
    print(f"Sem nenhum destes campos: {len(sem_campo)}")

    print(f"\nFunções COM campos organizacionais ({len(com_campo)}):")
    for r in com_campo:
        tag = " [composta]" if r["is_composite"] else ""
        detalhe = ", ".join(
            f"{campo}->[{', '.join(objs)}]"
            for campo, objs in r["objetos_por_campo"].items() if objs
        )
        print(f"  - {r['AGR_NAME']}{tag}: {detalhe}")

    print(f"\nFunções SEM estes campos ({len(sem_campo)}):")
    for r in sem_campo:
        tag = " [composta]" if r["is_composite"] else ""
        print(f"  - {r['AGR_NAME']}{tag}")

    return analise


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else None
    main(caminho)
