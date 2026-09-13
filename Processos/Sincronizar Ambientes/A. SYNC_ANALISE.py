# -*- coding: utf-8 -*-
"""
A. SYNC_ANALISE.py
======================================================================
FASE 1 (somente leitura, via RFC): compara as funções (AGR_USERS,
ocorrências diretas — COL_FLAG vazio) que um utilizador tem no
ambiente produtivo (PRD) com as que tem atualmente no ambiente de
qualidade (QAS / SAP_QAD_*), e mostra o que seria removido em QAS e o
que seria adicionado para ficar igual ao PRD.

Não altera o SAP. Apenas lê (RFC_READ_TABLE) em PRD e em QAS.
======================================================================
"""

import os
import sys
from pathlib import Path
from typing import Any, Dict, List

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

SUBSYSTEM_QAS = "S4QCLNT100"


def ler_funcoes_diretas(uname: str, ambiente: str) -> List[str]:
    """Lê AGR_USERS de `ambiente` (PRD ou QAD) e devolve as AGR_NAME diretas (COL_FLAG vazio)."""
    os.environ["SAP_TARGET_ENV"] = ambiente
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, read_table,
    )
    from pyrfc import Connection

    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for(ambiente)
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

    diretas = sorted({
        a.strip().upper()
        for _, a, _, _, c in rows
        if not c.strip()
    })
    return diretas


def comparar(uname: str) -> Dict[str, Any]:
    uname = uname.strip().upper()

    print("=" * 75)
    print("  FASE 1: ANALISE (RFC, SOMENTE LEITURA) — SINCRONIZAR PRD -> QAS")
    print("=" * 75)
    print(f"\nUtilizador: {uname}")

    print("\n[1/2] A ler funções diretas em SAP PRD...")
    prd = ler_funcoes_diretas(uname, "PRD")
    print(f"       PRD: {len(prd)} função(ões) direta(s).")

    print("[2/2] A ler funções diretas em SAP QAS (SAP_QAD_*)...")
    qas = ler_funcoes_diretas(uname, "QAD")
    print(f"       QAS: {len(qas)} função(ões) direta(s).")

    a_remover = sorted(set(qas) - set(prd))
    a_manter = sorted(set(qas) & set(prd))
    a_adicionar = sorted(set(prd) - set(qas))

    print("\n" + "-" * 75)
    print("RESUMO")
    print("-" * 75)
    print(f"Funções em PRD: {len(prd)}")
    print(f"Funções em QAS (atuais): {len(qas)}")
    print(f"Já iguais (nada a fazer): {len(a_manter)}")
    print(f"A remover de QAS (não existem em PRD): {len(a_remover)}")
    for r in a_remover:
        print(f"  - {r}")
    print(f"A adicionar em QAS (existem em PRD, faltam em QAS): {len(a_adicionar)}")
    for a in a_adicionar:
        print(f"  - {a}")

    return {
        "uname": uname,
        "prd": prd,
        "qas": qas,
        "a_manter": a_manter,
        "a_remover": a_remover,
        "a_adicionar": a_adicionar,
    }


if __name__ == "__main__":
    alvo = sys.argv[1] if len(sys.argv) > 1 else "S4244"
    comparar(alvo)
