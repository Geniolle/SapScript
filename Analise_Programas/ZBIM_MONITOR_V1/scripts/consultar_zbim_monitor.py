"""Consulta em modo leitura das faturas no Monitor ZBIM via RFC.

Permite listar itens pendentes de logística (MM) ou finanças (FI),
verificar status de workflow, motivos de bloqueio e divergências.
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))

from sap_rfc._rfc_common import build_connection_params, find_project_root, load_project_env


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Consulta read-only do monitor ZBIM.")
    parser.add_argument("--modo", choices=["log", "fi"], default="log", help="Modo de consulta: log (MM) ou fi (FI).")
    parser.add_argument("--empresa", default="", help="Filtro de Empresa (BUKRS).")
    parser.add_argument("--fatura", default="", help="Filtro de Documento de Fatura (BELNR).")
    parser.add_argument("--exercicio", default="", help="Filtro de Exercício (GJAHR).")
    parser.add_argument("--limite", type=int, default=20, help="Limite de registos a exibir.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    project_root = find_project_root()
    load_project_env(project_root)
    from pyrfc import Connection

    params = build_connection_params()
    conn = Connection(**params)

    query_table = "ZBIM_BLK_INVOICE" if args.modo == "log" else "ZBIM_BLK_INV_FI"
    options = []
    if args.empresa:
        options.append({"TEXT": f"BUKRS = '{args.empresa}'"})
    if args.fatura:
        prefix = "AND " if options else ""
        options.append({"TEXT": f"{prefix}BELNR = '{args.fatura}'"})
    if args.exercicio:
        prefix = "AND " if options else ""
        options.append({"TEXT": f"{prefix}GJAHR = '{args.exercicio}'"})

    try:
        conn.call("RFC_PING")
        print(f"Lendo tabela {query_table} (modo={args.modo.upper()})...")

        if args.modo == "log":
            fields = [
                {"FIELDNAME": "BUKRS"},
                {"FIELDNAME": "BELNR"},
                {"FIELDNAME": "GJAHR"},
                {"FIELDNAME": "BUZEI"},
                {"FIELDNAME": "EBELN"},
                {"FIELDNAME": "LIFNR"},
                {"FIELDNAME": "SPGRP"},
                {"FIELDNAME": "SPGRM"},
                {"FIELDNAME": "WI_ID"},
                {"FIELDNAME": "WI_STAT"},
                {"FIELDNAME": "WI_CRUSER"},
                {"FIELDNAME": "ZSTATUS"},
                {"FIELDNAME": "ZDIF_EM_FAT"},
                {"FIELDNAME": "ZWRBTR_DIF"},
            ]
        else:
            fields = [
                {"FIELDNAME": "BUKRS"},
                {"FIELDNAME": "BELNR"},
                {"FIELDNAME": "GJAHR"},
                {"FIELDNAME": "BUZEI"},
                {"FIELDNAME": "LIFNR"},
                {"FIELDNAME": "GROSS_AMOUNT"},
                {"FIELDNAME": "WAERS"},
                {"FIELDNAME": "WI_ID"},
                {"FIELDNAME": "WI_STAT"},
                {"FIELDNAME": "WI_CRUSER"},
                {"FIELDNAME": "ZSTATUS"},
            ]

        res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=query_table,
            DELIMITER="|",
            ROWCOUNT=args.limite,
            OPTIONS=options,
            FIELDS=fields,
        )

        rows = res.get("DATA", [])
        print(f"Total encontrado: {len(rows)} registo(s)\n")
        header = [f["FIELDNAME"] for f in fields]
        print(" | ".join(header))
        print("-" * 100)
        for r in rows:
            print(" | ".join([p.strip() for p in r["WA"].split("|")]))

    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
