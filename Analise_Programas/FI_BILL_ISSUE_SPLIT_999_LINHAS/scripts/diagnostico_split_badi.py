"""Diagnóstico de configuração e implementação da BAdI FI_BILL_ISSUE_SPLIT.

Verifica:
1. Métodos e código-fonte da classe ZCLFI_BILL_ISSUE_SPLIT via RPY_PROGRAM_READ.
2. Comparativo dos tipos de documento RH e ZV na tabela T003.
3. Determinação de conta de compensação de split (chave SPL) na tabela T030.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params, find_project_root, load_project_env


def main() -> int:
    project_root = find_project_root()
    load_project_env(project_root)

    try:
        from pyrfc import Connection  # type: ignore
    except ImportError:
        print("❌ PyRFC não está instalado no ambiente Python atual.", file=sys.stderr)
        return 1

    params = build_connection_params()
    print("=" * 70)
    print(f" Diagnóstico BAdI FI_BILL_ISSUE_SPLIT - Sistema: {params.get('ashost')} (Client {params.get('client')})")
    print("=" * 70)

    conn = Connection(**params)
    try:
        # 1. Verificar classe ZCLFI_BILL_ISSUE_SPLIT
        print("\n[1/3] Verificando métodos da classe ZCLFI_BILL_ISSUE_SPLIT...")
        res_trdir = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="TRDIR",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "NAME"}],
            OPTIONS=[{"TEXT": "NAME LIKE 'ZCLFI_BILL_ISSUE_SPLIT%CM%'"}],
        )

        for row in res_trdir.get("DATA", []):
            inc_name = row["WA"].strip()
            try:
                src = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=inc_name)
                lines = [l.get("LINE", "") for l in src.get("SOURCE_EXTENDED", [])]
                print(f"  • Include: {inc_name}")
                for l in lines:
                    if l.strip() and not l.strip().startswith("*\""):
                        print(f"      {l}")
            except Exception as e:
                print(f"    ⚠️ Erro ao ler {inc_name}: {e}")

        # 2. Verificar tipos de documento RH e ZV na T003
        print("\n[2/3] Verificando tipos de documento RH e ZV na T003...")
        try:
            res_t003 = conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="T003",
                DELIMITER="|",
                FIELDS=[
                    {"FIELDNAME": "BLART"},
                    {"FIELDNAME": "NUMKR"},
                    {"FIELDNAME": "XNETP"},
                    {"FIELDNAME": "STODT"},
                    {"FIELDNAME": "XGSBE"},
                ],
                OPTIONS=[{"TEXT": "BLART = 'RH' OR BLART = 'ZV'"}],
            )
            for row in res_t003.get("DATA", []):
                parts = [p.strip() for p in row["WA"].split("|")]
                print(f"  • Tipo {parts[0]}: Intervalo={parts[1]}, Líquido={parts[2]}, Divisão={parts[4]}")
        except Exception as e:
            print(f"  ⚠️ Não foi possível ler T003 via RFC: {e}")

        # 3. Verificar conta de compensação SPL na T030
        print("\n[3/3] Verificando conta de compensação de split (KTOSL = 'SPL')...")
        res_t030 = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="T030",
            DELIMITER="|",
            FIELDS=[
                {"FIELDNAME": "KTOPL"},
                {"FIELDNAME": "KTOSL"},
                {"FIELDNAME": "KONTH"},
            ],
            OPTIONS=[{"TEXT": "KTOSL = 'SPL'"}],
            ROWCOUNT=5,
        )
        rows_spl = res_t030.get("DATA", [])
        if rows_spl:
            for r in rows_spl:
                parts = [p.strip() for p in r["WA"].split("|")]
                print(f"  • Plano de Contas: {parts[0]} | Chave: {parts[1]} | Conta: {parts[2]}")
        else:
            print("  ⚠️ Nenhuma conta SPL encontrada na tabela T030.")

        print("\n✅ Diagnóstico concluído com sucesso.")
        return 0

    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
