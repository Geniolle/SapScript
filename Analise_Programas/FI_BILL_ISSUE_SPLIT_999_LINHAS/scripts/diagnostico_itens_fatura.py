"""Diagnóstico de volume de itens e fluxo contábil de fatura SD (VBRK / VBRP / VBFA).

Permite verificar se uma fatura atingiu o limite de 999 itens e analisar se
o split gerou múltiplos documentos contábeis ou se o documento contábil falhou.

Uso:
    python diagnostico_itens_fatura.py --fatura 2541036794
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params, find_project_root, load_project_env


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Diagnóstico de itens e fluxo de fatura SD.")
    parser.add_argument("--fatura", required=True, help="Número do documento de faturamento (ex.: 2541036794)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    vbeln = args.fatura.strip().zfill(10)

    project_root = find_project_root()
    load_project_env(project_root)

    try:
        from pyrfc import Connection  # type: ignore
    except ImportError:
        print("❌ PyRFC não está instalado no ambiente Python atual.", file=sys.stderr)
        return 1

    params = build_connection_params()
    print("=" * 70)
    print(f" Análise da Fatura {vbeln} no SAP PRD")
    print("=" * 70)

    conn = Connection(**params)
    try:
        # 1. Cabeçalho da Fatura (VBRK)
        res_vbrk = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="VBRK",
            DELIMITER="|",
            FIELDS=[
                {"FIELDNAME": "VBELN"},
                {"FIELDNAME": "FKART"},
                {"FIELDNAME": "BUKRS"},
                {"FIELDNAME": "FKDAT"},
                {"FIELDNAME": "NETWR"},
                {"FIELDNAME": "WAERK"},
                {"FIELDNAME": "RPLNR"},
                {"FIELDNAME": "FKSTO"},
            ],
            OPTIONS=[{"TEXT": f"VBELN = '{vbeln}'"}],
            ROWCOUNT=1,
        )
        data_vbrk = res_vbrk.get("DATA", [])
        if not data_vbrk:
            print(f"❌ Fatura {vbeln} não encontrada na tabela VBRK.")
            return 2

        parts = [p.strip() for p in data_vbrk[0]["WA"].split("|")]
        print(f"\n[1/3] Dados do Cabeçalho (VBRK):")
        print(f"  • Tipo de Fatura (FKART) : {parts[1]}")
        print(f"  • Empresa (BUKRS)         : {parts[2]}")
        print(f"  • Data Faturação (FKDAT) : {parts[3]}")
        print(f"  • Valor Líquido (NETWR)  : {parts[4]} {parts[5]}")
        print(f"  • Estornada (FKSTO)       : {'SIM' if parts[7] == 'X' else 'NÃO'}")

        # 2. Contagem de Itens (VBRP)
        res_vbrp = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="VBRP",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "POSNR"}],
            OPTIONS=[{"TEXT": f"VBELN = '{vbeln}'"}],
            ROWCOUNT=0,
        )
        item_count = len(res_vbrp.get("DATA", []))
        print(f"\n[2/3] Total de Itens de Fatura (VBRP):")
        print(f"  • Total de posições: {item_count}")
        if item_count >= 999:
            print(f"  ⚠️ ATENÇÃO: Esta fatura tem {item_count} itens (>= 999 posições).")
            print("     Necessita obrigatoriamente do mecanismo de split contábil para evitar erro F5 727.")
        else:
            print(f"  ℹ️ Fatura abaixo do limite de 999 itens ({item_count} posições).")

        # 3. Fluxo de Documentos Subsequentes (VBFA)
        res_vbfa = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="VBFA",
            DELIMITER="|",
            FIELDS=[
                {"FIELDNAME": "VBELV"},
                {"FIELDNAME": "VBELN"},
                {"FIELDNAME": "VBTYP_N"},
            ],
            OPTIONS=[
                {"TEXT": f"VBELV = '{vbeln}'"},
            ],
            ROWCOUNT=20,
        )
        data_vbfa = res_vbfa.get("DATA", [])
        print(f"\n[3/3] Fluxo de Documentos Subsequentes (VBFA):")
        fi_docs = []
        if data_vbfa:
            for r in data_vbfa:
                p = [x.strip() for x in r["WA"].split("|")]
                print(f"  • Sucessor: {p[1]} (Categoria VBTYP_N: '{p[2]}')")
                if p[2] == "R":  # Categoria R = Documento Contábil FI
                    fi_docs.append(p[1])
        else:
            print("  ⚠️ Nenhum documento subsequente registrado no fluxo VBFA.")

        if len(fi_docs) > 1:
            print(f"\n✅ Split confirmado! A fatura gerou {len(fi_docs)} documentos contábeis FI: {', '.join(fi_docs)}")
        elif len(fi_docs) == 1:
            print(f"\nℹ️ Gerado 1 documento contábil único: {fi_docs[0]}")
        else:
            print("\n❌ Nenhum documento contábil FI gerado (possível bloqueio por erro F5 727).")

        return 0

    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
