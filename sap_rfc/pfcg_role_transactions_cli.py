from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_role_transactions_service import analyze_pfcg_role_transactions_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only para análise de transações atribuídas a uma função PFCG em SAP PRD."
    )
    parser.add_argument(
        "--role-name",
        "--role",
        dest="role_name",
        required=True,
        help="Nome exato da função/perfil PFCG.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = analyze_pfcg_role_transactions_prd(args.role_name)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
