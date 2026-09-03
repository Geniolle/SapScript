from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_transaction_roles_service import analyze_transaction_roles_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: funcoes PFCG a que uma transacao esta atribuida em SAP PRD."
    )
    parser.add_argument(
        "--tcode",
        "--transaction",
        dest="tcode",
        required=True,
        help="Codigo de transacao (ex.: FB01).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = analyze_transaction_roles_prd(args.tcode)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
