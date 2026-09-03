from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_user_roles_service import analyze_user_roles_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: funcoes PFCG atribuidas a um utilizador em SAP PRD."
    )
    parser.add_argument(
        "--user",
        "--username",
        dest="username",
        required=True,
        help="Utilizador SAP (ex.: CLOPES).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = analyze_user_roles_prd(args.username)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
