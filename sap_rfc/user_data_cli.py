from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.user_data_service import analyze_user_data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: dados mestre / pessoais de um utilizador SAP."
    )
    parser.add_argument("--user", "--username", dest="username", required=True, help="Utilizador SAP.")
    parser.add_argument("--kind", dest="kind", required=True, choices=["master", "personal"], help="Tipo de dados.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = analyze_user_data(args.username, args.kind)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
