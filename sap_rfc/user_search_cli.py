from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.user_search_service import search_users


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: pesquisa utilizadores SAP por nome, email ou username."
    )
    parser.add_argument("--query", "--name", dest="query", required=True, help="Termo de pesquisa.")
    parser.add_argument("--type", dest="search_type", default="name", choices=["name", "email", "username"], help="Tipo de pesquisa.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = search_users(args.query, search_type=args.search_type)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())

