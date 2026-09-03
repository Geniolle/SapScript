from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.user_search_service import search_users_by_name


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: pesquisa utilizadores SAP por nome."
    )
    parser.add_argument("--query", "--name", dest="query", required=True, help="Nome ou parte do nome.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = search_users_by_name(args.query)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
