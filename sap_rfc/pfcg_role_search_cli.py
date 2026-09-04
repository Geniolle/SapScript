from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_role_search_service import search_pfcg_roles


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only para pesquisa de funções/perfis PFCG por padrão de nome (curinga '*')."
    )
    parser.add_argument(
        "--pattern",
        dest="pattern",
        required=True,
        help="Padrão de pesquisa (ex.: Z*EQUIPA*).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = search_pfcg_roles(args.pattern)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
