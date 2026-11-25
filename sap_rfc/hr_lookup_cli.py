from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.hr_lookup_service import lookup_hr_by_name, lookup_hr_data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: dados de RH (PA0002/PA0105) por PERNR ou nome, sempre contra PRD."
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--pernr", dest="pernr", help="Numero de colaborador (PERNR).")
    group.add_argument("--name", dest="name", help="Nome e/ou apelido do colaborador.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = lookup_hr_data(args.pernr) if args.pernr is not None else lookup_hr_by_name(args.name)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
