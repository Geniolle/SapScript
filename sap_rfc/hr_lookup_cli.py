from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.hr_lookup_service import lookup_hr_data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: dados de RH (PA0002/PA0105) por PERNR, sempre contra PRD."
    )
    parser.add_argument("--pernr", dest="pernr", required=True, help="Numero de colaborador (PERNR).")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = lookup_hr_data(args.pernr)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
