from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_transport_service import search_open_transport_requests


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only para pesquisa de Requests de transporte abertas (RFC)."
    )
    parser.add_argument("--environment", required=True, help="Ambiente alvo (apenas DEV é permitido).")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = search_open_transport_requests(args.environment)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
