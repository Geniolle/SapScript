from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_object_roles_service import analyze_object_roles_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: funcoes PFCG que contem um objeto de autorizacao em SAP PRD."
    )
    parser.add_argument(
        "--object",
        "--auth-object",
        dest="auth_object",
        required=True,
        help="Objeto de autorizacao (ex.: S_TCODE).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = analyze_object_roles_prd(args.auth_object)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
