from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_roles_field_service import analyze_pfcg_roles_field_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only: quais funções PFCG (de uma lista) têm um campo de autorização (ex.: ACTVT) em SAP PRD."
    )
    parser.add_argument(
        "--role-names",
        dest="role_names",
        required=True,
        help="Nomes das funções PFCG separados por vírgula.",
    )
    parser.add_argument(
        "--field",
        dest="field",
        required=True,
        help="Nome do campo de autorização a pesquisar (ex.: ACTVT).",
    )
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = parse_args()
    role_names = [name.strip() for name in args.role_names.split(",") if name.strip()]
    result = analyze_pfcg_roles_field_prd(role_names, args.field)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
