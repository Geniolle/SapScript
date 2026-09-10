from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_role_auth_objects_service import analyze_pfcg_role_auth_objects_prd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only para análise de objetos de autorização atribuídos a uma função PFCG em SAP PRD."
    )
    parser.add_argument(
        "--role-name",
        "--role",
        dest="role_name",
        required=True,
        help="Nome exato da função/perfil PFCG.",
    )
    return parser.parse_args()


def main() -> int:
    # Sem isto, no Windows o print() usa a codepage padrão da consola
    # (ex.: cp1252) mesmo quando o stdout é capturado por um subprocesso —
    # corrompe descrições com acentos (ex.: "não" -> "n?o") antes de o
    # processo pai decodificar como UTF-8.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = parse_args()
    result = analyze_pfcg_role_auth_objects_prd(args.role_name)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
