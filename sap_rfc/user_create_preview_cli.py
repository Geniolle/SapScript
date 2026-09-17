from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.user_create_service import preview_user_create


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON read-only para pré-visualização de criação de utilizador SAP via RFC."
    )
    parser.add_argument("--environment", required=True, help="Ambiente alvo (DEV, QAD ou PRD).")
    parser.add_argument("--username", required=True, help="Utilizador SAP a criar.")
    parser.add_argument("--first-name", required=True, help="Nome próprio.")
    parser.add_argument("--last-name", required=True, help="Apelido.")
    parser.add_argument("--email", default="", help="Email (opcional).")
    parser.add_argument("--ustyp", default="A", help="Tipo de utilizador (default: A - Dialog).")
    parser.add_argument("--group", default="", help="Grupo de utilizador / CLASS (default: vazio).")
    parser.add_argument("--valid-from", default="", help="Válido de (AAAAMMDD; default: hoje).")
    parser.add_argument("--valid-to", default="", help="Válido até (AAAAMMDD; default: sem limite).")
    parser.add_argument("--password", default="", help="Senha inicial (default: SAP_PASSE_PASSWD do .env).")
    parser.add_argument(
        "--role",
        dest="roles",
        action="append",
        default=[],
        help="Função PFCG a atribuir (pode ser repetido).",
    )
    parser.add_argument("--department", default="", help="Departamento (BAPIADDR3-DEPARTMENT, opcional).")
    parser.add_argument("--function", default="", help="Função/cargo (BAPIADDR3-FUNCTION, opcional).")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = preview_user_create(
        args.environment,
        args.username,
        args.first_name,
        args.last_name,
        args.email,
        args.ustyp,
        args.group,
        args.valid_from,
        args.valid_to,
        args.password,
        args.roles,
        args.department,
        args.function,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
