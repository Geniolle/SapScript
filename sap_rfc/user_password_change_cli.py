from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.user_create_service import change_password_rfc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON de ESCRITA para alterar a password de um utilizador SAP via RFC (BAPI_USER_CHANGE)."
    )
    parser.add_argument("--environment", required=True, help="Ambiente alvo (DEV, QAD ou PRD).")
    parser.add_argument("--username", required=True, help="Utilizador SAP a alterar.")
    parser.add_argument("--password", default="", help="Nova password (default: SAP_PASSE_PASSWD do .env).")
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Confirmação explícita e obrigatória de que esta execução deve escrever em SAP.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.confirm:
        print(
            json.dumps(
                {
                    "ok": False,
                    "status": "ERROR",
                    "error_type": "CONFIRMATION_REQUIRED",
                    "message": "Execução de escrita exige a flag --confirm.",
                },
                ensure_ascii=False,
            )
        )
        return 1

    result = change_password_rfc(args.environment, args.username, args.password)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
