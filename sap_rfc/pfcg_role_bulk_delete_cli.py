from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_role_delete_service import bulk_delete_pfcg_roles_rfc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON de ESCRITA para eliminação em massa de funções PFCG via RFC (apenas DEV)."
    )
    parser.add_argument("--environment", required=True, help="Ambiente alvo (apenas DEV é permitido).")
    parser.add_argument(
        "--role-name",
        dest="role_names",
        action="append",
        required=True,
        help="Nome exato de uma função/perfil PFCG a eliminar. Repita a flag para cada função do lote.",
    )
    parser.add_argument(
        "--transport-mode",
        default="LOCAL",
        help="Modo de transporte: LOCAL, CREATE_REQUEST ou EXISTING_REQUEST.",
    )
    parser.add_argument("--request", default="", help="Número da Request existente (modo EXISTING_REQUEST).")
    parser.add_argument(
        "--request-description",
        default="",
        help="Descrição da nova Request a criar (modo CREATE_REQUEST).",
    )
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Confirmação explícita e obrigatória de que esta execução deve eliminar em SAP.",
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
                    "message": "Execução de eliminação exige a flag --confirm.",
                },
                ensure_ascii=False,
            )
        )
        return 1

    result = bulk_delete_pfcg_roles_rfc(
        args.environment,
        args.role_names,
        args.transport_mode,
        args.request,
        args.request_description,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
