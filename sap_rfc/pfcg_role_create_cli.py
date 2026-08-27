from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.pfcg_role_create_service import create_pfcg_role_rfc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON de ESCRITA para criação individual de função PFCG via RFC (apenas DEV)."
    )
    parser.add_argument("--environment", required=True, help="Ambiente alvo (apenas DEV é permitido).")
    parser.add_argument("--role-name", required=True, help="Nome exato da função/perfil PFCG a criar.")
    parser.add_argument("--description", required=True, help="Descrição do Perfil de Autorização.")
    parser.add_argument(
        "--tcode",
        dest="tcodes",
        action="append",
        default=[],
        help="Transação a incluir (pode ser repetido).",
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

    result = create_pfcg_role_rfc(
        args.environment,
        args.role_name,
        args.description,
        args.tcodes,
        args.transport_mode,
        args.request,
        args.request_description,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
