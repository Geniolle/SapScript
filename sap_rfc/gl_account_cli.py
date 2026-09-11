from __future__ import annotations

import argparse
import json

from sap_rfc.gl_account_service import create_company_account_by_model


def main() -> int:
    parser = argparse.ArgumentParser(description="Criar/estender Conta Razão por modelo via RFC.")
    parser.add_argument("--environment", required=True, choices=("DEV", "QAD", "PRD"))
    parser.add_argument("--account", required=True)
    parser.add_argument("--company", required=True, help="Empresa de destino")
    parser.add_argument("--model-company", required=True)
    parser.add_argument("--alternative-account", default="")
    parser.add_argument("--confirm", action="store_true", help="Grava; sem esta opção executa apenas TESTMODE")
    args = parser.parse_args()
    result = create_company_account_by_model(
        environment=args.environment, account=args.account, target_company=args.company,
        model_company=args.model_company, alternative_account=args.alternative_account,
        test_only=not args.confirm,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
