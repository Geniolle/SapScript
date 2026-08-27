from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]
os.environ.setdefault("SAP_SCRIPT_PROJECT_DIR", str(PROJECT_ROOT))

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from sap_rfc.f110_service import OPERATION_PAGAMENTO, run_f110_proposal


BRANCH = "pagamento"

WEB_CONFIG = {
    "branch": BRANCH,
}

# Este processo só executa a PROPOSTA do F110 (Vorlauf). Nunca dispara o pagamento real.
WEB_PARAMS = [
    {
        "name": "environment",
        "label": "Ambiente",
        "type": "select",
        "required": True,
        "default": "DEV",
        "options": [
            {"value": "DEV", "label": "DEV"},
            {"value": "QAD", "label": "QAD"},
            {"value": "PRD", "label": "PRD"},
        ],
    },
    {"name": "company_code", "label": "Empresa", "type": "text", "required": True},
    {"name": "posting_date", "label": "Data de lançamento/base", "type": "date", "required": True},
    {"name": "next_due_date", "label": "Data do próximo vencimento", "type": "date", "required": True},
    {"name": "payment_method", "label": "Forma de pagamento", "type": "text", "required": True},
    {"name": "vendor_account", "label": "Fornecedor", "type": "text", "required": True},
    {
        "name": "document_number",
        "label": "Documento específico a compensar (opcional)",
        "type": "text",
        "required": False,
    },
]


def _error_payload(message: str, *, environment: str, company_code: str, payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "ok": False,
        "status": "ERRO",
        "message": message,
        "branch": BRANCH,
        "environment": environment,
        "company_code": company_code,
        "payload": payload,
    }


def main(
    environment: str = "",
    company_code: str = "",
    posting_date: str = "",
    next_due_date: str = "",
    payment_method: str = "",
    vendor_account: str = "",
    document_number: str = "",
) -> dict[str, Any]:
    payload = {
        "company_code": company_code,
        "posting_date": posting_date,
        "next_due_date": next_due_date,
        "payment_method": payment_method,
        "vendor_account": vendor_account,
        "document_number": document_number,
    }

    try:
        result = run_f110_proposal(
            environment,
            OPERATION_PAGAMENTO,
            company_code=company_code,
            payment_method=payment_method,
            account_number=vendor_account,
            posting_date=posting_date,
            next_due_date=next_due_date,
            document_number=document_number,
        )
        return asdict(result)
    except Exception as exc:
        return _error_payload(str(exc), environment=environment, company_code=company_code, payload=payload)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Executar proposta F110 (pagamento a fornecedor) via RFF110S.")
    parser.add_argument("--environment", default="")
    parser.add_argument("--company-code", dest="company_code", required=True)
    parser.add_argument("--posting-date", dest="posting_date", required=True)
    parser.add_argument("--next-due-date", dest="next_due_date", required=True)
    parser.add_argument("--payment-method", dest="payment_method", required=True)
    parser.add_argument("--vendor-account", dest="vendor_account", required=True)
    parser.add_argument("--document-number", dest="document_number", default="")
    return parser.parse_args()


def _cli() -> int:
    args = _parse_args()
    result = main(**vars(args))
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(_cli())
