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

from sap_rfc.fi_document_service import post_fi_document


BRANCH = "razao"

WEB_CONFIG = {
    "branch": BRANCH,
}

WEB_PARAMS = [
    {
        "name": "data_mode",
        "label": "Dados Default",
        "type": "select",
        "required": True,
        "default": "manual",
        "options": [
            {"value": "default", "label": "Dados Default do .env"},
            {"value": "manual", "label": "Dados Manuais"},
        ],
    },
    {"name": "company_code", "label": "Empresa", "type": "text", "required": True},
    {"name": "posting_date", "label": "Data de lançamento", "type": "date", "required": True},
    {"name": "document_date", "label": "Data do documento", "type": "date", "required": True},
    {
        "name": "currency",
        "label": "Moeda",
        "type": "select",
        "required": True,
        "default": "EUR",
        "options": [
            {"value": "EUR", "label": "EUR"},
            {"value": "USD", "label": "USD"},
            {"value": "BRL", "label": "BRL"},
        ],
    },
    {"name": "header_text", "label": "Texto do cabeçalho", "type": "text", "required": False},
    {"name": "reference", "label": "Referência", "type": "text", "required": False},
    {"name": "username", "label": "Utilizador", "type": "text", "required": False},
    {"name": "debit_gl_account", "label": "Conta a debitar", "type": "text", "required": True},
    {"name": "credit_gl_account", "label": "Conta a creditar", "type": "text", "required": True},
    {"name": "amount", "label": "Valor base", "type": "number", "required": True, "step": "0.01"},
    {"name": "tax_code", "label": "Código de imposto", "type": "text", "required": True},
    {"name": "tax_amount", "label": "Valor do imposto", "type": "number", "required": True, "step": "0.01"},
    {
        "name": "tax_direction",
        "label": "Direção do imposto",
        "type": "select",
        "required": True,
        "default": "credit",
        "options": [
            {"value": "credit", "label": "Crédito"},
            {"value": "debit", "label": "Débito"},
        ],
    },
    {"name": "tax_rate", "label": "Taxa de imposto", "type": "number", "required": False, "step": "0.01"},
    {"name": "tax_gl_account", "label": "Conta de imposto", "type": "text", "required": False},
    {"name": "item_text", "label": "Texto da linha", "type": "text", "required": False},
]


def _error_payload(message: str, *, company_code: str, payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "ok": False,
        "status": "ERRO",
        "message": message,
        "branch": BRANCH,
        "environment": payload.get("environment") or "QAD",
        "company_code": company_code,
        "payload": payload,
    }


def main(
    environment: str = "QAD",
    data_mode: str = "manual",
    company_code: str = "",
    posting_date: str = "",
    document_date: str = "",
    currency: str = "EUR",
    header_text: str = "",
    reference: str = "",
    username: str = "",
    debit_gl_account: str = "",
    credit_gl_account: str = "",
    amount: str = "",
    tax_code: str = "",
    tax_amount: str = "0",
    tax_direction: str = "credit",
    tax_rate: str = "",
    tax_gl_account: str = "",
    item_text: str = "",
) -> dict[str, Any]:
    payload = {
        "environment": environment,
        "data_mode": data_mode,
        "company_code": company_code,
        "posting_date": posting_date,
        "document_date": document_date,
        "currency": currency,
        "header_text": header_text,
        "reference": reference,
        "username": username,
        "debit_gl_account": debit_gl_account,
        "credit_gl_account": credit_gl_account,
        "amount": amount,
        "tax_code": tax_code,
        "tax_amount": tax_amount,
        "tax_direction": tax_direction,
        "tax_rate": tax_rate,
        "tax_gl_account": tax_gl_account,
        "item_text": item_text,
    }

    try:
        result = post_fi_document(environment, BRANCH, payload)
        return asdict(result)
    except Exception as exc:
        return _error_payload(str(exc), company_code=company_code, payload=payload)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Criar Documento FI de Razão em QAD via BAPI.")
    parser.add_argument("--environment", default="QAD")
    parser.add_argument("--data-mode", dest="data_mode", default="manual")
    parser.add_argument("--company-code", dest="company_code", required=True)
    parser.add_argument("--posting-date", dest="posting_date", required=True)
    parser.add_argument("--document-date", dest="document_date", required=True)
    parser.add_argument("--currency", default="EUR")
    parser.add_argument("--header-text", dest="header_text", default="")
    parser.add_argument("--reference", default="")
    parser.add_argument("--username", default="")
    parser.add_argument("--debit-gl-account", dest="debit_gl_account", required=True)
    parser.add_argument("--credit-gl-account", dest="credit_gl_account", required=True)
    parser.add_argument("--amount", required=True)
    parser.add_argument("--tax-code", dest="tax_code", required=True)
    parser.add_argument("--tax-amount", dest="tax_amount", default="0")
    parser.add_argument("--tax-direction", dest="tax_direction", default="credit")
    parser.add_argument("--tax-rate", dest="tax_rate", default="")
    parser.add_argument("--tax-gl-account", dest="tax_gl_account", default="")
    parser.add_argument("--item-text", dest="item_text", default="")
    return parser.parse_args()


def _cli() -> int:
    args = _parse_args()
    result = main(**vars(args))
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(_cli())
