from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date
from pathlib import Path
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    find_project_root,
    load_project_env,
    make_option_eq,
    make_read_only_guard,
    parse_rfc_table_rows,
    read_table,
)
from sap_rfc.fi_document_service import post_fi_document


TEST_CASE_ID = "A2-FI-FORN-QAD-IRF"
TEST_CASE_NAME = "Criar Documento FI de Fornecedor em QAD e validar WITH_ITEM"
WITH_ITEM_FIELDS = [
    "BUKRS",
    "BELNR",
    "GJAHR",
    "BUZEI",
    "WITHT",
    "WT_WITHCD",
    "WT_QBSHB",
    "WT_QSSHH",
]


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Cria um Documento FI de Fornecedor em QAD e valida se o documento "
            "gera registos em WITH_ITEM."
        )
    )
    parser.add_argument("--environment", default="QAD", help="Ambiente SAP alvo.")
    parser.add_argument("--company-code", default=os.getenv("SAP_QAD_FI_COMPANY_CODE", "2010"))
    parser.add_argument("--posting-date", default=date.today().isoformat())
    parser.add_argument("--document-date", default=date.today().isoformat())
    parser.add_argument("--currency", default=os.getenv("SAP_QAD_FI_CURRENCY", "EUR"))
    parser.add_argument(
        "--header-text",
        default="Teste real FI Fornecedor QAD IRF",
    )
    parser.add_argument(
        "--reference",
        default=f"IRF-{date.today().strftime('%Y%m%d')}",
    )
    parser.add_argument("--username", default=os.getenv("SAP_QAD_USER", ""))
    parser.add_argument("--vendor-account", default=os.getenv("SAP_FI_VENDOR_ACCOUNT", ""))
    parser.add_argument(
        "--expense-gl-account",
        default=os.getenv("SAP_FI_EXPENSE_GL_ACCOUNT", ""),
    )
    parser.add_argument("--amount", default=os.getenv("SAP_FI_TEST_AMOUNT", "100.00"))
    parser.add_argument("--tax-code", default=os.getenv("SAP_FI_TAX_CODE", "IVA"))
    parser.add_argument("--tax-amount", default=os.getenv("SAP_FI_TAX_AMOUNT", "23.00"))
    parser.add_argument("--tax-rate", default=os.getenv("SAP_FI_TAX_RATE", "23"))
    parser.add_argument("--tax-gl-account", default=os.getenv("SAP_FI_TAX_GL_ACCOUNT", ""))
    parser.add_argument("--item-text", default="Documento FI")
    parser.add_argument("--withholding-tax-type", default=os.getenv("SAP_FI_WITHHOLDING_TAX_TYPE", ""))
    parser.add_argument("--withholding-tax-code", default=os.getenv("SAP_FI_WITHHOLDING_TAX_CODE", ""))
    parser.add_argument(
        "--withholding-tax-base-amount",
        default=os.getenv("SAP_FI_WITHHOLDING_TAX_BASE_AMOUNT", ""),
    )
    parser.add_argument(
        "--withholding-tax-amount",
        default=os.getenv("SAP_FI_WITHHOLDING_TAX_AMOUNT", ""),
    )
    parser.add_argument(
        "--assert-with-item",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Falha se não houver linhas em WITH_ITEM para o documento criado.",
    )
    parser.add_argument(
        "--no-read-with-item",
        action="store_true",
        help="Cria o documento, mas não faz a leitura de WITH_ITEM.",
    )
    parser.add_argument(
        "--show-payload",
        action="store_true",
        help="Mostra o payload FI montado antes de chamar a RFC.",
    )
    return parser.parse_args()


def _normalize_document_number(value: str) -> str:
    return str(value or "").strip().zfill(10)


def _map_rows(fields: list[str], rows: list[list[str]]) -> list[dict[str, str]]:
    mapped: list[dict[str, str]] = []
    for row in rows:
        item = {field: row[index] if index < len(row) else "" for index, field in enumerate(fields)}
        mapped.append(item)
    return mapped


def _build_payload(args: argparse.Namespace) -> dict[str, Any]:
    payload = {
        "company_code": str(args.company_code or "").strip(),
        "posting_date": str(args.posting_date or "").strip(),
        "document_date": str(args.document_date or "").strip(),
        "currency": str(args.currency or "EUR").strip(),
        "header_text": str(args.header_text or "").strip(),
        "reference": str(args.reference or "").strip(),
        "username": str(args.username or "").strip(),
        "vendor_account": str(args.vendor_account or "").strip(),
        "expense_gl_account": str(args.expense_gl_account or "").strip(),
        "amount": str(args.amount or "").strip(),
        "tax_code": str(args.tax_code or "").strip(),
        "tax_amount": str(args.tax_amount or "").strip(),
        "tax_rate": str(args.tax_rate or "").strip(),
        "tax_gl_account": str(args.tax_gl_account or "").strip(),
        "item_text": str(args.item_text or "").strip(),
    }

    if str(args.withholding_tax_type or "").strip():
        payload["withholding_tax_type"] = str(args.withholding_tax_type).strip()
    if str(args.withholding_tax_code or "").strip():
        payload["withholding_tax_code"] = str(args.withholding_tax_code).strip()
    if str(args.withholding_tax_base_amount or "").strip():
        payload["withholding_tax_base_amount"] = str(args.withholding_tax_base_amount).strip()
    if str(args.withholding_tax_amount or "").strip():
        payload["withholding_tax_amount"] = str(args.withholding_tax_amount).strip()

    payload["data_mode"] = "manual"
    return payload


def _read_with_item(environment: str, company_code: str, document_number: str, posting_date: str) -> list[dict[str, str]]:
    project_root = find_project_root()
    load_project_env(project_root)
    connection = None
    guard = make_read_only_guard(("WITH_ITEM",))
    params = build_connection_params_for_env(environment)
    try:
        from pyrfc import Connection  # type: ignore

        connection = Connection(**params)  # type: ignore[misc]
        fields = WITH_ITEM_FIELDS
        options = [
            *make_option_eq("BUKRS", str(company_code).strip().upper()),
            *make_option_eq("BELNR", _normalize_document_number(document_number)),
            *make_option_eq("GJAHR", str(posting_date).strip()[:4]),
        ]
        rows = read_table(
            connection,
            guard,
            table_name="WITH_ITEM",
            fields=fields,
            options=options,
            rowcount=0,
        )
        return _map_rows(fields, rows)
    finally:
        if connection is not None:
            try:
                connection.close()
            except Exception:
                pass


def main() -> int:
    args = _parse_args()
    project_root = find_project_root()
    load_project_env(project_root)

    expected_fields = {
        "company_code": args.company_code,
        "posting_date": args.posting_date,
        "document_date": args.document_date,
        "currency": args.currency,
        "vendor_account": args.vendor_account,
        "expense_gl_account": args.expense_gl_account,
        "amount": args.amount,
        "tax_code": args.tax_code,
        "tax_amount": args.tax_amount,
        "tax_rate": args.tax_rate,
        "tax_gl_account": args.tax_gl_account,
        "withholding_tax_type": args.withholding_tax_type,
        "withholding_tax_code": args.withholding_tax_code,
        "withholding_tax_base_amount": args.withholding_tax_base_amount,
        "withholding_tax_amount": args.withholding_tax_amount,
    }

    payload = _build_payload(args)
    print(
        json.dumps(
            {
                "test_case_id": TEST_CASE_ID,
                "test_case_name": TEST_CASE_NAME,
                "environment": args.environment,
                "expected_fields": expected_fields,
            },
            ensure_ascii=False,
            indent=2,
        )
    )

    if args.show_payload:
        print(json.dumps(payload, ensure_ascii=False, indent=2))

    result = post_fi_document(args.environment, "fornecedor", payload)
    print(
        json.dumps(
            {
                "ok": result.ok,
                "status": result.status,
                "message": result.message,
                "company_code": result.company_code,
                "document_number": result.document_number,
            },
            ensure_ascii=False,
            indent=2,
        )
    )

    if not result.ok:
        print("Falha ao criar o documento FI; a leitura de WITH_ITEM não será executada.", file=sys.stderr)
        return 1

    if args.no_read_with_item:
        return 0

    if not result.document_number:
        print("Documento criado sem número; não foi possível validar WITH_ITEM.", file=sys.stderr)
        return 2

    with_item_rows = _read_with_item(args.environment, result.company_code or args.company_code, result.document_number, args.posting_date)
    print(
        json.dumps(
            {
                "with_item_fields": WITH_ITEM_FIELDS,
                "with_item_row_count": len(with_item_rows),
                "with_item_rows": with_item_rows,
            },
            ensure_ascii=False,
            indent=2,
        )
    )

    if args.assert_with_item and not with_item_rows:
        print(
            "WITH_ITEM vazio para o documento criado. Isso indica que a contabilização/IRF ainda não está correta.",
            file=sys.stderr,
        )
        return 3

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
