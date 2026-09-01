from __future__ import annotations

import os
import re
import sqlite3
from pathlib import Path
from typing import Any


def normalize_environment(environment: str | None = None) -> str:
    env = str(environment or os.getenv("SAP_FI_ENV") or os.getenv("SAP_DEFAULT_ENVIRONMENT") or "PRD").strip().upper()
    return env if env in {"DEV", "QAD", "PRD", "CUA"} else "PRD"


def env_default(environment: str, field_name: str, default: str = "") -> str:
    env = normalize_environment(environment)
    field_key = str(field_name or "").strip().upper()
    if not field_key:
        return default

    for candidate in (f"SAP_{env}_FI_{field_key}", f"SAP_FI_{field_key}"):
        value = os.getenv(candidate, "").strip()
        if value:
            return value
    return default


def env_alias_default(environment: str, aliases: tuple[str, ...], default: str = "") -> str:
    env = normalize_environment(environment)
    for alias in aliases:
        alias_key = str(alias or "").strip().upper()
        if not alias_key:
            continue
        for candidate in (f"SAP_{env}_{alias_key}", f"SAP_{alias_key}"):
            value = os.getenv(candidate, "").strip()
            if value:
                return value
    return default


def payment_method_default(environment: str, branch: str, fallback: str = "") -> str:
    branch_key = str(branch or "").strip().lower()
    aliases_by_branch = {
        "cliente": (
            "FI_FORM_PAGTO_CLIENTE",
            "FI_PAYMENT_METHOD_CLIENTE",
            "FI_FORM_PAGTO",
            "FI_PAYMENT_METHOD",
            "F110_PAYMENT_METHOD",
        ),
        "fornecedor": (
            "FI_FORM_PAGTO_FORNECEDOR",
            "FI_PAYMENT_METHOD_FORNECEDOR",
            "FI_FORM_PAGTO",
            "FI_PAYMENT_METHOD",
            "F110_PAYMENT_METHOD",
        ),
        "razao": (
            "FI_FORM_PAGTO_RAZAO",
            "FI_PAYMENT_METHOD_RAZAO",
            "FI_FORM_PAGTO",
            "FI_PAYMENT_METHOD",
            "F110_PAYMENT_METHOD",
        ),
    }
    return env_alias_default(environment, aliases_by_branch.get(branch_key, ()), fallback)


def _sequence_store_path() -> Path:
    data_dir = Path(__file__).resolve().parents[1] / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir / "fi_reference_sequence.sqlite3"


def _next_reference(prefix: str) -> str:
    safe_prefix = str(prefix or "RFC-TEST").strip().upper() or "RFC-TEST"
    match = re.match(r"^(.*?)-V(\d+)$", safe_prefix)
    base_prefix = match.group(1) if match else safe_prefix
    db_path = _sequence_store_path()
    connection = sqlite3.connect(db_path)
    try:
        connection.execute(
            "CREATE TABLE IF NOT EXISTS reference_sequence (name TEXT PRIMARY KEY, next_value INTEGER NOT NULL)"
        )
        connection.execute("BEGIN IMMEDIATE")
        row = connection.execute(
            "SELECT next_value FROM reference_sequence WHERE name = ?",
            (base_prefix,),
        ).fetchone()
        current = int(row[0]) if row else 1
        next_value = current + 1
        if row:
            connection.execute(
                "UPDATE reference_sequence SET next_value = ? WHERE name = ?",
                (next_value, base_prefix),
            )
        else:
            connection.execute(
                "INSERT INTO reference_sequence (name, next_value) VALUES (?, ?)",
                (base_prefix, next_value),
            )
        connection.commit()
        return f"{base_prefix}-V{current}"
    finally:
        connection.close()


def env_user(environment: str, default: str = "") -> str:
    env = normalize_environment(environment)
    return str(os.getenv(f"SAP_{env}_USER", default) or default).strip()


def build_connection_params(environment: str | None = None) -> dict[str, str]:
    env = normalize_environment(environment)
    required = [
        f"SAP_{env}_USER",
        f"SAP_{env}_PASSWD",
        f"SAP_{env}_ASHOST",
        f"SAP_{env}_SYSNR",
        f"SAP_{env}_CLIENT",
    ]
    missing = [name for name in required if not os.getenv(name, "").strip()]
    if missing:
        raise RuntimeError(f"Variáveis RFC em falta para {env}: {', '.join(missing)}")

    return {
        "user": os.environ[f"SAP_{env}_USER"],
        "passwd": os.environ[f"SAP_{env}_PASSWD"],
        "ashost": os.environ[f"SAP_{env}_ASHOST"],
        "sysnr": os.environ[f"SAP_{env}_SYSNR"],
        "client": os.environ[f"SAP_{env}_CLIENT"],
        "lang": os.getenv(f"SAP_{env}_LANG", "PT").strip() or "PT",
    }


def get_fi_default_context() -> dict[str, Any]:
    def env(name: str, default: str = "") -> str:
        return str(os.getenv(name, default) or default).strip()

    default_user = env("SAP_FI_USERNAME", env_user(normalize_environment(), ""))

    return {
        "common": {
            "company_code": env("SAP_FI_COMPANY_CODE", "2010"),
            "posting_date": env("SAP_FI_POSTING_DATE"),
            "document_date": env("SAP_FI_DOCUMENT_DATE"),
            "currency": env("SAP_FI_CURRENCY", "EUR"),
            "username": default_user,
            "payment_method": env("SAP_FI_FORM_PAGTO", env("SAP_F110_PAYMENT_METHOD")),
            "amount": env("SAP_FI_AMOUNT", "1.00"),
            "header_text": env("SAP_FI_HEADER_TEXT", "RFC-TEST"),
            "item_text": env("SAP_FI_ITEM_TEXT", "RFC-TEST"),
            "reference_prefix": env("SAP_FI_REFERENCE_PREFIX", "RFC-TEST"),
            "tax_code": env("SAP_FI_TAX_CODE"),
            "tax_amount": env("SAP_FI_TAX_AMOUNT", "0.00"),
            "tax_rate": env("SAP_FI_TAX_RATE"),
            "tax_gl_account": env("SAP_FI_TAX_GL_ACCOUNT"),
            "tax_direction": env("SAP_FI_TAX_DIRECTION", "credit"),
        },
        "branches": {
            "cliente": {
                "doc_type": env("SAP_FI_DOC_TYPE_CLIENTE", "DR"),
                "account": env("SAP_FI_CUSTOMER_ACCOUNT", "0010002949"),
                "counterparty": env("SAP_FI_REVENUE_GL_ACCOUNT", "0012990100"),
                "username": default_user,
                "payment_method": env(
                    "SAP_FI_FORM_PAGTO_CLIENTE",
                    env("SAP_F110_PAYMENT_METHOD_CLIENTE", env("SAP_FI_FORM_PAGTO", env("SAP_F110_PAYMENT_METHOD", "Q"))),
                ),
            },
            "fornecedor": {
                "doc_type": env("SAP_FI_DOC_TYPE_FORNECEDOR", "KR"),
                "account": env("SAP_FI_VENDOR_ACCOUNT", "0010000040"),
                "counterparty": env("SAP_FI_EXPENSE_GL_ACCOUNT", "0012010731"),
                "username": default_user,
                "payment_method": env(
                    "SAP_FI_FORM_PAGTO_FORNECEDOR",
                    env("SAP_F110_PAYMENT_METHOD_FORNECEDOR", env("SAP_FI_FORM_PAGTO", env("SAP_F110_PAYMENT_METHOD", "S"))),
                ),
            },
            "razao": {
                "doc_type": env("SAP_FI_DOC_TYPE_RAZAO", "SA"),
                "debit_gl_account": env("SAP_FI_DEBIT_GL_ACCOUNT", "0012010741"),
                "credit_gl_account": env("SAP_FI_CREDIT_GL_ACCOUNT", "0012010741"),
                "username": default_user,
                "payment_method": env(
                    "SAP_FI_FORM_PAGTO_RAZAO",
                    env("SAP_F110_PAYMENT_METHOD_RAZAO", env("SAP_FI_FORM_PAGTO", env("SAP_F110_PAYMENT_METHOD", ""))),
                ),
            },
        },
    }
