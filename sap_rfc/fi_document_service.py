from __future__ import annotations

import os
import sqlite3
import re
from datetime import date
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None


@dataclass
class FiDocumentResult:
    ok: bool
    status: str
    message: str
    branch: str
    company_code: str = ""
    document_number: str = ""
    check_return: list[dict[str, Any]] = field(default_factory=list)
    post_return: list[dict[str, Any]] = field(default_factory=list)
    commit_return: list[dict[str, Any]] = field(default_factory=list)
    payload: dict[str, Any] = field(default_factory=dict)


def _normalize_environment(environment: str | None = None) -> str:
    env = str(environment or os.getenv("SAP_FI_ENV") or os.getenv("SAP_DEFAULT_ENVIRONMENT") or "PRD").strip().upper()
    return env if env in {"DEV", "QAD", "PRD", "CUA"} else "PRD"


def build_connection_params(environment: str | None = None) -> dict[str, str]:
    env = _normalize_environment(environment)
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


def _require_pyrfc() -> None:
    if Connection is None:
        raise RuntimeError(f"PyRFC indisponível: {_PYRFC_IMPORT_ERROR}")


def _env_default(environment: str, field_name: str, default: str = "") -> str:
    env = _normalize_environment(environment)
    field_key = str(field_name or "").strip().upper()
    if not field_key:
        return default

    for candidate in (f"SAP_{env}_FI_{field_key}", f"SAP_FI_{field_key}"):
        value = os.getenv(candidate, "").strip()
        if value:
            return value
    return default


def _env_user(environment: str, default: str = "") -> str:
    env = _normalize_environment(environment)
    return str(os.getenv(f"SAP_{env}_USER", default) or default).strip()


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


def _apply_default_payload(environment: str, branch: str, payload: dict[str, Any]) -> dict[str, Any]:
    mode = str(payload.get("data_mode") or "manual").strip().lower()
    if mode not in {"default", "env"}:
        return payload

    merged = dict(payload)
    fields_by_branch = {
        "cliente": [
            "company_code",
            "posting_date",
            "document_date",
            "currency",
            "header_text",
            "reference",
            "username",
            "customer_account",
            "revenue_gl_account",
            "amount",
            "tax_code",
            "tax_amount",
            "tax_rate",
            "tax_gl_account",
            "item_text",
        ],
        "fornecedor": [
            "company_code",
            "posting_date",
            "document_date",
            "currency",
            "header_text",
            "reference",
            "username",
            "vendor_account",
            "expense_gl_account",
            "amount",
            "tax_code",
            "tax_amount",
            "tax_rate",
            "tax_gl_account",
            "item_text",
        ],
        "razao": [
            "company_code",
            "posting_date",
            "document_date",
            "currency",
            "header_text",
            "reference",
            "username",
            "debit_gl_account",
            "credit_gl_account",
            "amount",
            "tax_code",
            "tax_amount",
            "tax_direction",
            "tax_rate",
            "tax_gl_account",
            "item_text",
        ],
    }
    for field_name in fields_by_branch.get(str(branch or "").strip().lower(), []):
        current = str(merged.get(field_name) or "").strip()
        if current:
            continue
        fallback = "credit" if field_name == "tax_direction" else ""
        if field_name in {"tax_amount"}:
            fallback = "0"
        if field_name in {"currency"}:
            fallback = "EUR"
        merged[field_name] = _env_default(environment, field_name, fallback)
    if not str(merged.get("reference") or "").strip():
        merged["reference"] = _next_reference(
            _env_default(environment, "reference_prefix", "RFC-TEST")
        )
    return merged


def _to_decimal(value: Any, *, default: str = "0") -> Decimal:
    raw = str(value if value is not None else default).strip().replace(",", ".")
    if not raw:
        raw = default
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Valor numérico inválido: {value!r}") from exc


def _to_amount_text(value: Any) -> str:
    amount = _to_decimal(value)
    return f"{amount:.2f}"


def _to_date(value: Any) -> date:
    raw = str(value or "").strip()
    if not raw:
        raise ValueError("date value required")
    return date.fromisoformat(raw)


def _json_safe(value: Any) -> Any:
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {key: _json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_json_safe(item) for item in value]
    return value


def _build_header(payload: dict[str, Any], *, doc_type: str) -> dict[str, Any]:
    posting_date = str(payload.get("posting_date") or "").strip()
    document_date = str(payload.get("document_date") or "").strip()
    posting_date_value = _to_date(posting_date)
    document_date_value = _to_date(document_date)
    company_code = str(payload.get("company_code") or "").strip().upper()
    currency = str(payload.get("currency") or "EUR").strip().upper()
    header_text = str(payload.get("header_text") or "").strip()
    reference = str(payload.get("reference") or "").strip()
    username = str(payload.get("username") or _env_user(payload.get("environment"), "")).strip()

    if not posting_date:
        raise ValueError("posting_date é obrigatório.")
    if not document_date:
        raise ValueError("document_date é obrigatório.")
    if not company_code:
        raise ValueError("company_code é obrigatório.")

    return {
        "USERNAME": username,
        "COMP_CODE": company_code,
        "DOC_DATE": document_date_value,
        "PSTNG_DATE": posting_date_value,
        "DOC_TYPE": doc_type,
        "HEADER_TXT": header_text,
        "REF_DOC_NO": reference,
        "FISC_YEAR": f"{posting_date_value.year}",
        "BUS_ACT": "RFBU",
    }


def _build_tax_line(
    *,
    itemno: int,
    tax_code: str,
    tax_amount: Any,
    tax_rate: Any = "",
    gl_account: str = "",
    cond_key: str = "",
) -> dict[str, Any] | None:
    tax_code = str(tax_code or "").strip().upper()
    if not tax_code:
        return None

    line = {
        "ITEMNO_ACC": str(itemno),
        "TAX_CODE": tax_code,
    }
    if str(tax_rate or "").strip():
        line["TAX_RATE"] = str(tax_rate).strip()
    if str(gl_account or "").strip():
        line["GL_ACCOUNT"] = str(gl_account).strip().upper()
    if str(cond_key or "").strip():
        line["COND_KEY"] = str(cond_key).strip()
    return line


def _base_currency_row(itemno: int, currency: str, amount: Any) -> dict[str, Any]:
    return {
        "ITEMNO_ACC": str(itemno),
        "CURRENCY": str(currency).strip().upper(),
        "AMT_DOCCUR": _to_amount_text(amount),
    }


def _check_return_tables(response: dict[str, Any]) -> list[dict[str, Any]]:
    rows = response.get("RETURN") or []
    if isinstance(rows, dict):
        rows = [rows]
    return [dict(row) for row in rows if isinstance(row, dict)]


def _has_bapi_error(rows: list[dict[str, Any]]) -> bool:
    return any(str(row.get("TYPE") or "").strip().upper() in {"E", "A", "X"} for row in rows)


def _join_return_messages(rows: list[dict[str, Any]]) -> str:
    parts = []
    for row in rows:
        msg_type = str(row.get("TYPE") or "").strip().upper()
        msg = str(row.get("MESSAGE") or "").strip()
        if not msg:
            continue
        if msg_type:
            parts.append(f"{msg_type}: {msg}")
        else:
            parts.append(msg)
    return " | ".join(parts)


def _build_customer_payload(environment: str, payload: dict[str, Any]) -> dict[str, Any]:
    net_amount = _to_decimal(payload.get("amount"))
    tax_amount = _to_decimal(payload.get("tax_amount") or "0")
    gross_amount = net_amount + tax_amount
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    customer_account = str(payload.get("customer_account") or "").strip().upper()
    revenue_gl_account = str(payload.get("revenue_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()

    if not customer_account:
        raise ValueError("customer_account é obrigatório para documentos de Cliente.")
    if not revenue_gl_account:
        raise ValueError("revenue_gl_account é obrigatório para documentos de Cliente.")

    accountreceivable = [
        {
            "ITEMNO_ACC": "1",
            "CUSTOMER": customer_account,
            "ITEM_TEXT": item_text,
        }
    ]
    accountgl = [
        {
            "ITEMNO_ACC": "2",
            "GL_ACCOUNT": revenue_gl_account,
            "ITEM_TEXT": item_text,
        }
    ]
    accounttax = []
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=-tax_amount,
        tax_rate=payload.get("tax_rate"),
        gl_account=str(payload.get("tax_gl_account") or "").strip().upper(),
    )
    if tax_line:
        accounttax.append(tax_line)

    currencyamount = [
        _base_currency_row(1, currency, gross_amount),
        _base_currency_row(2, currency, -net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, -tax_amount))

    return {
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_cliente", "DR"),
        ),
        "ACCOUNTRECEIVABLE": accountreceivable,
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }


def _build_vendor_payload(environment: str, payload: dict[str, Any]) -> dict[str, Any]:
    net_amount = _to_decimal(payload.get("amount"))
    tax_amount = _to_decimal(payload.get("tax_amount") or "0")
    gross_amount = net_amount + tax_amount
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    vendor_account = str(payload.get("vendor_account") or "").strip().upper()
    expense_gl_account = str(payload.get("expense_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()

    if not vendor_account:
        raise ValueError("vendor_account é obrigatório para documentos de Fornecedor.")
    if not expense_gl_account:
        raise ValueError("expense_gl_account é obrigatório para documentos de Fornecedor.")

    accountpayable = [
        {
            "ITEMNO_ACC": "1",
            "VENDOR_NO": vendor_account,
            "ITEM_TEXT": item_text,
        }
    ]
    accountgl = [
        {
            "ITEMNO_ACC": "2",
            "GL_ACCOUNT": expense_gl_account,
            "ITEM_TEXT": item_text,
        }
    ]
    accounttax = []
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=tax_amount,
        tax_rate=payload.get("tax_rate"),
        gl_account=str(payload.get("tax_gl_account") or "").strip().upper(),
    )
    if tax_line:
        accounttax.append(tax_line)

    currencyamount = [
        _base_currency_row(1, currency, -gross_amount),
        _base_currency_row(2, currency, net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, tax_amount))

    return {
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_fornecedor", "KR"),
        ),
        "ACCOUNTPAYABLE": accountpayable,
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }


def _build_gl_payload(environment: str, payload: dict[str, Any]) -> dict[str, Any]:
    amount = _to_decimal(payload.get("amount"))
    tax_amount = _to_decimal(payload.get("tax_amount") or "0")
    tax_direction = str(payload.get("tax_direction") or "credit").strip().lower()
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    debit_gl_account = str(payload.get("debit_gl_account") or "").strip().upper()
    credit_gl_account = str(payload.get("credit_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()

    if not debit_gl_account:
        raise ValueError("debit_gl_account é obrigatório para documentos de Razão.")
    if not credit_gl_account:
        raise ValueError("credit_gl_account é obrigatório para documentos de Razão.")

    accountgl = [
        {
            "ITEMNO_ACC": "1",
            "GL_ACCOUNT": debit_gl_account,
            "ITEM_TEXT": item_text,
        },
        {
            "ITEMNO_ACC": "2",
            "GL_ACCOUNT": credit_gl_account,
            "ITEM_TEXT": item_text,
        },
    ]
    accounttax = []
    tax_line_amount = tax_amount if tax_direction == "debit" else -tax_amount
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=tax_line_amount,
        tax_rate=payload.get("tax_rate"),
        gl_account=str(payload.get("tax_gl_account") or "").strip().upper(),
    )
    if tax_line:
        accounttax.append(tax_line)

    currencyamount = [
        _base_currency_row(1, currency, amount),
        _base_currency_row(2, currency, -amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, tax_line_amount))

    return {
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_razao", "SA"),
        ),
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }


def _build_bapi_payload(branch: str, environment: str, payload: dict[str, Any]) -> dict[str, Any]:
    branch_key = str(branch or "").strip().lower()
    if branch_key == "cliente":
        return _build_customer_payload(environment, payload)
    if branch_key == "fornecedor":
        return _build_vendor_payload(environment, payload)
    if branch_key == "razao":
        return _build_gl_payload(environment, payload)
    raise ValueError(f"Tipo de documento FI não suportado: {branch}")


def _call_bapi(connection: Any, function_name: str, payload: dict[str, Any]) -> dict[str, Any]:
    response = connection.call(function_name, **payload)
    if not isinstance(response, dict):
        return {}
    return response


def _extract_document_number(response: dict[str, Any]) -> str:
    for key in ("OBJ_KEY", "BELNR", "DOC_NO", "DOCUMENTNUMBER", "DOC_NUMBER"):
        value = str(response.get(key) or "").strip()
        if value:
            return value
    return ""


def post_fi_document(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    _require_pyrfc()
    connection_params = build_connection_params(environment)
    payload = _apply_default_payload(environment, branch, payload)
    bapi_payload = _build_bapi_payload(branch, environment, payload)

    connection = Connection(**connection_params)  # type: ignore[misc]
    try:
        check_response = _call_bapi(connection, "BAPI_ACC_DOCUMENT_CHECK", bapi_payload)
        check_return = _check_return_tables(check_response)
        if _has_bapi_error(check_return):
            return FiDocumentResult(
                ok=False,
                status="ERRO",
                message=_join_return_messages(check_return) or "BAPI_ACC_DOCUMENT_CHECK devolveu erro.",
                branch=branch,
                company_code=str(payload.get("company_code") or "").strip().upper(),
                check_return=check_return,
                payload=_json_safe(bapi_payload),
            )

        post_response = _call_bapi(connection, "BAPI_ACC_DOCUMENT_POST", bapi_payload)
        post_return = _check_return_tables(post_response)
        if _has_bapi_error(post_return):
            return FiDocumentResult(
                ok=False,
                status="ERRO",
                message=_join_return_messages(post_return) or "BAPI_ACC_DOCUMENT_POST devolveu erro.",
                branch=branch,
                company_code=str(payload.get("company_code") or "").strip().upper(),
                check_return=check_return,
                post_return=post_return,
                payload=_json_safe(bapi_payload),
            )

        commit_response = _call_bapi(connection, "BAPI_TRANSACTION_COMMIT", {"WAIT": "X"})
        commit_return = _check_return_tables(commit_response)
        document_number = _extract_document_number(post_response) or _extract_document_number(check_response)

        message = _join_return_messages(post_return or check_return) or "Documento FI processado com sucesso."
        return FiDocumentResult(
            ok=True,
            status="SUCESSO",
            message=message,
            branch=branch,
            company_code=str(payload.get("company_code") or "").strip().upper(),
            document_number=document_number,
            check_return=check_return,
            post_return=post_return,
            commit_return=commit_return,
            payload=_json_safe(bapi_payload),
        )
    finally:
        try:
            connection.close()
        except Exception:
            pass
