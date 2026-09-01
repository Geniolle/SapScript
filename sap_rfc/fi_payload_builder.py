from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any

from .fi_config import env_alias_default, env_default, payment_method_default, normalize_environment, _next_reference


MONEY_QUANT = Decimal("0.01")


@dataclass
class TaxCalculationResult:
    tax_code: str
    base_amount: Decimal
    tax_amount: Decimal
    tax_rate: str = ""
    tax_gl_account: str = ""
    cond_key: str = ""
    acct_key: str = ""
    tax_date: str = ""
    taxjurcode: str = ""
    taxjurcode_deep: str = ""
    taxjurcode_level: str = ""
    source: str = "payload"


def _format_itemno_tax(itemno: int) -> str:
    return f"{int(itemno):06d}"


def _tax_rfc_name(environment: str) -> str:
    return env_default(
        environment,
        "tax_calc_rfc",
        os.getenv("SAP_FI_TAX_CALC_RFC", "").strip() or "BBP_CALCULATE_TAX_FRM_NET_40B",
    )


def _tax_calc_response_rows(response: dict[str, Any]) -> list[dict[str, Any]]:
    rows = response.get("T_MWDAT") or response.get("T_MWDAT[]") or []
    if isinstance(rows, dict):
        rows = [rows]
    return [dict(row) for row in rows if isinstance(row, dict)]


def _extract_decimal(value: Any, default: str = "0") -> Decimal:
    raw = str(value if value is not None else default).strip()
    if not raw:
        raw = default
    raw = raw.replace(",", ".")
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Valor monetário inválido: {value!r}") from exc


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
    username = str(payload.get("username") or env_default(payload.get("environment"), "user", "")).strip()

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
    acct_key: str = "",
    itemno_tax: str = "",
    tax_date: Any = "",
    taxjurcode: str = "",
    taxjurcode_deep: str = "",
    taxjurcode_level: str = "",
) -> dict[str, Any] | None:
    tax_code = str(tax_code or "").strip().upper()
    if not tax_code:
        return None

    line = {"ITEMNO_ACC": str(itemno), "TAX_CODE": tax_code}
    if str(tax_rate or "").strip():
        line["TAX_RATE"] = str(tax_rate).strip()
    if str(gl_account or "").strip():
        line["GL_ACCOUNT"] = str(gl_account).strip().upper()
    if str(cond_key or "").strip():
        line["COND_KEY"] = str(cond_key).strip()
    if str(acct_key or "").strip():
        line["ACCT_KEY"] = str(acct_key).strip()
    if str(itemno_tax or "").strip():
        line["ITEMNO_TAX"] = _format_itemno_tax(int(str(itemno_tax).strip()))
    if str(tax_date or "").strip():
        line["TAX_DATE"] = str(tax_date).strip()
    if str(taxjurcode or "").strip():
        line["TAXJURCODE"] = str(taxjurcode).strip()
    if str(taxjurcode_deep or "").strip():
        line["TAXJURCODE_DEEP"] = str(taxjurcode_deep).strip()
    if str(taxjurcode_level or "").strip():
        line["TAXJURCODE_LEVEL"] = str(taxjurcode_level).strip()
    return line


def _build_withholding_tax_line(
    *,
    itemno: int,
    wt_type: str,
    wt_code: str,
    base_amount: Any,
    manual_amount: Any = "",
) -> dict[str, Any] | None:
    wt_type = str(wt_type or "").strip().upper()
    wt_code = str(wt_code or "").strip().upper()
    if not wt_type or not wt_code:
        return None

    line = {
        "ITEMNO_ACC": str(itemno),
        "WT_TYPE": wt_type,
        "WT_CODE": wt_code,
        "BAS_AMT_LC": _to_amount_text(base_amount),
        "BAS_AMT_TC": _to_amount_text(base_amount),
        "BAS_AMT_IND": "X",
    }
    if str(manual_amount or "").strip():
        amount_text = _to_amount_text(manual_amount)
        line["MAN_AMT_LC"] = amount_text
        line["MAN_AMT_TC"] = amount_text
        line["MAN_AMT_IND"] = "X"
        line["AWH_AMT_LC"] = amount_text
        line["AWH_AMT_TC"] = amount_text
    return line


def _read_first_rfc_table_row(connection: Any, table_name: str, fields: list[str], options: list[dict[str, str]]) -> dict[str, str] | None:
    response = connection.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table_name,
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": field} for field in fields],
        OPTIONS=options,
        ROWCOUNT=1,
    )
    rows = response.get("DATA") or []
    if not rows:
        return None
    wa = str(rows[0].get("WA") or "")
    parts = [part.strip() for part in wa.split("|")]
    if len(parts) < len(fields):
        parts += [""] * (len(fields) - len(parts))
    return {field: parts[index] if index < len(parts) else "" for index, field in enumerate(fields)}


def _resolve_master_withholding_tax(
    connection: Any | None,
    *,
    table_name: str,
    company_code: str,
    account_field: str,
    account_number: str,
) -> dict[str, str]:
    if connection is None:
        return {}
    company_code = str(company_code or "").strip().upper()
    account_number = str(account_number or "").strip().upper()
    if not company_code or not account_number:
        return {}
    try:
        row = _read_first_rfc_table_row(
            connection,
            table_name,
            ["BUKRS", account_field, "WITHT", "WT_WITHCD"],
            [
                {"TEXT": f"BUKRS = '{company_code}'"},
                {"TEXT": f"AND {account_field} = '{account_number}'"},
            ],
        )
    except Exception:
        return {}
    if not row:
        return {}
    result = {
        "withholding_tax_type": str(row.get("WITHT") or "").strip().upper(),
        "withholding_tax_code": str(row.get("WT_WITHCD") or "").strip().upper(),
    }
    if result["withholding_tax_type"] and result["withholding_tax_code"]:
        try:
            country_row = _read_first_rfc_table_row(
                connection,
                "T001",
                ["BUKRS", "LAND1"],
                [{"TEXT": f"BUKRS = '{company_code}'"}],
            )
            country = str(country_row.get("LAND1") or "").strip().upper() if country_row else ""
            if country:
                tax_row = _read_first_rfc_table_row(
                    connection,
                    "T059Z",
                    ["LAND1", "WITHT", "WT_WITHCD", "QPROZ", "QSATZ"],
                    [
                        {"TEXT": f"LAND1 = '{country}'"},
                        {"TEXT": f"AND WITHT = '{result['withholding_tax_type']}'"},
                        {"TEXT": f"AND WT_WITHCD = '{result['withholding_tax_code']}'"},
                    ],
                )
                if tax_row:
                    result["withholding_tax_rate"] = str(tax_row.get("QSATZ") or "").strip()
        except Exception:
            pass
    return result


def _base_currency_row(
    itemno: int,
    currency: str,
    amount: Any,
    *,
    curr_type: str = "00",
    amt_base: Any = "",
    tax_amt: Any = "",
) -> dict[str, Any]:
    row = {
        "ITEMNO_ACC": str(itemno),
        "CURR_TYPE": str(curr_type or "00").strip() or "00",
        "CURRENCY": str(currency).strip().upper(),
        "AMT_DOCCUR": _to_amount_text(amount),
    }
    if str(amt_base or "").strip():
        row["AMT_BASE"] = _to_amount_text(amt_base)
    if str(tax_amt or "").strip():
        row["TAX_AMT"] = _to_amount_text(tax_amt)
    return row


def _check_return_tables(response: dict[str, Any]) -> list[dict[str, Any]]:
    rows = response.get("RETURN") or response.get("RETURN[]") or []
    if isinstance(rows, dict):
        rows = [rows]
    return [dict(row) for row in rows if isinstance(row, dict)]


def _join_return_messages(rows: list[dict[str, Any]]) -> str:
    parts: list[str] = []
    for row in rows:
        message = str(row.get("MESSAGE") or "").strip()
        if message:
            parts.append(message)
    return " | ".join(parts)


def _validate_currencyamount_balance(currencyamount: list[dict[str, Any]]) -> None:
    balances: dict[tuple[str, str], Decimal] = {}
    for row in currencyamount:
        curr_type = str(row.get("CURR_TYPE") or "00").strip() or "00"
        currency = str(row.get("CURRENCY") or "").strip().upper()
        key = (curr_type, currency)
        balances[key] = balances.get(key, Decimal("0")) + _extract_decimal(row.get("AMT_DOCCUR") or "0")

    errors: list[str] = []
    for (curr_type, currency), balance in balances.items():
        rounded = balance.quantize(MONEY_QUANT)
        if rounded != Decimal("0.00"):
            label = f"{currency or 'MOEDA'} / CURR_TYPE {curr_type}"
            errors.append(f"{label} saldo = {rounded:.2f}")

    if errors:
        raise ValueError("Documento FI não balanceado:\n" + "\n".join(errors))


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
            "payment_method",
            "withholding_tax_type",
            "withholding_tax_code",
            "withholding_tax_base_amount",
            "withholding_tax_amount",
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
            "payment_method",
            "withholding_tax_type",
            "withholding_tax_code",
            "withholding_tax_base_amount",
            "withholding_tax_amount",
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
        if field_name == "payment_method":
            fallback = payment_method_default(environment, branch, fallback)
        merged[field_name] = env_default(environment, field_name, fallback)
    if not str(merged.get("reference") or "").strip():
        merged["reference"] = _next_reference(env_default(environment, "reference_prefix", "RFC-TEST"))
    return merged


def _resolve_tax_calculation(
    connection: Any,
    environment: str,
    *,
    company_code: str,
    posting_date: str,
    currency: str,
    base_amount: Decimal,
    tax_code: str,
    tax_rate_hint: Any = "",
    tax_amount_hint: Any = "",
    tax_gl_account_hint: str = "",
    taxjurcode_hint: str = "",
) -> TaxCalculationResult | None:
    tax_code = str(tax_code or "").strip().upper()
    if not tax_code:
        return None

    result = TaxCalculationResult(
        tax_code=tax_code,
        base_amount=abs(base_amount).quantize(MONEY_QUANT),
        tax_amount=_extract_decimal(tax_amount_hint or "0").quantize(MONEY_QUANT),
        tax_rate=str(tax_rate_hint or "").strip(),
        tax_gl_account=str(tax_gl_account_hint or "").strip().upper(),
        cond_key=env_default(environment, "tax_cond_key", ""),
        acct_key=env_default(environment, "tax_acct_key", ""),
        tax_date=str(posting_date or "").strip(),
        taxjurcode=str(taxjurcode_hint or env_default(environment, "taxjurcode", "")).strip().upper(),
        source="payload",
    )

    rfc_name = _tax_rfc_name(environment)
    if not rfc_name:
        return result

    try:
        response = connection.call(
            rfc_name,
            I_BUKRS=str(company_code).strip().upper(),
            I_MWSKZ=tax_code,
            I_TXJCD=str(taxjurcode_hint or "").strip().upper(),
            I_WAERS=str(currency or "EUR").strip().upper(),
            I_WRBTR=f"{abs(base_amount):.2f}",
            I_PRSDT=_to_date(posting_date),
            I_PROTOKOLL="X",
        )
    except Exception:
        return result

    rows = _tax_calc_response_rows(dict(response or {}))
    row = rows[0] if rows else {}

    calc_base = row.get("KAWRT")
    calc_tax_amount = row.get("WMWST") or row.get("FWSTE") or row.get("E_FWSTE")
    calc_tax_rate = row.get("MSATZ")
    calc_tax_gl_account = row.get("HKONT")
    calc_cond_key = row.get("KSCHL")
    calc_acct_key = row.get("KTOSL")
    calc_taxjurcode = row.get("TXJCD")
    calc_taxjurcode_deep = row.get("TXJCD_DEEP")
    calc_taxjurcode_level = row.get("TXJLV")

    if calc_base not in (None, ""):
        result.base_amount = abs(_extract_decimal(calc_base)).quantize(MONEY_QUANT)
    if calc_tax_amount not in (None, ""):
        result.tax_amount = abs(_extract_decimal(calc_tax_amount)).quantize(MONEY_QUANT)
    if calc_tax_rate not in (None, ""):
        result.tax_rate = str(calc_tax_rate).strip()
    if calc_tax_gl_account not in (None, ""):
        result.tax_gl_account = str(calc_tax_gl_account).strip().upper()
    if calc_cond_key not in (None, ""):
        result.cond_key = str(calc_cond_key).strip().upper()
    if calc_acct_key not in (None, ""):
        result.acct_key = str(calc_acct_key).strip().upper()
    if calc_taxjurcode not in (None, ""):
        result.taxjurcode = str(calc_taxjurcode).strip().upper()
    if calc_taxjurcode_deep not in (None, ""):
        result.taxjurcode_deep = str(calc_taxjurcode_deep).strip().upper()
    if calc_taxjurcode_level not in (None, ""):
        result.taxjurcode_level = str(calc_taxjurcode_level).strip().upper()
    result.source = rfc_name

    manual_tax_amount = _extract_decimal(tax_amount_hint or "0").quantize(MONEY_QUANT)
    if manual_tax_amount and manual_tax_amount != Decimal("0.00") and result.tax_amount != manual_tax_amount:
        raise ValueError(
            f"IVA informado divergente do calculado pelo SAP: informado={manual_tax_amount:.2f} calculado={result.tax_amount:.2f}"
        )
    manual_tax_rate = str(tax_rate_hint or "").strip()
    if manual_tax_rate and result.tax_rate and manual_tax_rate != result.tax_rate:
        raise ValueError(
            f"Tax rate informado divergente do calculado pelo SAP: informado={manual_tax_rate} calculado={result.tax_rate}"
        )

    if not result.tax_gl_account:
        result.tax_gl_account = str(tax_gl_account_hint or "").strip().upper()

    return result


def _build_customer_payload(environment: str, payload: dict[str, Any], connection: Any | None = None) -> dict[str, Any]:
    net_amount = _to_decimal(payload.get("amount"))
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    customer_account = str(payload.get("customer_account") or "").strip().upper()
    revenue_gl_account = str(payload.get("revenue_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()
    tax_result = None
    if tax_code:
        tax_result = _resolve_tax_calculation(
            connection,
            environment,
            company_code=str(payload.get("company_code") or "").strip().upper(),
            posting_date=str(payload.get("posting_date") or "").strip(),
            currency=currency,
            base_amount=net_amount,
            tax_code=tax_code,
            tax_rate_hint=payload.get("tax_rate"),
            tax_amount_hint=payload.get("tax_amount"),
            tax_gl_account_hint=str(payload.get("tax_gl_account") or "").strip().upper(),
        )
    tax_amount = tax_result.tax_amount if tax_result else _extract_decimal(payload.get("tax_amount") or "0")
    if not tax_code:
        tax_amount = Decimal("0")
    gross_amount = net_amount + tax_amount

    if not customer_account:
        raise ValueError("customer_account é obrigatório para documentos de Cliente.")
    if not revenue_gl_account:
        raise ValueError("revenue_gl_account é obrigatório para documentos de Cliente.")

    withholding_tax_type = env_default(environment, "withholding_tax_type", str(payload.get("withholding_tax_type") or ""))
    withholding_tax_code = env_default(environment, "withholding_tax_code", str(payload.get("withholding_tax_code") or ""))
    if connection and (not withholding_tax_type or not withholding_tax_code):
        master_wt = _resolve_master_withholding_tax(
            connection,
            table_name="KNBW",
            company_code=str(payload.get("company_code") or ""),
            account_field="KUNNR",
            account_number=customer_account,
        )
        withholding_tax_type = withholding_tax_type or master_wt.get("withholding_tax_type", "")
        withholding_tax_code = withholding_tax_code or master_wt.get("withholding_tax_code", "")
        if not str(payload.get("withholding_tax_amount") or "").strip():
            withholding_tax_rate = str(master_wt.get("withholding_tax_rate") or "").strip()
            if withholding_tax_rate:
                tax_amount = (net_amount * _to_decimal(withholding_tax_rate) / Decimal("100")).quantize(MONEY_QUANT)
                payload["withholding_tax_amount"] = _to_amount_text(tax_amount)

    accountreceivable = [{"ITEMNO_ACC": "1", "CUSTOMER": customer_account, "ITEM_TEXT": item_text}]
    payment_method = payment_method_default(environment, "cliente", str(payload.get("payment_method") or "").strip().upper())
    if payment_method:
        accountreceivable[0]["PYMT_METH"] = payment_method
    if withholding_tax_code:
        accountreceivable[0]["W_TAX_CODE"] = withholding_tax_code
    accountgl = [{"ITEMNO_ACC": "2", "GL_ACCOUNT": revenue_gl_account, "ITEM_TEXT": item_text}]
    accounttax = []
    if tax_code:
        accountgl[0]["TAX_CODE"] = tax_code
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=-tax_amount,
        tax_rate=tax_result.tax_rate if tax_result and tax_result.tax_rate else payload.get("tax_rate"),
        gl_account=(tax_result.tax_gl_account if tax_result and tax_result.tax_gl_account else str(payload.get("tax_gl_account") or "").strip().upper()),
        cond_key=tax_result.cond_key if tax_result else "",
        acct_key=tax_result.acct_key if tax_result else "",
        itemno_tax="2",
        tax_date=tax_result.tax_date if tax_result else str(payload.get("posting_date") or "").strip(),
        taxjurcode=tax_result.taxjurcode if tax_result else "",
        taxjurcode_deep=tax_result.taxjurcode_deep if tax_result else "",
        taxjurcode_level=tax_result.taxjurcode_level if tax_result else "",
    )
    if accountgl:
        accountgl[0]["ITEMNO_TAX"] = _format_itemno_tax(3)
    if tax_line:
        accounttax.append(tax_line)
    withholding_tax_line = _build_withholding_tax_line(
        itemno=1,
        wt_type=withholding_tax_type,
        wt_code=withholding_tax_code,
        base_amount=env_default(environment, "withholding_tax_base_amount", str(payload.get("withholding_tax_base_amount") or gross_amount)),
        manual_amount=env_default(environment, "withholding_tax_amount", str(payload.get("withholding_tax_amount") or "")),
    )

    currencyamount = [
        _base_currency_row(1, currency, gross_amount),
        _base_currency_row(2, currency, -net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, -tax_amount, amt_base=net_amount, tax_amt=-tax_amount))
    _validate_currencyamount_balance(currencyamount)

    result = {
        "DOCUMENTHEADER": _build_header(payload, doc_type=env_default(environment, "doc_type_cliente", "DR")),
        "ACCOUNTRECEIVABLE": accountreceivable,
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }
    if withholding_tax_line:
        result["ACCOUNTWT"] = [withholding_tax_line]
    return result


def _build_vendor_payload(environment: str, payload: dict[str, Any], connection: Any | None = None) -> dict[str, Any]:
    net_amount = _to_decimal(payload.get("amount"))
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    vendor_account = str(payload.get("vendor_account") or "").strip().upper()
    expense_gl_account = str(payload.get("expense_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()
    tax_result = None
    if tax_code:
        tax_result = _resolve_tax_calculation(
            connection,
            environment,
            company_code=str(payload.get("company_code") or "").strip().upper(),
            posting_date=str(payload.get("posting_date") or "").strip(),
            currency=currency,
            base_amount=net_amount,
            tax_code=tax_code,
            tax_rate_hint=payload.get("tax_rate"),
            tax_amount_hint=payload.get("tax_amount"),
            tax_gl_account_hint=str(payload.get("tax_gl_account") or "").strip().upper(),
        )
    tax_amount = tax_result.tax_amount if tax_result else _extract_decimal(payload.get("tax_amount") or "0")
    if not tax_code:
        tax_amount = Decimal("0")
    gross_amount = net_amount + tax_amount

    if not vendor_account:
        raise ValueError("vendor_account é obrigatório para documentos de Fornecedor.")
    if not expense_gl_account:
        raise ValueError("expense_gl_account é obrigatório para documentos de Fornecedor.")

    withholding_tax_type = env_default(environment, "withholding_tax_type", str(payload.get("withholding_tax_type") or ""))
    withholding_tax_code = env_default(environment, "withholding_tax_code", str(payload.get("withholding_tax_code") or ""))
    if connection and (not withholding_tax_type or not withholding_tax_code):
        master_wt = _resolve_master_withholding_tax(
            connection,
            table_name="LFBW",
            company_code=str(payload.get("company_code") or ""),
            account_field="LIFNR",
            account_number=vendor_account,
        )
        withholding_tax_type = withholding_tax_type or master_wt.get("withholding_tax_type", "")
        withholding_tax_code = withholding_tax_code or master_wt.get("withholding_tax_code", "")
        if not str(payload.get("withholding_tax_amount") or "").strip():
            withholding_tax_rate = str(master_wt.get("withholding_tax_rate") or "").strip()
            if withholding_tax_rate:
                tax_amount = (net_amount * _to_decimal(withholding_tax_rate) / Decimal("100")).quantize(MONEY_QUANT)
                payload["withholding_tax_amount"] = _to_amount_text(tax_amount)

    accountpayable = [{"ITEMNO_ACC": "1", "VENDOR_NO": vendor_account, "ITEM_TEXT": item_text}]
    payment_method = payment_method_default(environment, "fornecedor", str(payload.get("payment_method") or "").strip().upper())
    if payment_method:
        accountpayable[0]["PYMT_METH"] = payment_method
    if withholding_tax_code:
        accountpayable[0]["W_TAX_CODE"] = withholding_tax_code
    accountgl = [{"ITEMNO_ACC": "2", "GL_ACCOUNT": expense_gl_account, "ITEM_TEXT": item_text}]
    accounttax = []
    if tax_code:
        accountgl[0]["TAX_CODE"] = tax_code
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=tax_amount,
        tax_rate=tax_result.tax_rate if tax_result and tax_result.tax_rate else payload.get("tax_rate"),
        gl_account=(tax_result.tax_gl_account if tax_result and tax_result.tax_gl_account else str(payload.get("tax_gl_account") or "").strip().upper()),
        cond_key=tax_result.cond_key if tax_result else "",
        acct_key=tax_result.acct_key if tax_result else "",
        itemno_tax="2",
        tax_date=tax_result.tax_date if tax_result else str(payload.get("posting_date") or "").strip(),
        taxjurcode=tax_result.taxjurcode if tax_result else "",
        taxjurcode_deep=tax_result.taxjurcode_deep if tax_result else "",
        taxjurcode_level=tax_result.taxjurcode_level if tax_result else "",
    )
    if accountgl:
        accountgl[0]["ITEMNO_TAX"] = _format_itemno_tax(3)
    if tax_line:
        accounttax.append(tax_line)
    withholding_tax_line = _build_withholding_tax_line(
        itemno=1,
        wt_type=withholding_tax_type,
        wt_code=withholding_tax_code,
        base_amount=env_default(environment, "withholding_tax_base_amount", str(payload.get("withholding_tax_base_amount") or net_amount)),
        manual_amount=env_default(environment, "withholding_tax_amount", str(payload.get("withholding_tax_amount") or "")),
    )

    currencyamount = [
        _base_currency_row(1, currency, -gross_amount),
        _base_currency_row(2, currency, net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, tax_amount, amt_base=net_amount, tax_amt=tax_amount))
    _validate_currencyamount_balance(currencyamount)

    result = {
        "DOCUMENTHEADER": _build_header(payload, doc_type=env_default(environment, "doc_type_fornecedor", "KR")),
        "ACCOUNTPAYABLE": accountpayable,
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }
    if withholding_tax_line:
        result["ACCOUNTWT"] = [withholding_tax_line]
    return result


def _build_gl_payload(environment: str, payload: dict[str, Any], connection: Any | None = None) -> dict[str, Any]:
    amount = _to_decimal(payload.get("amount"))
    tax_direction = str(payload.get("tax_direction") or "credit").strip().lower()
    tax_code = str(payload.get("tax_code") or "").strip().upper()
    debit_gl_account = str(payload.get("debit_gl_account") or "").strip().upper()
    credit_gl_account = str(payload.get("credit_gl_account") or "").strip().upper()
    item_text = str(payload.get("item_text") or payload.get("header_text") or "").strip()
    currency = str(payload.get("currency") or "EUR").strip().upper()
    tax_result = None
    if tax_code:
        tax_result = _resolve_tax_calculation(
            connection,
            environment,
            company_code=str(payload.get("company_code") or "").strip().upper(),
            posting_date=str(payload.get("posting_date") or "").strip(),
            currency=currency,
            base_amount=amount,
            tax_code=tax_code,
            tax_rate_hint=payload.get("tax_rate"),
            tax_amount_hint=payload.get("tax_amount"),
            tax_gl_account_hint=str(payload.get("tax_gl_account") or "").strip().upper(),
        )
    tax_amount = tax_result.tax_amount if tax_result else _extract_decimal(payload.get("tax_amount") or "0")
    if not tax_code:
        tax_amount = Decimal("0")

    if not debit_gl_account:
        raise ValueError("debit_gl_account é obrigatório para documentos de Razão.")
    if not credit_gl_account:
        raise ValueError("credit_gl_account é obrigatório para documentos de Razão.")

    accountgl = [
        {"ITEMNO_ACC": "1", "GL_ACCOUNT": debit_gl_account, "ITEM_TEXT": item_text},
        {"ITEMNO_ACC": "2", "GL_ACCOUNT": credit_gl_account, "ITEM_TEXT": item_text},
    ]
    accounttax = []
    tax_line_amount = tax_amount if tax_direction == "debit" else -tax_amount
    taxable_itemno = "1" if tax_direction == "debit" else "2"
    gross_amount = amount + tax_amount
    tax_line = _build_tax_line(
        itemno=3,
        tax_code=tax_code,
        tax_amount=tax_line_amount,
        tax_rate=tax_result.tax_rate if tax_result and tax_result.tax_rate else payload.get("tax_rate"),
        gl_account=(tax_result.tax_gl_account if tax_result and tax_result.tax_gl_account else str(payload.get("tax_gl_account") or "").strip().upper()),
        cond_key=tax_result.cond_key if tax_result else "",
        acct_key=tax_result.acct_key if tax_result else "",
        itemno_tax=taxable_itemno,
        tax_date=tax_result.tax_date if tax_result else str(payload.get("posting_date") or "").strip(),
        taxjurcode=tax_result.taxjurcode if tax_result else "",
        taxjurcode_deep=tax_result.taxjurcode_deep if tax_result else "",
        taxjurcode_level=tax_result.taxjurcode_level if tax_result else "",
    )
    if accountgl:
        accountgl[int(taxable_itemno) - 1]["ITEMNO_TAX"] = _format_itemno_tax(3)
        accountgl[int(taxable_itemno) - 1]["TAX_CODE"] = tax_code if tax_code else accountgl[int(taxable_itemno) - 1].get("TAX_CODE", "")
    if tax_line:
        accounttax.append(tax_line)

    currencyamount = [
        _base_currency_row(1, currency, amount if tax_direction == "debit" else gross_amount),
        _base_currency_row(2, currency, -gross_amount if tax_direction == "debit" else -amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, tax_line_amount, amt_base=amount, tax_amt=tax_line_amount))
    _validate_currencyamount_balance(currencyamount)

    return {
        "DOCUMENTHEADER": _build_header(payload, doc_type=env_default(environment, "doc_type_razao", "SA")),
        "ACCOUNTGL": accountgl,
        "ACCOUNTTAX": accounttax,
        "CURRENCYAMOUNT": currencyamount,
    }


def _build_bapi_payload(branch: str, environment: str, payload: dict[str, Any], connection: Any | None = None) -> dict[str, Any]:
    branch_key = str(branch or "").strip().lower()
    if branch_key == "cliente":
        return _build_customer_payload(environment, payload, connection=connection)
    if branch_key == "fornecedor":
        return _build_vendor_payload(environment, payload, connection=connection)
    if branch_key == "razao":
        return _build_gl_payload(environment, payload, connection=connection)
    raise ValueError(f"Tipo de documento FI não suportado: {branch}")
