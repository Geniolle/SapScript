from __future__ import annotations

import os
import sqlite3
import re
import platform
import shutil
from datetime import date
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
import dataclasses
import json
import subprocess
import sys


try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None

from .fi_config import build_connection_params as fi_build_connection_params
from .fi_payload_builder import (
    _apply_default_payload as fi_apply_default_payload,
    _build_bapi_payload as fi_build_bapi_payload,
)


def _bridge_python_executable() -> Path | None:
    repo_root = Path(__file__).resolve().parents[1]
    candidates = [
        repo_root / ".venv-rfc" / "Scripts" / "python.exe",
        repo_root / ".venv-rfc" / "Scripts" / "pythonw.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _run_post_fi_document_via_bridge(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    python_exe = _bridge_python_executable()
    if python_exe is None:
        runtime = "WSL/Linux" if _is_wsl_runtime() else platform.platform()
        raise RuntimeError(
            "Execução FI via bridge indisponível neste runtime "
            f"({runtime}). Configure SAP_FI_BRIDGE_PYTHON ou WORKFLOW_PYTHON_EXEC "
            "com um Python compatível com PyRFC, ou execute este worker no Windows "
            "com .venv-rfc\\Scripts\\python.exe disponível."
        )

    repo_root = Path(__file__).resolve().parents[1]
    bridge_script = (
        "from pathlib import Path\n"
        "import dataclasses\n"
        "import json\n"
        "import sys\n"
        f"repo_root = Path(r'{repo_root}')\n"
        "sys.path.insert(0, str(repo_root))\n"
        "from sap_rfc.fi_document_service import _post_fi_document_core as _post\n"
        "environment = json.loads(sys.stdin.readline())\n"
        "branch = json.loads(sys.stdin.readline())\n"
        "payload = json.loads(sys.stdin.read() or '{}')\n"
        "result = _post(environment, branch, payload)\n"
        "print(json.dumps(dataclasses.asdict(result), ensure_ascii=False))\n"
    )

    env = os.environ.copy()
    env["SAP_FI_BRIDGE_ACTIVE"] = "1"

    proc = subprocess.run(
        [str(python_exe), "-c", bridge_script],
        input="\n".join(
            [
                json.dumps(environment, ensure_ascii=False),
                json.dumps(branch, ensure_ascii=False),
                json.dumps(payload, ensure_ascii=False),
            ]
        ),
        capture_output=True,
        text=True,
        cwd=str(repo_root),
        env=env,
    )
    stdout = proc.stdout.strip()
    stderr = proc.stderr.strip()
    if proc.returncode != 0:
        detail = stderr or stdout or f"Processo RFC de apoio falhou com código {proc.returncode}."
        raise RuntimeError(detail)
    if not stdout:
        raise RuntimeError("Processo RFC de apoio devolveu resposta vazia.")
    data = json.loads(stdout)
    return FiDocumentResult(**data)


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
    return _env_default(
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
        cond_key=_env_default(environment, "tax_cond_key", ""),
        acct_key=_env_default(environment, "tax_acct_key", ""),
        tax_date=str(posting_date or "").strip(),
        taxjurcode=str(taxjurcode_hint or _env_default(environment, "taxjurcode", "")).strip().upper(),
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


def _is_windows_runtime() -> bool:
    return os.name == "nt" or platform.system().lower() == "windows"


def _is_wsl_runtime() -> bool:
    if _is_windows_runtime():
        return False

    release = platform.release().lower()
    version = platform.version().lower()
    return (
        "microsoft" in release
        or "wsl" in release
        or "microsoft" in version
        or "wsl" in version
        or bool(os.environ.get("WSL_INTEROP"))
    )


def _bridge_python_executable() -> str | None:
    configured_candidates = [
        os.getenv("SAP_FI_BRIDGE_PYTHON", "").strip(),
        os.getenv("WORKFLOW_PYTHON_EXEC", "").strip(),
    ]
    for candidate in configured_candidates:
        if not candidate:
            continue
        resolved = Path(candidate).expanduser()
        if resolved.exists():
            return str(resolved)
        found = shutil.which(candidate)
        if found:
            return found

    if not _is_windows_runtime():
        return None

    repo_root = Path(__file__).resolve().parents[1]
    windows_candidates = [
        repo_root / ".venv-rfc" / "Scripts" / "python.exe",
        repo_root / ".venv-rfc" / "Scripts" / "pythonw.exe",
    ]
    for candidate in windows_candidates:
        if candidate.exists():
            return str(candidate)

    return None


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


def _env_alias_default(environment: str, aliases: tuple[str, ...], default: str = "") -> str:
    env = _normalize_environment(environment)
    for alias in aliases:
        alias_key = str(alias or "").strip().upper()
        if not alias_key:
            continue
        for candidate in (f"SAP_{env}_{alias_key}", f"SAP_{alias_key}"):
            value = os.getenv(candidate, "").strip()
            if value:
                return value
    return default


def _payment_method_default(environment: str, branch: str, fallback: str = "") -> str:
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
    return _env_alias_default(environment, aliases_by_branch.get(branch_key, ()), fallback)


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

    withholding_tax_type = _env_default(
        environment,
        "withholding_tax_type",
        str(payload.get("withholding_tax_type") or ""),
    )
    withholding_tax_code = _env_default(
        environment,
        "withholding_tax_code",
        str(payload.get("withholding_tax_code") or ""),
    )
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

    accountreceivable = [
        {
            "ITEMNO_ACC": "1",
            "CUSTOMER": customer_account,
            "ITEM_TEXT": item_text,
        }
    ]
    payment_method = _payment_method_default(environment, "cliente", str(payload.get("payment_method") or "").strip().upper())
    if payment_method:
        accountreceivable[0]["PYMT_METH"] = payment_method
    if withholding_tax_code:
        accountreceivable[0]["W_TAX_CODE"] = withholding_tax_code
    accountgl = [
        {
            "ITEMNO_ACC": "2",
            "GL_ACCOUNT": revenue_gl_account,
            "ITEM_TEXT": item_text,
        }
    ]
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
        base_amount=_env_default(
            environment,
            "withholding_tax_base_amount",
            str(payload.get("withholding_tax_base_amount") or gross_amount),
        ),
        manual_amount=_env_default(
            environment,
            "withholding_tax_amount",
            str(payload.get("withholding_tax_amount") or ""),
        ),
    )

    currencyamount = [
        _base_currency_row(1, currency, gross_amount),
        _base_currency_row(2, currency, -net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, -tax_amount, amt_base=net_amount, tax_amt=-tax_amount))
    _validate_currencyamount_balance(currencyamount)

    result = {
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_cliente", "DR"),
        ),
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

    withholding_tax_type = _env_default(
        environment,
        "withholding_tax_type",
        str(payload.get("withholding_tax_type") or ""),
    )
    withholding_tax_code = _env_default(
        environment,
        "withholding_tax_code",
        str(payload.get("withholding_tax_code") or ""),
    )
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

    accountpayable = [
        {
            "ITEMNO_ACC": "1",
            "VENDOR_NO": vendor_account,
            "ITEM_TEXT": item_text,
        }
    ]
    payment_method = _payment_method_default(environment, "fornecedor", str(payload.get("payment_method") or "").strip().upper())
    if payment_method:
        accountpayable[0]["PYMT_METH"] = payment_method
    if withholding_tax_code:
        accountpayable[0]["W_TAX_CODE"] = withholding_tax_code
    accountgl = [
        {
            "ITEMNO_ACC": "2",
            "GL_ACCOUNT": expense_gl_account,
            "ITEM_TEXT": item_text,
        }
    ]
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
        base_amount=_env_default(
            environment,
            "withholding_tax_base_amount",
            str(payload.get("withholding_tax_base_amount") or net_amount),
        ),
        manual_amount=_env_default(
            environment,
            "withholding_tax_amount",
            str(payload.get("withholding_tax_amount") or ""),
        ),
    )

    currencyamount = [
        _base_currency_row(1, currency, -gross_amount),
        _base_currency_row(2, currency, net_amount),
    ]
    if tax_line:
        currencyamount.append(_base_currency_row(3, currency, tax_amount, amt_base=net_amount, tax_amt=tax_amount))
    _validate_currencyamount_balance(currencyamount)

    result = {
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_fornecedor", "KR"),
        ),
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
        "DOCUMENTHEADER": _build_header(
            payload,
            doc_type=_env_default(environment, "doc_type_razao", "SA"),
        ),
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


def _post_fi_document_core(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    if Connection is None:
        if os.getenv("SAP_FI_BRIDGE_ACTIVE") == "1":
            raise RuntimeError(
                "PyRFC indisponível no processo de bridge. Verifique se o Python do bridge "
                "tem SAP NetWeaver RFC SDK + pyrfc instalados."
            )
        return _run_post_fi_document_via_bridge(environment, branch, payload)

    _require_pyrfc()
    connection_params = fi_build_connection_params(environment)
    payload = fi_apply_default_payload(environment, branch, payload)

    connection = Connection(**connection_params)  # type: ignore[misc]
    try:
        bapi_payload = fi_build_bapi_payload(branch, environment, payload, connection=connection)
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


def post_fi_document(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    if Connection is None:
        return _run_post_fi_document_via_bridge(environment, branch, payload)
    return _post_fi_document_core(environment, branch, payload)
