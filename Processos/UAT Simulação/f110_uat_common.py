# -*- coding: utf-8 -*-
"""
Helpers comuns para a simulacao UAT da F110/RFF110S.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

try:
    from pyrfc import Connection
    HAS_PYRFC = True
except Exception as exc:  # pragma: no cover - depends on local SAP SDK
    HAS_PYRFC = False
    PYRFC_IMPORT_ERROR = exc

try:
    from sap_agent.config import SapConnectionConfig
except Exception:  # pragma: no cover - fallback when package import path differs
    SapConnectionConfig = None  # type: ignore[assignment]


# =============================================================================
# (1) ENVIRONMENT AND RFC CONNECTION
# =============================================================================

ROOT_DIR = Path(__file__).resolve().parents[2]
DEFAULT_SYSTEM_KEY = "QAD"


def load_project_dotenv() -> None:
    candidates = [
        ROOT_DIR / ".env",
        Path.cwd() / ".env",
    ]
    for env_path in candidates:
        if env_path.exists():
            load_dotenv(env_path, override=False)


def _first_env(*names: str, default: str = "", required: bool = True) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    if required:
        raise RuntimeError(f"Falta definir uma das variaveis de ambiente: {', '.join(names)}")
    return default


def normalize_system_key(system_key: str | None) -> str:
    value = str(system_key or "").strip().upper()
    return value or DEFAULT_SYSTEM_KEY


def candidate_env_keys(system_key: str) -> list[str]:
    normalized = normalize_system_key(system_key)
    mapping = {
        "QAD": ["QAD", "S4Q", "S4QCLNT100"],
        "S4Q": ["S4Q", "QAD", "S4QCLNT100"],
        "S4QCLNT100": ["S4QCLNT100", "S4Q", "QAD"],
    }
    keys = mapping.get(normalized, [normalized])
    if normalized not in keys:
        keys.insert(0, normalized)
    return keys


def resolve_connection_params(system_key: str | None = None) -> dict[str, str]:
    requested_key = normalize_system_key(system_key or os.getenv("SAP_SYSTEM") or os.getenv("WORKFLOW_SAP_KEY"))
    keys_to_try = candidate_env_keys(requested_key)

    generic_ready = all(
        os.getenv(name, "").strip()
        for name in ("SAP_USER", "SAP_PASSWD", "SAP_ASHOST", "SAP_SYSNR", "SAP_CLIENT")
    )
    if generic_ready and SapConnectionConfig is not None:
        cfg = SapConnectionConfig.from_env()
        return cfg.as_pyrfc_params()

    ashost = _first_env(*(f"SAP_ASHOST_{key}" for key in keys_to_try), "SAP_ASHOST")
    sysnr = _first_env(*(f"SAP_SYSNR_{key}" for key in keys_to_try), "SAP_SYSNR", default="00", required=False) or "00"
    client = _first_env(*(f"SAP_CLIENT_{key}" for key in keys_to_try), "SAP_CLIENT")
    user = _first_env(*(f"SAP_USER_{key}" for key in keys_to_try), "SAP_USER")
    passwd = _first_env(
        *(f"SAP_PASSWORD_{key}" for key in keys_to_try),
        *(f"SAP_PASSWD_{key}" for key in keys_to_try),
        "SAP_PASSWORD",
        "SAP_PASSWD",
    )
    lang = _first_env(
        *(f"SAP_LANGUAGE_{key}" for key in keys_to_try),
        "SAP_LANG",
        "SAP_LANGUAGE",
        default="PT",
        required=False,
    ) or "PT"

    return {
        "ashost": ashost,
        "sysnr": sysnr,
        "client": client,
        "user": user,
        "passwd": passwd,
        "lang": lang,
    }


def open_rfc_connection(system_key: str | None = None) -> tuple[Connection, str]:
    if not HAS_PYRFC:
        raise RuntimeError(f"A biblioteca pyrfc nao esta disponivel: {PYRFC_IMPORT_ERROR}")

    params = resolve_connection_params(system_key)
    conn = Connection(**params)
    return conn, params["user"]


# =============================================================================
# (2) RFC TABLE ACCESS
# =============================================================================

def parse_rfc_table_result(result: dict[str, Any]) -> list[dict[str, str]]:
    sap_fields = [entry["FIELDNAME"] for entry in result.get("FIELDS", [])]
    rows: list[dict[str, str]] = []
    for row in result.get("DATA", []):
        values = str(row.get("WA", "")).split("|")
        rows.append(
            {
                field: values[index].strip() if index < len(values) else ""
                for index, field in enumerate(sap_fields)
            }
        )
    return rows


def read_table(
    conn: Connection,
    table_name: str,
    fields: list[str],
    options: list[str] | None = None,
    rowcount: int = 10,
) -> list[dict[str, str]]:
    fields_payload = [{"FIELDNAME": field} for field in fields]
    options_payload = [{"TEXT": option} for option in (options or [])]
    result = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table_name,
        DELIMITER="|",
        FIELDS=fields_payload,
        OPTIONS=options_payload,
        ROWCOUNT=rowcount,
    )
    return parse_rfc_table_result(result)


def read_table_with_fallbacks(
    conn: Connection,
    table_name: str,
    field_sets: list[list[str]],
    options: list[str] | None = None,
    rowcount: int = 10,
) -> tuple[list[dict[str, str]], str]:
    last_error: Exception | None = None
    for fields in field_sets:
        try:
            return read_table(conn, table_name, fields, options=options, rowcount=rowcount), ",".join(fields)
        except Exception as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    return [], ""


def call_bapi_ap_acc_getopenitems(
    conn: Connection,
    company_code: str,
    vendor: str,
    keydate: str,
    noteditems: str = " ",
) -> tuple[list[dict[str, str]], dict[str, Any]]:
    result = conn.call(
        "BAPI_AP_ACC_GETOPENITEMS",
        COMPANYCODE=company_code,
        KEYDATE=keydate,
        NOTEDITEMS=noteditems,
        VENDOR=vendor,
    )
    return [dict(row) for row in result.get("LINEITEMS", [])], dict(result or {})


def zero_pad_if_numeric(value: str | None, size: int = 10) -> str:
    text = str(value or "").strip()
    if text.isdigit():
        return text.zfill(size)
    return text


def parse_yyyymmdd(value: str | None) -> str:
    text = str(value or "").strip()
    if len(text) != 8 or not text.isdigit():
        return ""
    return text
