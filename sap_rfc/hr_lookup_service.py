"""Read-only: dados de RH (PA0002 + PA0105) de um colaborador por numero de
pessoal (PERNR), para pre-preencher a criacao de utilizador SAP.

Os dados de RH (infotipos PA*) so existem no mestre de PRD neste sistema —
mesmo quando o utilizador SAP vai ser criado noutro ambiente (ex.: QAD), a
leitura de RH e sempre feita contra PRD.

O username SAP sugerido e derivado do PERNR com o prefixo "S" (convencao da
empresa), nunca lido de uma tabela de RH.
"""

from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_eq,
    make_read_only_guard,
    read_table,
)

ALLOWED_TABLES = ("PA0002", "PA0105")
HR_SOURCE_ENV = "PRD"
_EMAIL_SUBTYPE = "0010"

_PERNR_RE = re.compile(r"^\d{1,8}$")


def validate_pernr(raw: str) -> str:
    value = re.sub(r"\D", "", str(raw or ""))
    if not value:
        raise ValueError("Indique um numero de colaborador (PERNR) valido.")
    if not _PERNR_RE.match(value):
        raise ValueError("PERNR invalido: use apenas digitos (max. 8).")
    return value.zfill(8)


def derive_username(pernr: str) -> str:
    return f"S{pernr}"


def _error_result(pernr: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "pernr": pernr,
        "error_type": error_type,
        "message": message,
        "system": HR_SOURCE_ENV,
    }
    if details:
        payload["details"] = details
    return payload


def lookup_hr_data(raw_pernr: str) -> dict[str, Any]:
    try:
        pernr = validate_pernr(raw_pernr)
    except ValueError as exc:
        return _error_result(str(raw_pernr or "").strip(), "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(HR_SOURCE_ENV)
    except Exception as exc:
        return _error_result(pernr, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(pernr, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(pernr, error_type, message, details=format_exception(exc))

    try:
        try:
            personal_rows = read_table(
                connection, guard, table_name="PA0002",
                fields=["PERNR", "VORNA", "NACHN"],
                options=make_option_eq("PERNR", pernr), rowcount=1,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(pernr, "AUTHORIZATION_ERROR", "Sem autorizacao para ler dados pessoais de RH (PA0002).", details=format_exception(exc))
            error_type, message = classify_rfc_error(exc)
            return _error_result(pernr, error_type, message, details=format_exception(exc))

        if not personal_rows:
            return {
                "ok": False, "status": "NAO_ENCONTRADO", "pernr": pernr,
                "error_type": "NOT_FOUND",
                "message": f"Colaborador {pernr} nao encontrado nos dados de RH (PA0002) de {HR_SOURCE_ENV}.",
                "system": HR_SOURCE_ENV,
            }

        _, first_name, last_name = personal_rows[0]
        first_name = first_name.strip()
        last_name = last_name.strip()

        email = ""
        try:
            comm_rows = read_table(
                connection, guard, table_name="PA0105",
                fields=["PERNR", "SUBTY", "USRID_LONG"],
                options=make_option_eq("PERNR", pernr), rowcount=20,
            )
            for row in comm_rows:
                if row[1].strip() == _EMAIL_SUBTYPE and row[2].strip():
                    email = row[2].strip()
                    break
        except Exception as exc:
            if not is_authorization_error(exc):
                raise

        return {
            "ok": True,
            "status": "OK",
            "pernr": pernr,
            "username": derive_username(pernr),
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "system": HR_SOURCE_ENV,
        }
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
