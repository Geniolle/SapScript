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
# PA0105/0010 pode ter mais do que uma linha (email pessoal e email da empresa).
# Preferir sempre o email do dominio da empresa quando existir mais do que um.
_COMPANY_EMAIL_DOMAIN = "@SALSAJEANS.COM"

_PERNR_RE = re.compile(r"^\d{1,8}$")


def validate_pernr(raw: str) -> str:
    value = re.sub(r"\D", "", str(raw or ""))
    if not value:
        raise ValueError("Indique um numero de colaborador (PERNR) valido.")
    if not _PERNR_RE.match(value):
        raise ValueError("PERNR invalido: use apenas digitos (max. 8).")
    return value.zfill(8)


def derive_username(pernr: str) -> str:
    # O username SAP sugerido não deve ter zeros à esquerda (ex.: S2760 e não S00002760)
    clean_digits = str(pernr or "").strip().lstrip("0")
    return f"S{clean_digits}" if clean_digits else "S0"


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
            candidate_emails = [
                row[2].strip() for row in comm_rows
                if row[1].strip() == _EMAIL_SUBTYPE and row[2].strip()
            ]
            company_email = next(
                (addr for addr in candidate_emails if addr.upper().endswith(_COMPANY_EMAIL_DOMAIN)),
                None,
            )
            email = company_email or (candidate_emails[0] if candidate_emails else "")
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


_NAME_RE = re.compile(r"^[A-Za-zÀ-ÿ' \-]{2,60}$")
_MAX_NAME_MATCHES = 10


def _case_variants(value: str) -> list[str]:
    variants: list[str] = []
    for v in (value, value.upper(), value.title()):
        if v not in variants:
            variants.append(v)
    return variants


def _exact_options_name_hr(query: str) -> list[dict[str, str]]:
    """Primeiro e último nome com igualdade exata em PA0002 (evita LIKE '%...%', lento em RFC_READ_TABLE).

    Um único termo pesquisa em VORNA (nome) ou NACHN (apelido), com igualdade exata.
    """
    parts = query.split()
    last = parts[-1]
    first = parts[0] if len(parts) > 1 else None
    clauses: list[str] = []
    if first is None:
        for v in _case_variants(last):
            clauses.append(f"NACHN = '{v}'")
            clauses.append(f"VORNA = '{v}'")
    else:
        for last_v in _case_variants(last):
            for first_v in _case_variants(first):
                clauses.append(f"VORNA = '{first_v}' AND NACHN = '{last_v}'")
    rows: list[dict[str, str]] = []
    for clause in dict.fromkeys(clauses):
        rows.append({"TEXT": f"{'OR ' if rows else ''}{clause}"})
    return rows


def lookup_hr_by_name(raw_name: str) -> dict[str, Any]:
    """Pesquisa por nome em PA0002 (PRD). Um único resultado segue para lookup_hr_data; vários são listados."""
    name = " ".join(str(raw_name or "").split())
    if not _NAME_RE.match(name):
        return _error_result("", "INVALID_INPUT", "Indique nome ou apelido válido (letras, espaços, hífen ou apóstrofo).")

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(HR_SOURCE_ENV)
    except Exception as exc:
        return _error_result("", "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result("", error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
        try:
            rows = read_table(
                connection, guard, table_name="PA0002",
                fields=["PERNR", "VORNA", "NACHN"],
                options=_exact_options_name_hr(name), rowcount=50,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result("", "AUTHORIZATION_ERROR", "Sem autorizacao para ler dados pessoais de RH (PA0002).", details=format_exception(exc))
            error_type, message = classify_rfc_error(exc)
            return _error_result("", error_type, message, details=format_exception(exc))
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result("", error_type, message, details=format_exception(exc))
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    matches: dict[str, dict[str, str]] = {}
    for pernr_raw, first, last in rows:
        pernr_value = pernr_raw.strip()
        if pernr_value and pernr_value not in matches:
            matches[pernr_value] = {"pernr": pernr_value.zfill(8), "name": f"{first.strip()} {last.strip()}".strip()}

    if not matches:
        return {
            "ok": False, "status": "NAO_ENCONTRADO", "pernr": None,
            "error_type": "NOT_FOUND",
            "message": f"Nenhum colaborador com o nome '{name}' nos dados de RH (PA0002) de {HR_SOURCE_ENV}.",
            "system": HR_SOURCE_ENV,
        }

    if len(matches) == 1:
        return lookup_hr_data(next(iter(matches.values()))["pernr"])

    listed = list(matches.values())[:_MAX_NAME_MATCHES]
    names = "; ".join(f"{m['pernr']} {m['name']}" for m in listed)
    return {
        "ok": False, "status": "VARIAS_CORRESPONDENCIAS", "pernr": None,
        "error_type": "AMBIGUOUS",
        "message": f"Vários colaboradores encontrados: {names}. Indique o número de colaborador (ex.: S5887).",
        "matches": listed,
        "system": HR_SOURCE_ENV,
    }
