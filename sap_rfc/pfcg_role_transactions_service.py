from __future__ import annotations

import os
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    choose_best_text,
    classify_import_error,
    classify_rfc_error,
    fetch_composite_members,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_eq,
    make_option_in,
    make_read_only_guard,
    read_table,
    role_exists,
    validate_role_name,
)

# AGR_DEFINE é necessário apenas para confirmar a existência da função antes de
# consultar as restantes tabelas (evita disparar consultas adicionais para uma
# função inexistente).
ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TCODES", "AGR_AGRS", "TSTCT")


def _error_result(role_name: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "role": role_name,
        "error_type": error_type,
        "message": message,
        "system": SYSTEM_NAME,
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _fetch_tcodes(connection: Any, guard: Any, roles: list[str]) -> set[str]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_TCODES",
        fields=["AGR_NAME", "TCODE"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    return {row[1].strip() for row in rows if row[1].strip()}


def _fetch_tcode_descriptions(connection: Any, guard: Any, tcodes: list[str]) -> dict[str, str]:
    if not tcodes:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="TSTCT",
        fields=["SPRSL", "TCODE", "TTEXT"],
        options=make_option_in("TCODE", tcodes),
        rowcount=0,
    )
    by_tcode: dict[str, list[dict[str, str]]] = {}
    for spras, tcode, ttext in rows:
        by_tcode.setdefault(tcode, []).append({"SPRSL": spras, "TTEXT": ttext})

    descriptions: dict[str, str] = {}
    for tcode, candidates in by_tcode.items():
        best = choose_best_text(candidates)
        if best and str(best.get("TTEXT", "")).strip():
            descriptions[tcode] = str(best["TTEXT"]).strip()
    return descriptions


def analyze_pfcg_role_transactions_prd(role_name: str) -> dict[str, Any]:
    normalized_role = str(role_name or "").strip().upper()

    try:
        normalized_role = validate_role_name(role_name)
    except ValueError as exc:
        return _error_result(normalized_role, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params()
    except Exception as exc:
        return _error_result(normalized_role, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(normalized_role, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(normalized_role, error_type, message, details=format_exception(exc))

    try:
        try:
            if not role_exists(connection, guard, normalized_role):
                return {
                    "ok": True,
                    "status": "NAO_EXISTE",
                    "role": normalized_role,
                    "count": 0,
                    "transactions": [],
                    "system": SYSTEM_NAME,
                    "client": params["client"],
                }
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_DEFINE_AUTHORIZATION_ERROR",
                    "Sem autorização para consultar AGR_DEFINE.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        composite_members: list[str] = []
        composite_warning = None
        try:
            composite_members = fetch_composite_members(connection, guard, normalized_role)
        except Exception as exc:
            composite_warning = (
                "Sem autorização para consultar AGR_AGRS."
                if is_authorization_error(exc)
                else f"Não foi possível verificar se a função é composta: {classify_rfc_error(exc)[1]}"
            )

        is_composite = bool(composite_members)
        roles_to_scan = [normalized_role, *composite_members] if is_composite else [normalized_role]

        try:
            tcodes = _fetch_tcodes(connection, guard, roles_to_scan)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_TCODES_AUTHORIZATION_ERROR",
                    "Sem autorização para consultar AGR_TCODES.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_TCODES_{error_type}", message, details=format_exception(exc))

        sorted_tcodes = sorted(tcodes)

        descriptions: dict[str, str] = {}
        description_warning = None
        try:
            descriptions = _fetch_tcode_descriptions(connection, guard, sorted_tcodes)
        except Exception as exc:
            description_warning = (
                "Sem autorização para consultar TSTCT."
                if is_authorization_error(exc)
                else f"Não foi possível obter descrições das transações: {classify_rfc_error(exc)[1]}"
            )

        transactions = [
            {"tcode": tcode, "description": descriptions.get(tcode)}
            for tcode in sorted_tcodes
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "role": normalized_role,
            "count": len(transactions),
            "transactions": transactions,
            "system": SYSTEM_NAME,
            "client": params["client"],
            "is_composite": is_composite,
        }
        if is_composite:
            payload["composite_members"] = composite_members
        warnings = [w for w in (composite_warning, description_warning) if w]
        if warnings:
            payload["warning"] = " ".join(warnings)
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
