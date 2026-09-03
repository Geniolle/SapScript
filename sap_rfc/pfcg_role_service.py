from __future__ import annotations

import os
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    build_connection_params_for,
    resolve_target_env,
    choose_best_text,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_eq,
    make_read_only_guard,
    normalize_spras,
    read_table,
    validate_role_name,
)

ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TEXTS")


def _error_result(role_name: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "role": role_name,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def analyze_pfcg_role_prd(role_name: str) -> dict[str, Any]:
    normalized_role = str(role_name or "").strip().upper()

    try:
        normalized_role = validate_role_name(role_name)
    except ValueError as exc:
        return _error_result(normalized_role, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
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
            define_rows = read_table(
                connection,
                guard,
                table_name="AGR_DEFINE",
                fields=["AGR_NAME"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=5,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_DEFINE_AUTHORIZATION_ERROR",
                    "Não foi possível determinar se a função existe: sem autorização para consultar AGR_DEFINE.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        if not define_rows:
            return {
                "ok": True,
                "status": "NAO_EXISTE",
                "role": normalized_role,
                "description": None,
                "language": None,
                "system": target_env,
                "client": params["client"],
            }

        description = None
        language = None
        warning = None
        try:
            text_rows_raw = read_table(
                connection,
                guard,
                table_name="AGR_TEXTS",
                fields=["AGR_NAME", "SPRAS", "TEXT"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=20,
            )
            text_rows = [
                {"AGR_NAME": row[0], "SPRAS": row[1], "TEXT": row[2]}
                for row in text_rows_raw
            ]
            best = choose_best_text(text_rows)
            if best and str(best.get("TEXT", "")).strip():
                description = str(best["TEXT"]).strip()
                language = normalize_spras(str(best.get("SPRAS", "") or ""))
        except Exception as exc:
            if is_authorization_error(exc):
                warning = "Sem autorização para consultar AGR_TEXTS."
            else:
                warning = classify_rfc_error(exc)[1]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "EXISTE",
            "role": normalized_role,
            "description": description,
            "language": language,
            "system": target_env,
            "client": params["client"],
        }
        if warning:
            payload["warning"] = warning
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
