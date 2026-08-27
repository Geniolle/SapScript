from __future__ import annotations

import os
from datetime import date
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    classify_import_error,
    classify_rfc_error,
    fetch_composite_members,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_in,
    make_read_only_guard,
    read_table,
    role_exists,
    validate_role_name,
)

ALLOWED_TABLES = ("AGR_DEFINE", "AGR_USERS", "AGR_AGRS")


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


def _classify_assignment_status(from_dat: str, to_dat: str, today: date) -> str:
    today_int = today.year * 10000 + today.month * 100 + today.day
    try:
        from_int = int(from_dat) if from_dat.strip().isdigit() else 0
    except ValueError:
        from_int = 0
    try:
        to_int = int(to_dat) if to_dat.strip().isdigit() else 99991231
    except ValueError:
        to_int = 99991231

    if today_int < from_int:
        return "FUTURO"
    if today_int > to_int:
        return "EXPIRADO"
    return "ATIVO"


def _fetch_users(connection: Any, guard: Any, roles: list[str]) -> list[tuple[str, str, str]]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_USERS",
        fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    unique: dict[tuple[str, str, str], None] = {}
    for _agr_name, uname, from_dat, to_dat in rows:
        uname = uname.strip()
        if not uname:
            continue
        unique.setdefault((uname, from_dat.strip(), to_dat.strip()), None)
    return list(unique.keys())


def analyze_pfcg_role_users_prd(role_name: str) -> dict[str, Any]:
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
                    "users": [],
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
            raw_users = _fetch_users(connection, guard, roles_to_scan)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_USERS_AUTHORIZATION_ERROR",
                    "Sem autorização para consultar AGR_USERS.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_USERS_{error_type}", message, details=format_exception(exc))

        today = date.today()
        users = [
            {
                "username": uname,
                "valid_from": from_dat,
                "valid_to": to_dat,
                "assignment_status": _classify_assignment_status(from_dat, to_dat, today),
            }
            for uname, from_dat, to_dat in sorted(raw_users, key=lambda item: item[0])
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "role": normalized_role,
            "count": len(users),
            "users": users,
            "system": SYSTEM_NAME,
            "client": params["client"],
            "is_composite": is_composite,
        }
        if is_composite:
            payload["composite_members"] = composite_members
        if composite_warning:
            payload["warning"] = composite_warning
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
