"""Read-only: dada uma função PFCG, devolve os objetos de autorização (ex.:
S_TCODE, Z_BASIS_BASE) que a função contém, em SAP PRD.

Analogo a `pfcg_role_transactions_service` mas sobre `AGR_1251` (valores de
objetos de autorização por função), filtrando por `AGR_NAME`. Enriquece com
`TOBJT` -> texto do objeto de autorização. Se a função for composta
(Sammelrolle), agrega também os objetos das funções-membro.
"""

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

ALLOWED_TABLES = ("AGR_DEFINE", "AGR_1251", "AGR_AGRS", "TOBJT")


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


def _fetch_auth_objects(connection: Any, guard: Any, roles: list[str]) -> set[str]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_1251",
        fields=["AGR_NAME", "OBJECT", "DELETED"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    objects: set[str] = set()
    for _agr_name, obj, deleted in rows:
        if str(deleted).strip().upper() == "X":
            continue
        obj = obj.strip()
        if obj:
            objects.add(obj)
    return objects


def _fetch_object_descriptions(connection: Any, guard: Any, auth_objects: list[str]) -> dict[str, str]:
    if not auth_objects:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="TOBJT",
        # Campo real da TOBJT é "OBJECT", não "OBJCT" (confirmado via
        # RFC_READ_TABLE com FIELDS=[] em PRD).
        fields=["LANGU", "OBJECT", "TTEXT"],
        options=make_option_in("OBJECT", auth_objects),
        rowcount=0,
    )
    by_object: dict[str, list[dict[str, str]]] = {}
    for langu, objct, ttext in rows:
        by_object.setdefault(objct.strip(), []).append({"SPRSL": langu, "TTEXT": ttext})

    descriptions: dict[str, str] = {}
    for obj, candidates in by_object.items():
        best = choose_best_text(candidates)
        if best and str(best.get("TTEXT", "")).strip():
            descriptions[obj] = str(best["TTEXT"]).strip()
    return descriptions


def analyze_pfcg_role_auth_objects_prd(role_name: str) -> dict[str, Any]:
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
            if not role_exists(connection, guard, normalized_role):
                return {
                    "ok": True,
                    "status": "NAO_EXISTE",
                    "role": normalized_role,
                    "count": 0,
                    "auth_objects": [],
                    "system": target_env,
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
            auth_objects = _fetch_auth_objects(connection, guard, roles_to_scan)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_1251_AUTHORIZATION_ERROR",
                    "Sem autorização para consultar AGR_1251.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_1251_{error_type}", message, details=format_exception(exc))

        sorted_objects = sorted(auth_objects)

        descriptions: dict[str, str] = {}
        description_warning = None
        try:
            descriptions = _fetch_object_descriptions(connection, guard, sorted_objects)
        except Exception as exc:
            description_warning = (
                "Sem autorização para consultar TOBJT."
                if is_authorization_error(exc)
                else f"Não foi possível obter descrições dos objetos: {classify_rfc_error(exc)[1]}"
            )

        auth_object_rows = [
            {"auth_object": obj, "description": descriptions.get(obj)}
            for obj in sorted_objects
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "role": normalized_role,
            "count": len(auth_object_rows),
            "auth_objects": auth_object_rows,
            "system": target_env,
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
