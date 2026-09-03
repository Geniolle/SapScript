"""Read-only: dado um objeto de autorizacao, devolve as funcoes PFCG (roles) que
o contem em SAP PRD.

Analogo a `pfcg_transaction_roles_service` mas sobre `AGR_1251` (valores de
objetos de autorizacao por funcao), filtrando por `OBJECT`. Enriquece com:
  - `AGR_TEXTS` -> descricao das funcoes;
  - `AGR_AGRS`  -> funcoes compostas que as incluem;
  - `TOBJT`     -> texto do objeto de autorizacao.
"""

from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    choose_best_text,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_eq,
    make_option_in,
    make_read_only_guard,
    read_table,
)

ALLOWED_TABLES = ("AGR_1251", "AGR_TEXTS", "AGR_AGRS", "TOBJT")

_AUTHOBJ_RE = re.compile(r"^[A-Za-z0-9_/]{1,40}$")


def validate_auth_object(raw: str) -> str:
    value = str(raw or "").strip().upper()
    if not value:
        raise ValueError("Indique um objeto de autorizacao.")
    if not _AUTHOBJ_RE.match(value):
        raise ValueError("Objeto de autorizacao invalido: use apenas letras, numeros, _ ou /.")
    return value


def _error_result(auth_object: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "auth_object": auth_object,
        "error_type": error_type,
        "message": message,
        "system": SYSTEM_NAME,
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _is_custom_role(name: str) -> bool:
    """Apenas funcoes personalizadas (comecam por Z)."""
    return name.strip().upper().startswith("Z")


def _fetch_roles_with_object(connection: Any, guard: Any, auth_object: str) -> list[str]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_1251",
        fields=["AGR_NAME", "OBJECT", "DELETED"],
        options=make_option_eq("OBJECT", auth_object),
        rowcount=0,
    )
    roles: set[str] = set()
    for agr_name, _obj, deleted in rows:
        if str(deleted).strip().upper() == "X":
            continue
        name = agr_name.strip()
        if name and _is_custom_role(name):
            roles.add(name)
    return sorted(roles)


def _fetch_role_descriptions(connection: Any, guard: Any, roles: list[str]) -> dict[str, str]:
    if not roles:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="AGR_TEXTS",
        fields=["AGR_NAME", "SPRAS", "TEXT"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    by_role: dict[str, list[dict[str, str]]] = {}
    for agr_name, spras, text in rows:
        by_role.setdefault(agr_name.strip(), []).append({"SPRAS": spras, "TEXT": text})
    out: dict[str, str] = {}
    for role, candidates in by_role.items():
        best = choose_best_text(candidates)
        if best and str(best.get("TEXT", "")).strip():
            out[role] = str(best["TEXT"]).strip()
    return out


def _fetch_parent_composites(connection: Any, guard: Any, child_roles: list[str]) -> dict[str, list[str]]:
    if not child_roles:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="AGR_AGRS",
        fields=["AGR_NAME", "CHILD_AGR"],
        options=make_option_in("CHILD_AGR", child_roles),
        rowcount=0,
    )
    out: dict[str, list[str]] = {}
    for parent, child in rows:
        parent, child = parent.strip(), child.strip()
        if parent and child and _is_custom_role(parent):
            out.setdefault(child, []).append(parent)
    return {k: sorted(set(v)) for k, v in out.items()}


def _fetch_object_text(connection: Any, guard: Any, auth_object: str) -> str | None:
    rows = read_table(
        connection,
        guard,
        table_name="TOBJT",
        fields=["LANGU", "OBJCT", "TTEXT"],
        options=make_option_eq("OBJCT", auth_object),
        rowcount=20,
    )
    candidates = [{"SPRSL": langu, "TTEXT": ttext} for langu, _obj, ttext in rows]
    best = choose_best_text(candidates)
    if best and str(best.get("TTEXT", "")).strip():
        return str(best["TTEXT"]).strip()
    return None


def analyze_object_roles_prd(auth_object: str) -> dict[str, Any]:
    try:
        norm_obj = validate_auth_object(auth_object)
    except ValueError as exc:
        return _error_result(str(auth_object or "").strip().upper(), "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params()
    except Exception as exc:
        return _error_result(norm_obj, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_obj, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_obj, error_type, message, details=format_exception(exc))

    try:
        try:
            roles = _fetch_roles_with_object(connection, guard, norm_obj)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    norm_obj,
                    "AGR_1251_AUTHORIZATION_ERROR",
                    "Sem autorizacao para consultar AGR_1251.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_obj, f"AGR_1251_{error_type}", message, details=format_exception(exc))

        object_text = None
        warnings: list[str] = []
        try:
            object_text = _fetch_object_text(connection, guard, norm_obj)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar TOBJT."
                if is_authorization_error(exc)
                else f"Sem texto do objeto: {classify_rfc_error(exc)[1]}"
            )

        if not roles:
            return {
                "ok": True,
                "status": "NAO_ENCONTRADO",
                "auth_object": norm_obj,
                "auth_object_text": object_text,
                "count": 0,
                "roles": [],
                "system": SYSTEM_NAME,
                "client": params["client"],
                **({"warning": " ".join(warnings)} if warnings else {}),
            }

        descriptions: dict[str, str] = {}
        try:
            descriptions = _fetch_role_descriptions(connection, guard, roles)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar AGR_TEXTS."
                if is_authorization_error(exc)
                else f"Sem descricoes das funcoes: {classify_rfc_error(exc)[1]}"
            )

        parents: dict[str, list[str]] = {}
        try:
            parents = _fetch_parent_composites(connection, guard, roles)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar AGR_AGRS."
                if is_authorization_error(exc)
                else f"Sem funcoes compostas: {classify_rfc_error(exc)[1]}"
            )

        role_rows = [
            {
                "role": role,
                "description": descriptions.get(role),
                "composite_parents": parents.get(role, []),
            }
            for role in roles
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "auth_object": norm_obj,
            "auth_object_text": object_text,
            "count": len(role_rows),
            "roles": role_rows,
            "system": SYSTEM_NAME,
            "client": params["client"],
        }
        if warnings:
            payload["warning"] = " ".join(warnings)
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
