"""Read-only: dada uma lista de funções PFCG e um campo de autorização (ex.:
ACTVT), indica quais das funções têm esse campo atribuído e em que objetos de
autorização, em SAP PRD.

Nota: ACTVT ("Atividade") não é um objeto de autorização — é um FIELD que
aparece dentro de muitos objetos de autorização diferentes (S_TCODE não tem,
mas a maioria dos objetos de FI/MM/PP tem). Esta análise verifica a presença
do campo em qualquer objeto da função, não de um objeto específico.

Análogo a `pfcg_role_auth_objects_service` mas agrega múltiplas funções e
filtra por FIELD em vez de listar todos os objetos.
"""

from __future__ import annotations

import os
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params_for,
    resolve_target_env,
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

ALLOWED_TABLES = ("AGR_DEFINE", "AGR_1251", "AGR_AGRS")


def _error_result(role_names: list[str], field_name: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "roles": role_names,
        "field": field_name,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _fetch_field_matches(connection: Any, guard: Any, roles: list[str], field_name: str) -> list[dict[str, str]]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_1251",
        fields=["AGR_NAME", "OBJECT", "FIELD", "LOW", "HIGH", "DELETED"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    matches: list[dict[str, str]] = []
    for agr_name, obj, field, low, high, deleted in rows:
        if field.strip().upper() != field_name:
            continue
        if deleted.strip().upper() == "X":
            continue
        matches.append({"role": agr_name.strip(), "object": obj.strip(), "low": low.strip(), "high": high.strip()})
    return matches


def analyze_pfcg_roles_field_prd(role_names: list[str], field_name: str) -> dict[str, Any]:
    raw_roles = [str(name or "").strip().upper() for name in role_names if str(name or "").strip()]
    field = str(field_name or "").strip().upper()

    if not raw_roles:
        return _error_result(raw_roles, field, "INVALID_INPUT", "Indique pelo menos uma função PFCG.")
    if not field:
        return _error_result(raw_roles, field, "INVALID_INPUT", "Indique o campo de autorização a pesquisar (ex.: ACTVT).")

    normalized_roles: list[str] = []
    for raw in raw_roles:
        try:
            normalized_roles.append(validate_role_name(raw))
        except ValueError as exc:
            return _error_result(raw_roles, field, "INVALID_INPUT", f"Função inválida '{raw}': {exc}")

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
    except Exception as exc:
        return _error_result(normalized_roles, field, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(normalized_roles, field, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(normalized_roles, field, error_type, message, details=format_exception(exc))

    try:
        results: dict[str, dict[str, Any]] = {}
        warnings: list[str] = []

        for role in normalized_roles:
            try:
                if not role_exists(connection, guard, role):
                    results[role] = {"exists": False}
                    continue
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(
                        normalized_roles, field,
                        "AGR_DEFINE_AUTHORIZATION_ERROR",
                        "Sem autorização para consultar AGR_DEFINE.",
                        details=format_exception(exc),
                    )
                error_type, message = classify_rfc_error(exc)
                return _error_result(normalized_roles, field, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

            composite_members: list[str] = []
            try:
                composite_members = fetch_composite_members(connection, guard, role)
            except Exception as exc:
                warnings.append(
                    f"{role}: sem autorização para consultar AGR_AGRS."
                    if is_authorization_error(exc)
                    else f"{role}: não foi possível verificar se é composta ({classify_rfc_error(exc)[1]})."
                )

            roles_to_scan = [role, *composite_members] if composite_members else [role]

            try:
                matches = _fetch_field_matches(connection, guard, roles_to_scan, field)
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(
                        normalized_roles, field,
                        "AGR_1251_AUTHORIZATION_ERROR",
                        "Sem autorização para consultar AGR_1251.",
                        details=format_exception(exc),
                    )
                error_type, message = classify_rfc_error(exc)
                return _error_result(normalized_roles, field, f"AGR_1251_{error_type}", message, details=format_exception(exc))

            objects_with_field = sorted({m["object"] for m in matches})
            results[role] = {
                "exists": True,
                "is_composite": bool(composite_members),
                "has_field": bool(matches),
                "objects_with_field": objects_with_field,
                "detail": matches,
            }

        with_field = [role for role, data in results.items() if data.get("has_field")]
        without_field = [role for role, data in results.items() if data.get("exists") and not data.get("has_field")]
        not_found = [role for role, data in results.items() if not data.get("exists")]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "roles": normalized_roles,
            "field": field,
            "with_field": with_field,
            "without_field": without_field,
            "not_found": not_found,
            "results": results,
            "system": target_env,
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
