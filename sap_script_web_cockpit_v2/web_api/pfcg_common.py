"""Estado em memoria e sanitizacao partilhados pelas rotas PFCG
(/api/salsa-it-agent/pfcg/* e /api/pfcg/*). Extraido de main.py na Fase 3."""
from __future__ import annotations

from typing import Any

from fastapi import HTTPException


# Estado dos fluxos preview -> confirm (chave = job_id do preview).
PFCG_EXCEL_SELECTIONS: dict[str, dict[str, str]] = {}
PFCG_RFC_CREATE_PREVIEWS: dict[str, dict[str, Any]] = {}
PFCG_COMPOSTA_CREATE_PREVIEWS: dict[str, dict[str, Any]] = {}

PFCG_RFC_CREATE_ENVIRONMENT = "DEV"
PFCG_RFC_DELETE_PREVIEWS: dict[str, dict[str, Any]] = {}
PFCG_RFC_DELETE_ENVIRONMENT = "DEV"
PFCG_RFC_BULK_DELETE_PREVIEWS: dict[str, dict[str, Any]] = {}


def _validate_pfcg_role_name_or_400(role_name: str) -> str:
    try:
        from sap_rfc import validate_role_name
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail="Não foi possível carregar a validação PFCG no backend.",
        ) from exc

    try:
        return validate_role_name(role_name)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


PFCG_SYSTEMS = ("DEV", "QAD", "PRD", "CUA")


def _validate_pfcg_system_or_400(system: str) -> str:
    value = str(system or "PRD").strip().upper() or "PRD"
    if value not in PFCG_SYSTEMS:
        raise HTTPException(
            status_code=400,
            detail=f"Sistema invalido: {value}. Use DEV, QAD, PRD ou CUA.",
        )
    return value


def _safe_pfcg_failed_message() -> str:
    return "Não foi possível concluir a análise PFCG."


def _safe_pfcg_sub_result(result: dict[str, Any], *, items_key: str, item_fields: tuple[str, ...]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "role": str(result.get("role") or ""),
        "count": result.get("count"),
        "system": result.get("system"),
        "client": result.get("client"),
        "is_composite": bool(result.get("is_composite")),
    }
    if safe_result["is_composite"]:
        composite_members = result.get("composite_members")
        safe_result["composite_members"] = composite_members if isinstance(composite_members, list) else []
    if result.get("warning"):
        safe_result["warning"] = result.get("warning")

    raw_items = result.get(items_key)
    if isinstance(raw_items, list):
        safe_result[items_key] = [
            {field: item.get(field) for field in item_fields}
            for item in raw_items
            if isinstance(item, dict)
        ]
    else:
        safe_result[items_key] = []

    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")

    return safe_result


def _safe_pfcg_role_search_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "pattern": str(result.get("pattern") or ""),
        "count": result.get("count"),
        "system": result.get("system"),
        "client": result.get("client"),
    }
    if result.get("warning"):
        safe_result["warning"] = result.get("warning")

    raw_roles = result.get("roles")
    safe_result["roles"] = [
        {"role": str(item.get("role") or ""), "description": item.get("description")}
        for item in raw_roles
        if isinstance(item, dict)
    ] if isinstance(raw_roles, list) else []

    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")

    return safe_result


def _safe_pfcg_rfc_delete_result(result: dict[str, Any]) -> dict[str, Any]:
    return {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or "ERROR"),
        "environment": str(result.get("environment") or PFCG_RFC_DELETE_ENVIRONMENT),
        "role": str(result.get("role") or ""),
        "description": result.get("description"),
        "tcodes": result.get("tcodes") or [],
        "tcodes_count": result.get("tcodes_count"),
        "users_count": result.get("users_count"),
        "transport": result.get("transport"),
        "transport_mode": result.get("transport_mode"),
        "transport_request": result.get("transport_request"),
        "error_type": result.get("error_type"),
        "message": result.get("message"),
    }


def _safe_pfcg_rfc_bulk_delete_result(result: dict[str, Any]) -> dict[str, Any]:
    raw_items = result.get("items")
    items = [
        {
            "role": str(item.get("role") or ""),
            "ok": bool(item.get("ok")),
            "status": str(item.get("status") or ""),
            "description": item.get("description"),
            "users_count": item.get("users_count"),
            "message": item.get("message"),
        }
        for item in raw_items
        if isinstance(item, dict)
    ] if isinstance(raw_items, list) else []

    return {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or "ERROR"),
        "environment": str(result.get("environment") or PFCG_RFC_DELETE_ENVIRONMENT),
        "roles": result.get("roles") or [],
        "items": items,
        "found_count": result.get("found_count"),
        "not_found_count": result.get("not_found_count"),
        "deleted_count": result.get("deleted_count"),
        "failed_count": result.get("failed_count"),
        "transport": result.get("transport"),
        "transport_mode": result.get("transport_mode"),
        "transport_request": result.get("transport_request"),
        "error_type": result.get("error_type"),
        "message": result.get("message"),
    }


def _safe_pfcg_rfc_create_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "environment": result.get("environment"),
        "role": result.get("role"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        if result.get("missing_tcodes"):
            safe_result["missing_tcodes"] = result.get("missing_tcodes")
        return safe_result

    # Campos apenas do fluxo de sucesso (preview e/ou criação real)
    for field in (
        "description",
        "tcodes",
        "tcodes_count",
        "tcodes_requested",
        "tcodes_created",
        "profile_generated",
        "transport",
        "transport_mode",
        "transport_request",
        "transport_request_created",
    ):
        if field in result:
            safe_result[field] = result.get(field)
    return safe_result


def _safe_pfcg_transport_search_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "environment": result.get("environment"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        return safe_result

    safe_result["owner"] = result.get("owner")
    safe_result["requests_count"] = result.get("requests_count")
    safe_result["requests"] = [
        {
            "request": row.get("request"),
            "description": row.get("description"),
            "trtype": row.get("trtype"),
            "target_system": row.get("target_system"),
            "state": row.get("state"),
        }
        for row in (result.get("requests") or [])
        if isinstance(row, dict)
    ]
    return safe_result
