"""Estado em memoria e sanitizacao partilhados pelas rotas de criacao de utilizador
(/api/salsa-it-agent/user/create/*). Mesmo padrao usado em pfcg_common.py."""
from __future__ import annotations

from typing import Any

from fastapi import HTTPException

# Estado do fluxo preview -> confirm (chave = job_id do preview). Guarda também a
# senha inicial validada (nunca devolvida ao frontend) para o passo de confirmação
# não precisar de a receber outra vez do browser.
USER_CREATE_PREVIEWS: dict[str, dict[str, Any]] = {}

USER_CREATE_SYSTEMS = ("DEV", "QAD", "PRD")


def _validate_user_create_system_or_400(system: str) -> str:
    value = str(system or "").strip().upper()
    if value not in USER_CREATE_SYSTEMS:
        raise HTTPException(
            status_code=400,
            detail=f"Sistema inválido: {value}. Use {', '.join(USER_CREATE_SYSTEMS)}.",
        )
    return value


def _validate_username_or_400(username: str) -> str:
    try:
        from sap_rfc.user_data_service import validate_username
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail="Não foi possível carregar a validação de utilizador no backend.",
        ) from exc

    try:
        return validate_username(username)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def _safe_user_create_failed_message() -> str:
    return "Não foi possível concluir a criação do utilizador."


def _validate_pernr_or_400(pernr: str) -> str:
    try:
        from sap_rfc.hr_lookup_service import validate_pernr
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail="Não foi possível carregar a validação de PERNR no backend.",
        ) from exc

    try:
        return validate_pernr(pernr)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def _safe_hr_lookup_failed_message() -> str:
    return "Não foi possível consultar os dados de RH."


def _safe_hr_lookup_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "pernr": result.get("pernr"),
        "system": result.get("system"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        return safe_result

    for field in ("username", "first_name", "last_name", "email"):
        if field in result:
            safe_result[field] = result.get(field)
    return safe_result


def _safe_user_create_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "environment": result.get("environment"),
        "username": result.get("username"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        if result.get("missing_roles"):
            safe_result["missing_roles"] = result.get("missing_roles")
        if result.get("sap_return_messages"):
            safe_result["sap_return_messages"] = result.get("sap_return_messages")
        return safe_result

    # Campos apenas do fluxo de sucesso (preview e/ou criação real). A senha nunca
    # é incluída aqui — só uma indicação da origem (Excel/pedido vs. .env).
    for field in (
        "first_name",
        "last_name",
        "email",
        "ustyp",
        "group",
        "valid_from",
        "valid_to",
        "department",
        "function",
        "roles",
        "roles_requested",
        "roles_assigned",
        "roles_missing",
        "role_assignment_messages",
        "password_source",
        "message",
    ):
        if field in result:
            safe_result[field] = result.get(field)
    return safe_result
