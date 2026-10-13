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


import re

_USER_RE = re.compile(r"^[A-Za-z0-9_.\-]{1,12}$")


def _validate_username_or_400(username: str) -> str:
    value = str(username or "").strip().upper()
    if not value:
        raise HTTPException(status_code=400, detail="Indique um utilizador SAP.")
    if not _USER_RE.match(value):
        raise HTTPException(
            status_code=400,
            detail="Utilizador invalido: use apenas letras, numeros, _, . ou - (max. 12).",
        )
    return value


def _safe_user_create_failed_message() -> str:
    return "Não foi possível concluir a criação do utilizador."


def _safe_user_password_change_failed_message() -> str:
    return "Não foi possível concluir a alteração de senha do utilizador."


def _safe_user_unlock_failed_message() -> str:
    return "Não foi possível concluir o desbloqueio do utilizador."


def _validate_pernr_or_400(pernr: str) -> str:
    value = str(pernr or "").strip()
    if not value:
        raise HTTPException(status_code=400, detail="Indique o número de colaborador (PERNR).")
    digits = re.sub(r"\D", "", value)
    if not digits:
        raise HTTPException(status_code=400, detail="PERNR inválido: use apenas dígitos numéricos.")
    return digits.zfill(8)


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
        for field in (
            "uflag",
            "failed_logon_count",
            "locked",
            "login_failed_locked",
            "lock_reasons",
            "lock_status",
            "lock_check_error_type",
            "lock_check_message",
            "unlocked",
            "unlock_message",
            "unlock_sap_return_messages",
        ):
            if field in result:
                safe_result[field] = result.get(field)
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
        "uflag",
        "failed_logon_count",
        "locked",
        "login_failed_locked",
        "lock_reasons",
        "lock_status",
        "lock_check_error_type",
        "lock_check_message",
        "unlocked",
        "unlock_message",
        "unlock_sap_return_messages",
        "message",
    ):
        if field in result:
            safe_result[field] = result.get(field)
    return safe_result
