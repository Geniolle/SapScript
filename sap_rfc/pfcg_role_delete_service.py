from __future__ import annotations

import os
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    KNOWN_ENVIRONMENTS,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    load_project_env,
    make_option_eq,
    make_write_guard,
    read_table,
    role_exists,
    validate_role_name,
)
from sap_rfc.pfcg_transport_service import (
    create_transport_request,
    validate_request_description,
    validate_transport_request,
)

MAX_ROLE_NAME_LENGTH = 30

ALLOWED_DELETE_ENVIRONMENTS = KNOWN_ENVIRONMENTS  # escrita permitida em DEV/QAD/PRD/CUA (autorizado explicitamente)

TRANSPORT_MODE_LOCAL = "LOCAL"
TRANSPORT_MODE_CREATE_REQUEST = "CREATE_REQUEST"
TRANSPORT_MODE_EXISTING_REQUEST = "EXISTING_REQUEST"
VALID_TRANSPORT_MODES = (TRANSPORT_MODE_LOCAL, TRANSPORT_MODE_CREATE_REQUEST, TRANSPORT_MODE_EXISTING_REQUEST)

DELETE_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "BPC_DELETE_SINGLE_ROLE",
    "UJ0_DELETE_SINGLE_ROLE",
)
DELETE_ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TEXTS", "AGR_TCODES", "AGR_USERS", "TSTC")


def validate_role_name_for_delete(role_name: str) -> str:
    normalized = validate_role_name(role_name)
    if len(normalized) > MAX_ROLE_NAME_LENGTH:
        raise ValueError(
            f"Nome da função excede o limite de {MAX_ROLE_NAME_LENGTH} caracteres (AGR_DEFINE-AGR_NAME)."
        )
    return normalized


def validate_transport_inputs(transport_mode: str, request_number: str, request_description: str) -> dict[str, str]:
    mode = str(transport_mode or TRANSPORT_MODE_LOCAL).strip().upper()
    if mode not in VALID_TRANSPORT_MODES:
        raise ValueError(
            f"Modo de transporte inválido: '{mode}'. Valores aceites: {', '.join(VALID_TRANSPORT_MODES)}."
        )

    if mode == TRANSPORT_MODE_LOCAL:
        return {"transport_mode": mode, "request_number": "", "request_description": ""}

    if mode == TRANSPORT_MODE_EXISTING_REQUEST:
        request_clean = str(request_number or "").strip().upper()
        if not request_clean:
            raise ValueError("Selecione uma Request de transporte existente e aberta.")
        return {"transport_mode": mode, "request_number": request_clean, "request_description": ""}

    normalized_description = validate_request_description(request_description)
    return {"transport_mode": mode, "request_number": "", "request_description": normalized_description}


def _assert_delete_environment_allowed(environment: str) -> str:
    env = str(environment or "").strip().upper()
    if env not in ALLOWED_DELETE_ENVIRONMENTS:
        raise ValueError(
            f"Eliminação individual via RFC só é permitida em DEV. "
            f"Ambiente '{env}' está bloqueado (QAD e PRD não são permitidos para escrita/eliminação)."
        )
    return env


def _error_result(
    environment: str,
    role_name: str,
    error_type: str,
    message: str,
    *,
    details: str | None = None,
    status: str = "ERROR",
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": status,
        "environment": environment,
        "role": role_name,
        "error_type": error_type,
        "message": message,
    }
    if details:
        payload["details"] = details
    return payload


def preview_pfcg_role_delete(
    environment: str,
    role_name: str,
    transport_mode: str = TRANSPORT_MODE_LOCAL,
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Read-only preview da eliminação de função PFCG em DEV."""
    raw_role_upper = str(role_name or "").strip().upper()
    try:
        env = _assert_delete_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), raw_role_upper, "ENVIRONMENT_BLOCKED", str(exc))

    try:
        normalized_role = validate_role_name_for_delete(role_name)
    except ValueError as exc:
        return _error_result(env, raw_role_upper, "INVALID_INPUT", str(exc))

    try:
        transport = validate_transport_inputs(transport_mode, request_number, request_description)
    except ValueError as exc:
        return _error_result(env, normalized_role, "INVALID_TRANSPORT_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, normalized_role, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    guard = make_write_guard(DELETE_ALLOWED_FUNCTIONS, DELETE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    try:
        if not role_exists(connection, guard, normalized_role):
            return _error_result(
                env,
                normalized_role,
                "ROLE_NOT_FOUND",
                f"A função '{normalized_role}' não existe no ambiente {env}.",
                status="NOT_FOUND",
            )

        # 1. Descrição (AGR_TEXTS)
        description = ""
        try:
            text_rows = read_table(
                connection,
                guard,
                table_name="AGR_TEXTS",
                fields=["TEXT"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=1,
            )
            if text_rows and text_rows[0] and text_rows[0][0]:
                description = str(text_rows[0][0]).strip()
        except Exception:
            pass

        # 2. Transações (AGR_TCODES)
        tcodes: list[str] = []
        try:
            tcode_rows = read_table(
                connection,
                guard,
                table_name="AGR_TCODES",
                fields=["TCODE"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=200,
            )
            tcodes = sorted({row[0].strip().upper() for row in tcode_rows if row and row[0].strip()})
        except Exception:
            pass

        # 3. Utilizadores atribuídos (AGR_USERS)
        users_count = 0
        try:
            user_rows = read_table(
                connection,
                guard,
                table_name="AGR_USERS",
                fields=["UNAME"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=200,
            )
            users_count = len({row[0].strip().upper() for row in user_rows if row and row[0].strip()})
        except Exception:
            pass

        transport_preview: dict[str, Any] = {"transport_mode": transport["transport_mode"]}
        if transport["transport_mode"] == TRANSPORT_MODE_EXISTING_REQUEST:
            request_check = validate_transport_request(env, transport["request_number"])
            if not request_check.get("ok"):
                return _error_result(
                    env,
                    normalized_role,
                    request_check.get("error_type") or "REQUEST_NOT_VALID",
                    request_check.get("message") or "A Request de transporte selecionada não é válida.",
                )
            transport_preview.update(
                {
                    "request_number": request_check.get("request"),
                    "request_description": request_check.get("description"),
                    "request_state": request_check.get("state"),
                    "request_target_system": request_check.get("target_system"),
                }
            )
        elif transport["transport_mode"] == TRANSPORT_MODE_CREATE_REQUEST:
            transport_preview.update(
                {
                    "request_description": transport["request_description"],
                    "request_category": "Customizing",
                }
            )

        return {
            "ok": True,
            "status": "PREVIEW_READY",
            "environment": env,
            "role": normalized_role,
            "description": description or "(Sem descrição)",
            "tcodes": tcodes,
            "tcodes_count": len(tcodes),
            "users_count": users_count,
            "transport": transport_preview,
        }
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass


def delete_pfcg_role_rfc(
    environment: str,
    role_name: str,
    transport_mode: str = TRANSPORT_MODE_LOCAL,
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Elimina a função PFCG via RFC (BPC_DELETE_SINGLE_ROLE), apenas em DEV."""
    raw_role_upper = str(role_name or "").strip().upper()
    try:
        env = _assert_delete_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), raw_role_upper, "ENVIRONMENT_BLOCKED", str(exc))

    try:
        normalized_role = validate_role_name_for_delete(role_name)
    except ValueError as exc:
        return _error_result(env, raw_role_upper, "INVALID_INPUT", str(exc))

    try:
        transport = validate_transport_inputs(transport_mode, request_number, request_description)
    except ValueError as exc:
        return _error_result(env, normalized_role, "INVALID_TRANSPORT_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, normalized_role, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    guard = make_write_guard(DELETE_ALLOWED_FUNCTIONS, DELETE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    try:
        if not role_exists(connection, guard, normalized_role):
            return _error_result(
                env,
                normalized_role,
                "ROLE_NOT_FOUND",
                f"A função '{normalized_role}' não existe em {env}.",
                status="NOT_FOUND",
            )

        resolved_request_number = ""
        if transport["transport_mode"] == TRANSPORT_MODE_EXISTING_REQUEST:
            request_check = validate_transport_request(env, transport["request_number"])
            if not request_check.get("ok"):
                return _error_result(
                    env,
                    normalized_role,
                    request_check.get("error_type") or "REQUEST_NOT_VALID",
                    request_check.get("message") or "A Request de transporte selecionada não é válida.",
                )
            resolved_request_number = str(request_check.get("request") or "")
        elif transport["transport_mode"] == TRANSPORT_MODE_CREATE_REQUEST:
            request_create_result = create_transport_request(env, transport["request_description"])
            if not request_create_result.get("ok"):
                return _error_result(
                    env,
                    normalized_role,
                    request_create_result.get("error_type") or "REQUEST_CREATE_FAILED",
                    request_create_result.get("message") or "Falha ao criar nova Request de transporte.",
                    details=request_create_result.get("details"),
                )
            resolved_request_number = str(request_create_result.get("request") or "")

        guard.assert_function_allowed("BPC_DELETE_SINGLE_ROLE")
        call_kwargs: dict[str, Any] = {
            "IV_NAME_PREFIX": "",
            "IV_PREF_AGR_NAME": normalized_role,
            "NO_DIALOG": "X",
            "REQUEST": resolved_request_number,
        }

        try:
            rfc_result = connection.call("BPC_DELETE_SINGLE_ROLE", **call_kwargs)
            messages = rfc_result.get("ET_MESSAGES") or []
            error_msgs = [
                str(m.get("MESSAGE") or m)
                for m in messages
                if isinstance(m, dict) and m.get("TYPE") in {"E", "A"}
            ]
            if error_msgs:
                return _error_result(
                    env,
                    normalized_role,
                    "SAP_DELETE_ERROR",
                    f"Erro retornado pelo SAP ao eliminar função: {'; '.join(error_msgs)}",
                )
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(
                env,
                normalized_role,
                f"DELETE_FAILED_{error_type}",
                f"Falha ao eliminar a função via RFC: {message}",
                details=format_exception(exc),
            )

    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    # Verificação pós-execução via ligação independente
    post_check_connection = None
    try:
        post_check_connection = Connection(**params)
        still_exists = role_exists(post_check_connection, guard, normalized_role)
        if still_exists:
            return _error_result(
                env,
                normalized_role,
                "DELETE_VERIFICATION_FAILED",
                f"A chamada RFC concluiu mas a função {normalized_role} ainda consta em AGR_DEFINE em {env}.",
            )
    except Exception as exc:
        return _error_result(
            env,
            normalized_role,
            "DELETE_POST_CHECK_ERROR",
            f"Função eliminada mas ocorreu erro ao validar no SAP: {exc}",
            details=format_exception(exc),
        )
    finally:
        try:
            if post_check_connection is not None:
                post_check_connection.close()
        except Exception:
            pass

    return {
        "ok": True,
        "status": "DELETED",
        "environment": env,
        "role": normalized_role,
        "transport_mode": transport["transport_mode"],
        "transport_request": resolved_request_number or None,
        "message": f"Função {normalized_role} eliminada com sucesso em {env}.",
    }
