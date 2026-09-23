"""Criação de utilizador SAP via RFC (BAPI_USER_CREATE1 + BAPI_USER_ACTGROUPS_ASSIGN).

O CUA foi desligado em todos os sistemas (DEV/QAD/PRD deixaram de estar ligados a um
central SPA) — por isso a criação é sempre feita diretamente no sistema-alvo, nunca
via um sistema central. Segue o mesmo padrão de escrita RFC já usado em
`sap_rfc/pfcg_role_create_service.py` (guard com whitelist explícita de funções/tabelas,
verificação pós-escrita numa ligação nova e independente).

Interface das BAPIs confirmada por introspeção real (`get_function_description`) em
DEV/QAD/PRD em 2026-09-17 — todas RFC-enabled (FMODE='R' em TFDIR) nos três sistemas:
  BAPI_USER_EXISTENCE_CHECK, BAPI_USER_CREATE1, BAPI_USER_ACTGROUPS_ASSIGN,
  BAPI_TRANSACTION_COMMIT, BAPI_TRANSACTION_ROLLBACK.
"""

from __future__ import annotations

import os
import re
from datetime import date
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    load_project_env,
    make_option_eq,
    make_write_guard,
    read_table,
    role_exists,
)
from sap_rfc.user_data_service import validate_username

# Limites reais das estruturas BAPI (confirmados via get_function_description):
#   BAPIADDR3-FIRSTNAME/LASTNAME -> CHAR40, BAPIADDR3-E_MAIL -> CHAR241
#   BAPIAGR-AGR_NAME -> CHAR30 (igual a AGR_DEFINE-AGR_NAME)
MAX_NAME_LENGTH = 40
MAX_EMAIL_LENGTH = 241

DEFAULT_USTYP = "A"
ALLOWED_USTYP = {"A", "B", "C", "L", "S"}

# CUA deixou de existir: escrita direta permitida nos três sistemas reais.
ALLOWED_CREATE_ENVIRONMENTS = ("DEV", "QAD", "PRD")

# Sentinela SAP-standard para "sem data fim" em atribuições de função (AGR_USERS-TO_DAT).
ROLE_ASSIGNMENT_NO_END_DATE = "99991231"

WRITE_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "BAPI_USER_EXISTENCE_CHECK",
    "BAPI_USER_CREATE1",
    "BAPI_USER_ACTGROUPS_ASSIGN",
    "BAPI_TRANSACTION_COMMIT",
    "BAPI_TRANSACTION_ROLLBACK",
)
WRITE_ALLOWED_TABLES = ("USR02", "AGR_USERS", "AGR_DEFINE")

# BAPI_USER_CHANGE: mesma whitelist de escrita, mas sem BAPI_USER_CREATE1/
# BAPI_USER_ACTGROUPS_ASSIGN (não fazem sentido neste fluxo de reset de password).
CHANGE_PASSWORD_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "BAPI_USER_EXISTENCE_CHECK",
    "BAPI_USER_CHANGE",
    "BAPI_TRANSACTION_COMMIT",
    "BAPI_TRANSACTION_ROLLBACK",
)
CHANGE_PASSWORD_ALLOWED_TABLES = ("USR02",)

# BAPI_USER_UNLOCK: fluxo de escrita separado do reset de password. Mantem a
# mesma leitura pos-escrita de USR02 para confirmar se o bloqueio ficou removido.
USER_UNLOCK_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "BAPI_USER_EXISTENCE_CHECK",
    "BAPI_USER_UNLOCK",
    "BAPI_TRANSACTION_COMMIT",
    "BAPI_TRANSACTION_ROLLBACK",
)
USER_UNLOCK_ALLOWED_TABLES = ("USR02",)

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_DATE_DIGITS_RE = re.compile(r"^\d{8}$")


def _assert_create_environment_allowed(environment: str) -> str:
    env = str(environment or "").strip().upper()
    if env not in ALLOWED_CREATE_ENVIRONMENTS:
        raise ValueError(
            f"Ambiente '{env}' não permitido. Use um de: {', '.join(ALLOWED_CREATE_ENVIRONMENTS)}."
        )
    return env


def validate_person_name(raw: str, label: str) -> str:
    value = str(raw or "").strip()
    if not value:
        raise ValueError(f"Informe o {label} do utilizador.")
    if len(value) > MAX_NAME_LENGTH:
        raise ValueError(f"O {label} excede o limite real de {MAX_NAME_LENGTH} caracteres (BAPIADDR3).")
    return value


def validate_email(raw: str) -> str:
    value = str(raw or "").strip()
    if not value:
        return ""
    if len(value) > MAX_EMAIL_LENGTH:
        raise ValueError(f"O email excede o limite real de {MAX_EMAIL_LENGTH} caracteres (BAPIADDR3-E_MAIL).")
    if not _EMAIL_RE.match(value):
        raise ValueError(f"Email inválido: '{value}'.")
    return value


def validate_ustyp(raw: str) -> str:
    value = str(raw or DEFAULT_USTYP).strip().upper()
    if value not in ALLOWED_USTYP:
        raise ValueError(f"Tipo de utilizador inválido: '{value}'. Use um de: {', '.join(sorted(ALLOWED_USTYP))}.")
    return value


def validate_group(raw: str) -> str:
    # Vazio por default: confirmado com o utilizador de referência S198, que também
    # não tem CLASS preenchido em DEV/QAD/PRD — nunca inventar um grupo default.
    return str(raw or "").strip().upper()


def validate_department(raw: str) -> str:
    # BAPIADDR3-DEPARTMENT: CHAR40, mesmo limite real das colunas de nome.
    value = str(raw or "").strip()
    if len(value) > MAX_NAME_LENGTH:
        raise ValueError(f"O departamento excede o limite real de {MAX_NAME_LENGTH} caracteres (BAPIADDR3).")
    return value


def validate_function(raw: str) -> str:
    # BAPIADDR3-FUNCTION: CHAR40, mesmo limite real das colunas de nome.
    value = str(raw or "").strip()
    if len(value) > MAX_NAME_LENGTH:
        raise ValueError(f"A função excede o limite real de {MAX_NAME_LENGTH} caracteres (BAPIADDR3).")
    return value


def _normalize_sap_date(raw: Any) -> str:
    value = str(raw or "").strip().replace("-", "").replace("/", "")
    if not value:
        return ""
    if not _DATE_DIGITS_RE.match(value):
        raise ValueError(f"Data inválida: '{raw}'. Use AAAAMMDD.")
    return value


def validate_valid_from(raw: Any) -> str:
    value = _normalize_sap_date(raw)
    # GLTGV é obrigatório na LOGONDATA; sem valor indicado, assume-se hoje.
    return value or date.today().strftime("%Y%m%d")


def validate_valid_to(raw: Any) -> str:
    # Vazio == sem limite (mesmo padrão observado em S198: GLTGB devolvido como
    # '00000000' em PRD/QAD quando nunca foi definida data fim).
    return _normalize_sap_date(raw)


def normalize_roles(raw: Any) -> list[str]:
    if isinstance(raw, (list, tuple, set)):
        parts = [str(item) for item in raw]
    else:
        parts = re.split(r"[;,\n]+", str(raw or ""))
    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        value = part.strip().upper()
        if not value or value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


def resolve_default_password() -> str:
    """Senha inicial default (aplicada quando o Excel/pedido não indica uma):
    vem de SAP_PASSE_PASSWD no `.env` do projeto — nunca hardcoded no código."""
    return os.getenv("SAP_PASSE_PASSWD", "").strip()


def validate_password(raw: Any) -> str:
    value = str(raw or "").strip() or resolve_default_password()
    if not value:
        raise ValueError(
            "Sem senha inicial: indique uma ou configure SAP_PASSE_PASSWD no .env do projeto."
        )
    return value


def _error_result(
    environment: str,
    username: str,
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
        "username": username,
        "error_type": error_type,
        "message": message,
    }
    if details:
        payload["details"] = details
    return payload


def _validate_inputs(
    environment: str,
    username: str,
    first_name: str,
    last_name: str,
    email: str,
    ustyp: str,
    group: str,
    valid_from: Any,
    valid_to: Any,
    password: Any,
    roles: Any,
    department: str = "",
    function: str = "",
) -> tuple[str, str, str, str, str, str, str, str, str, str, list[str], str, str] | dict[str, Any]:
    """Returns (env, user, first, last, mail, ustyp, group, gltgv, gltgb, password, roles, department, function) or an error dict."""
    raw_user_upper = str(username or "").strip().upper()
    try:
        env = _assert_create_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), raw_user_upper, "ENVIRONMENT_BLOCKED", str(exc))

    try:
        norm_user = validate_username(username)
    except ValueError as exc:
        return _error_result(env, raw_user_upper, "INVALID_INPUT", str(exc))

    try:
        first = validate_person_name(first_name, "nome")
        last = validate_person_name(last_name, "apelido")
        mail = validate_email(email)
        ustyp_v = validate_ustyp(ustyp)
        group_v = validate_group(group)
        gltgv = validate_valid_from(valid_from)
        gltgb = validate_valid_to(valid_to)
        pwd = validate_password(password)
        roles_list = normalize_roles(roles)
        department_v = validate_department(department)
        function_v = validate_function(function)
    except ValueError as exc:
        return _error_result(env, norm_user, "INVALID_INPUT", str(exc))

    return env, norm_user, first, last, mail, ustyp_v, group_v, gltgv, gltgb, pwd, roles_list, department_v, function_v


def _user_exists(connection: Any, guard: Any, username: str) -> bool:
    guard.assert_function_allowed("BAPI_USER_EXISTENCE_CHECK")
    result = connection.call("BAPI_USER_EXISTENCE_CHECK", USERNAME=username)
    # Confirmado por introspeção real em QAD (2026-09-17): TYPE vem sempre 'I' (mensagem
    # informativa), tanto para "existe" como "não existe" — não dá para usar TYPE. O sinal
    # correto é NUMBER da classe de mensagem 01: '088' = existe, '124' = não existe
    # (independente do idioma de logon, ao contrário do texto em MESSAGE).
    ret = result.get("RETURN") or {}
    return str(ret.get("NUMBER") or "").strip() == "088"


def _decode_user_lock_status(raw_uflag: Any, raw_locnt: Any = "") -> dict[str, Any]:
    raw_uflag_text = str(raw_uflag or "").strip()
    try:
        uflag = int(raw_uflag_text or "0")
    except ValueError:
        uflag = 0

    raw_locnt_text = str(raw_locnt or "").strip()
    try:
        failed_logon_count = int(raw_locnt_text or "0")
    except ValueError:
        failed_logon_count = None

    reasons: list[str] = []
    if uflag & 32:
        reasons.append("Bloqueado globalmente por administrador")
    if uflag & 64:
        reasons.append("Bloqueado localmente por administrador")
    if uflag & 128:
        reasons.append("Bloqueado por tentativas excessivas de logon incorreto")

    if uflag and not reasons:
        reasons.append("Bloqueado (motivo SAP não mapeado)")

    return {
        "uflag": raw_uflag_text or "0",
        "failed_logon_count": failed_logon_count,
        "locked": bool(uflag),
        "login_failed_locked": bool(uflag & 128),
        "lock_reasons": reasons,
        "lock_status": "; ".join(reasons) if reasons else "Ativo",
    }


def _read_user_lock_status(connection: Any, guard: Any, username: str) -> dict[str, Any]:
    rows = read_table(
        connection,
        guard,
        table_name="USR02",
        fields=["BNAME", "UFLAG", "LOCNT"],
        options=make_option_eq("BNAME", username),
        rowcount=1,
    )
    if not rows:
        return {
            "lock_status": "Utilizador não encontrado na verificação pós-alteração",
            "locked": None,
            "login_failed_locked": None,
            "uflag": None,
            "failed_logon_count": None,
            "lock_reasons": [],
        }
    row = rows[0]
    return _decode_user_lock_status(row[1] if len(row) > 1 else "", row[2] if len(row) > 2 else "")


def preview_user_create(
    environment: str,
    username: str,
    first_name: str,
    last_name: str,
    email: str,
    ustyp: str = DEFAULT_USTYP,
    group: str = "",
    valid_from: Any = "",
    valid_to: Any = "",
    password: Any = "",
    roles: Any = None,
    department: str = "",
    function: str = "",
) -> dict[str, Any]:
    """Read-only preview: valida ambiente/campos, confirma que o utilizador ainda não
    existe e que as funções indicadas existem no sistema-alvo. Nunca escreve em SAP."""
    validated = _validate_inputs(
        environment, username, first_name, last_name, email, ustyp, group,
        valid_from, valid_to, password, roles, department, function,
    )
    if isinstance(validated, dict):
        return validated
    env, norm_user, first, last, mail, ustyp_v, group_v, gltgv, gltgb, pwd, roles_list, department_v, function_v = validated

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, norm_user, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    try:
        try:
            already_exists = _user_exists(connection, guard, norm_user)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, norm_user, f"EXISTENCE_CHECK_{error_type}", message, details=format_exception(exc))

        if already_exists:
            return _error_result(
                env, norm_user, "USER_ALREADY_EXISTS",
                f"O utilizador {norm_user} já existe em {env}. Não é possível criar novamente.",
            )

        missing_roles: list[str] = []
        if roles_list:
            try:
                missing_roles = [role for role in roles_list if not role_exists(connection, guard, role)]
            except Exception as exc:
                error_type, message = classify_rfc_error(exc)
                return _error_result(env, norm_user, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        if missing_roles:
            payload = _error_result(
                env, norm_user, "ROLES_NOT_FOUND",
                f"As seguintes funções não existem em {env}: {', '.join(missing_roles)}",
            )
            payload["missing_roles"] = missing_roles
            return payload

        return {
            "ok": True,
            "status": "PREVIEW_READY",
            "environment": env,
            "username": norm_user,
            "first_name": first,
            "last_name": last,
            "email": mail or "-",
            "ustyp": ustyp_v,
            "group": group_v or "(vazio)",
            "valid_from": gltgv,
            "valid_to": gltgb or "sem limite",
            "roles": roles_list,
            "department": department_v,
            "function": function_v,
            "password_source": "Excel/pedido" if str(password or "").strip() else "SAP_PASSE_PASSWD (.env)",
        }
    finally:
        try:
            connection.close()
        except Exception:
            pass


def create_user_rfc(
    environment: str,
    username: str,
    first_name: str,
    last_name: str,
    email: str,
    ustyp: str = DEFAULT_USTYP,
    group: str = "",
    valid_from: Any = "",
    valid_to: Any = "",
    password: Any = "",
    roles: Any = None,
    department: str = "",
    function: str = "",
) -> dict[str, Any]:
    """Criação REAL de utilizador via RFC (BAPI_USER_CREATE1), com atribuição de
    funções (BAPI_USER_ACTGROUPS_ASSIGN) quando indicadas, seguida de
    BAPI_TRANSACTION_COMMIT. Único ponto de entrada de escrita RFC para este fluxo."""
    validated = _validate_inputs(
        environment, username, first_name, last_name, email, ustyp, group,
        valid_from, valid_to, password, roles, department, function,
    )
    if isinstance(validated, dict):
        return validated
    env, norm_user, first, last, mail, ustyp_v, group_v, gltgv, gltgb, pwd, roles_list, department_v, function_v = validated

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, norm_user, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    try:
        try:
            already_exists = _user_exists(connection, guard, norm_user)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, norm_user, f"EXISTENCE_CHECK_{error_type}", message, details=format_exception(exc))

        if already_exists:
            return _error_result(
                env, norm_user, "USER_ALREADY_EXISTS",
                f"O utilizador {norm_user} já existe em {env}. Não é possível criar novamente.",
            )

        guard.assert_function_allowed("BAPI_USER_CREATE1")
        try:
            create_result = connection.call(
                "BAPI_USER_CREATE1",
                USERNAME=norm_user,
                ADDRESS={
                    "FIRSTNAME": first,
                    "LASTNAME": last,
                    "E_MAIL": mail,
                    "DEPARTMENT": department_v,
                    "FUNCTION": function_v,
                },
                LOGONDATA={"USTYP": ustyp_v, "CLASS": group_v, "GLTGV": gltgv, "GLTGB": gltgb},
                PASSWORD={"BAPIPWD": pwd},
            )
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return _error_result(env, norm_user, f"BAPI_USER_CREATE1_{error_type}", message, details=format_exception(exc))

        return_rows = create_result.get("RETURN") or []
        error_messages = [
            str(row.get("MESSAGE") or "").strip()
            for row in return_rows
            if str(row.get("TYPE") or "").strip().upper() in {"E", "A"} and str(row.get("MESSAGE") or "").strip()
        ]
        if error_messages:
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return {
                "ok": False,
                "status": "CREATE_FAILED",
                "environment": env,
                "username": norm_user,
                "sap_return_messages": error_messages,
                "message": "Não foi possível criar o utilizador (BAPI_USER_CREATE1 devolveu erro).",
            }

        role_assign_messages: list[str] = []
        if roles_list:
            guard.assert_function_allowed("BAPI_USER_ACTGROUPS_ASSIGN")
            activity_groups = [
                {
                    "AGR_NAME": role,
                    "FROM_DAT": gltgv,
                    "TO_DAT": gltgb or ROLE_ASSIGNMENT_NO_END_DATE,
                }
                for role in roles_list
            ]
            try:
                assign_result = connection.call(
                    "BAPI_USER_ACTGROUPS_ASSIGN",
                    USERNAME=norm_user,
                    ACTIVITYGROUPS=activity_groups,
                )
                role_assign_messages = [
                    str(row.get("MESSAGE") or "").strip()
                    for row in (assign_result.get("RETURN") or [])
                    if str(row.get("TYPE") or "").strip().upper() in {"E", "A"} and str(row.get("MESSAGE") or "").strip()
                ]
            except Exception as exc:
                _, assign_error_message = classify_rfc_error(exc)
                role_assign_messages = [assign_error_message]

        # Utilizador criado (e, se aplicável, funções atribuídas) só fica persistido
        # depois deste commit explícito — sem ele, tudo fica preso na transação LUW aberta.
        guard.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
        try:
            connection.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception as exc:
            _, commit_error_message = classify_rfc_error(exc)
            return _error_result(
                env, norm_user, "COMMIT_FAILED",
                f"BAPI_USER_CREATE1 não devolveu erro, mas o COMMIT falhou: {commit_error_message}",
            )
    finally:
        try:
            connection.close()
        except Exception:
            pass

    # Verificação pós-escrita com ligação NOVA e independente: só uma leitura nova
    # sem qualquer relação com a ligação que escreveu prova persistência real em BD
    # (isolamento de transação impede que outra sessão veja dados não confirmados).
    verify_connection = None
    try:
        verify_connection = Connection(**params)
        verify_guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)

        user_rows = read_table(
            verify_connection, verify_guard, table_name="USR02",
            fields=["BNAME"], options=make_option_eq("BNAME", norm_user), rowcount=1,
        )
        if not user_rows:
            payload = _error_result(
                env, norm_user, "WRITE_NOT_PERSISTED",
                "O COMMIT não gerou exceção, mas o utilizador não foi encontrado numa ligação "
                "nova e independente (USR02). Não repetir a criação sem investigar antes.",
            )
            if error_messages:
                payload["sap_return_messages"] = error_messages
            return payload

        created_roles: list[str] = []
        if roles_list:
            role_rows = read_table(
                verify_connection, verify_guard, table_name="AGR_USERS",
                fields=["AGR_NAME"], options=make_option_eq("UNAME", norm_user), rowcount=0,
            )
            created_roles = sorted({row[0].strip().upper() for row in role_rows if row[0].strip()})
        missing_roles_after = sorted(set(roles_list) - set(created_roles))

        result_payload: dict[str, Any] = {
            "ok": True,
            "status": "CREATED" if not missing_roles_after and not role_assign_messages else "CREATED_PARTIAL",
            "environment": env,
            "username": norm_user,
            "first_name": first,
            "last_name": last,
            "email": mail or "-",
            "ustyp": ustyp_v,
            "group": group_v or "(vazio)",
            "valid_from": gltgv,
            "valid_to": gltgb or "sem limite",
            "department": department_v,
            "function": function_v,
            "roles_requested": roles_list,
            "roles_assigned": created_roles,
        }
        if missing_roles_after or role_assign_messages:
            result_payload["message"] = (
                "Utilizador criado com sucesso, mas nem todas as funções pedidas ficaram "
                "confirmadas na verificação pós-escrita independente."
            )
            if missing_roles_after:
                result_payload["roles_missing"] = missing_roles_after
            if role_assign_messages:
                result_payload["role_assignment_messages"] = role_assign_messages
        return result_payload
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, norm_user, f"POST_VALIDATION_{error_type}", message, details=format_exception(exc))
    finally:
        try:
            if verify_connection is not None:
                verify_connection.close()
        except Exception:
            pass


def unlock_user_rfc(environment: str, username: str) -> dict[str, Any]:
    """Desbloqueia um utilizador SAP existente via BAPI_USER_UNLOCK.

    Este fluxo nunca altera a password. Depois da BAPI e do commit, abre uma
    ligacao nova para validar o estado final em USR02-UFLAG/LOCNT.
    """
    try:
        env = _assert_create_environment_allowed(environment)
        norm_user = validate_username(username)
    except ValueError as exc:
        raw_user_upper = str(username or "").strip().upper()
        return _error_result(str(environment or "").strip().upper(), raw_user_upper, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, norm_user, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    guard = make_write_guard(USER_UNLOCK_ALLOWED_FUNCTIONS, USER_UNLOCK_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    try:
        try:
            exists = _user_exists(connection, guard, norm_user)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, norm_user, f"EXISTENCE_CHECK_{error_type}", message, details=format_exception(exc))

        if not exists:
            return _error_result(
                env, norm_user, "USER_NOT_FOUND",
                f"O utilizador {norm_user} não existe em {env}. Não é possível desbloquear.",
            )

        guard.assert_function_allowed("BAPI_USER_UNLOCK")
        try:
            unlock_result = connection.call("BAPI_USER_UNLOCK", USERNAME=norm_user)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return _error_result(env, norm_user, f"BAPI_USER_UNLOCK_{error_type}", message, details=format_exception(exc))

        return_rows = unlock_result.get("RETURN") or []
        error_messages = [
            str(row.get("MESSAGE") or "").strip()
            for row in return_rows
            if str(row.get("TYPE") or "").strip().upper() in {"E", "A"} and str(row.get("MESSAGE") or "").strip()
        ]
        if error_messages:
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return {
                "ok": False,
                "status": "UNLOCK_FAILED",
                "environment": env,
                "username": norm_user,
                "sap_return_messages": error_messages,
                "message": "Não foi possível desbloquear o utilizador (BAPI_USER_UNLOCK devolveu erro).",
            }

        guard.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
        try:
            connection.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception as exc:
            _, commit_error_message = classify_rfc_error(exc)
            return _error_result(
                env, norm_user, "COMMIT_FAILED",
                f"BAPI_USER_UNLOCK não devolveu erro, mas o COMMIT falhou: {commit_error_message}",
            )
    finally:
        try:
            connection.close()
        except Exception:
            pass

    payload: dict[str, Any] = {
        "environment": env,
        "username": norm_user,
    }

    verify_connection = None
    try:
        verify_connection = Connection(**params)
        verify_guard = make_write_guard(USER_UNLOCK_ALLOWED_FUNCTIONS, USER_UNLOCK_ALLOWED_TABLES)
        payload.update(_read_user_lock_status(verify_connection, verify_guard, norm_user))
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        payload["lock_status"] = "Não foi possível verificar o bloqueio após desbloquear o utilizador."
        payload["lock_check_error_type"] = error_type
        payload["lock_check_message"] = message
    finally:
        try:
            if verify_connection is not None:
                verify_connection.close()
        except Exception:
            pass

    if payload.get("locked") is False:
        payload["ok"] = True
        payload["status"] = "USER_UNLOCKED"
        payload["message"] = "Utilizador desbloqueado e sem bloqueio ativo em USR02."
    elif payload.get("locked") is True:
        payload["ok"] = False
        payload["status"] = "UNLOCK_NOT_CONFIRMED"
        payload["message"] = f"Desbloqueio executado, mas o utilizador continua bloqueado: {payload.get('lock_status')}."
    else:
        payload["ok"] = False
        payload["status"] = "UNLOCK_CHECK_FAILED"
        payload["message"] = "Desbloqueio executado, mas não foi possível verificar o estado em USR02."

    return payload


def change_password_rfc(environment: str, username: str, password: Any = "") -> dict[str, Any]:
    """Redefine a password de um utilizador SAP já existente via BAPI_USER_CHANGE.

    Ao contrário de BAPI_USER_CREATE1 (cujo PASSWORD-BAPIPWD é suficiente sozinho,
    confirmado por introspeção real), BAPI_USER_CHANGE exige também PASSWORDX-BAPIPWD='X'
    para sinalizar que o campo PASSWORD deve mesmo ser aplicado - sem isso a chamada
    é aceite mas a password não muda. Único ponto de entrada de escrita RFC para este
    fluxo; usado quando a password inicial definida na criação não ficou operacional."""
    try:
        env = _assert_create_environment_allowed(environment)
        norm_user = validate_username(username)
    except ValueError as exc:
        raw_user_upper = str(username or "").strip().upper()
        return _error_result(str(environment or "").strip().upper(), raw_user_upper, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, norm_user, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        pwd = validate_password(password)
    except ValueError as exc:
        return _error_result(env, norm_user, "INVALID_INPUT", str(exc))

    try:
        from pyrfc import Connection
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    guard = make_write_guard(CHANGE_PASSWORD_ALLOWED_FUNCTIONS, CHANGE_PASSWORD_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, norm_user, error_type, message, details=format_exception(exc))

    try:
        try:
            exists = _user_exists(connection, guard, norm_user)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, norm_user, f"EXISTENCE_CHECK_{error_type}", message, details=format_exception(exc))

        if not exists:
            return _error_result(
                env, norm_user, "USER_NOT_FOUND",
                f"O utilizador {norm_user} não existe em {env}. Não é possível alterar a password.",
            )

        guard.assert_function_allowed("BAPI_USER_CHANGE")
        try:
            change_result = connection.call(
                "BAPI_USER_CHANGE",
                USERNAME=norm_user,
                PASSWORD={"BAPIPWD": pwd},
                PASSWORDX={"BAPIPWD": "X"},
                PRODUCTIVE_PWD=pwd,
            )
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return _error_result(env, norm_user, f"BAPI_USER_CHANGE_{error_type}", message, details=format_exception(exc))

        return_rows = change_result.get("RETURN") or []
        error_messages = [
            str(row.get("MESSAGE") or "").strip()
            for row in return_rows
            if str(row.get("TYPE") or "").strip().upper() in {"E", "A"} and str(row.get("MESSAGE") or "").strip()
        ]
        if error_messages:
            try:
                guard.assert_function_allowed("BAPI_TRANSACTION_ROLLBACK")
                connection.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return {
                "ok": False,
                "status": "CHANGE_FAILED",
                "environment": env,
                "username": norm_user,
                "sap_return_messages": error_messages,
                "message": "Não foi possível alterar a password (BAPI_USER_CHANGE devolveu erro).",
            }

        guard.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
        try:
            connection.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception as exc:
            _, commit_error_message = classify_rfc_error(exc)
            return _error_result(
                env, norm_user, "COMMIT_FAILED",
                f"BAPI_USER_CHANGE não devolveu erro, mas o COMMIT falhou: {commit_error_message}",
            )
    finally:
        try:
            connection.close()
        except Exception:
            pass

    payload: dict[str, Any] = {
        "ok": True,
        "status": "PASSWORD_CHANGED",
        "environment": env,
        "username": norm_user,
        "password_source": "Excel/pedido" if str(password or "").strip() else "SAP_PASSE_PASSWD (.env)",
    }

    verify_connection = None
    try:
        verify_connection = Connection(**params)
        verify_guard = make_write_guard(CHANGE_PASSWORD_ALLOWED_FUNCTIONS, CHANGE_PASSWORD_ALLOWED_TABLES)
        payload.update(_read_user_lock_status(verify_connection, verify_guard, norm_user))
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        payload["lock_status"] = "Não foi possível verificar o bloqueio após alterar a password."
        payload["lock_check_error_type"] = error_type
        payload["lock_check_message"] = message
    finally:
        try:
            if verify_connection is not None:
                verify_connection.close()
        except Exception:
            pass

    if payload.get("locked") is True:
        unlock_result = unlock_user_rfc(env, norm_user)

        for lock_key in (
            "uflag",
            "failed_logon_count",
            "locked",
            "login_failed_locked",
            "lock_reasons",
            "lock_status",
            "lock_check_error_type",
            "lock_check_message",
        ):
            if lock_key in unlock_result:
                payload[lock_key] = unlock_result[lock_key]

        if unlock_result.get("ok") is True and unlock_result.get("locked") is False:
            payload["status"] = "PASSWORD_CHANGED_AND_UNLOCKED"
            payload["unlocked"] = True
            payload["message"] = (
                "Password alterada e utilizador desbloqueado com sucesso (USR02 sem bloqueio ativo)."
            )
        elif unlock_result.get("locked") is True:
            payload["status"] = "PASSWORD_CHANGED"
            payload["unlocked"] = False
            payload["message"] = (
                f"Password alterada, mas o utilizador continua bloqueado: {payload.get('lock_status')}."
            )
        else:
            payload["status"] = "PASSWORD_CHANGED"
            payload["unlocked"] = False
            if unlock_result.get("message"):
                payload["unlock_message"] = unlock_result["message"]
            if unlock_result.get("sap_return_messages"):
                payload["unlock_sap_return_messages"] = unlock_result["sap_return_messages"]
            payload["message"] = (
                f"Password alterada, mas não foi possível desbloquear o utilizador: "
                f"{unlock_result.get('message') or 'falha na operação de desbloqueio'}."
            )
    elif payload.get("locked") is False:
        payload["status"] = "PASSWORD_CHANGED"
        payload["unlocked"] = False
        payload["message"] = "Password alterada e utilizador sem bloqueio ativo em USR02."
    else:
        payload["status"] = "PASSWORD_CHANGED"
        payload["unlocked"] = False
        payload["message"] = "Password alterada, mas não foi possível verificar o bloqueio após alterar a password."

    return payload

