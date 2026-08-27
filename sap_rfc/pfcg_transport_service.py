from __future__ import annotations

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
)

# Limite real confirmado via DDIF_FIELDINFO_GET: E07T-AS4TEXT -> CHAR60.
# CTS_WBO_CREATE_REQUEST aceita DESCRIPTION até 80 (TEXT80) na assinatura RFC, mas o
# armazenamento real em E07T trunca para 60 caracteres — por isso a validação aplicativa
# usa 60 e REJEITA (nunca trunca automaticamente) descrições maiores que isso.
MAX_REQUEST_DESCRIPTION_LENGTH = 60

# Categoria funcionalmente correta para este processo (PFCG = objeto de Customizing).
# Confirmada por convenção real já existente no projeto (Processos/criar_request.py e
# "A. PFCG_CREATE.py" usam sempre tipo="1"=Customizing como default) e pelos valores reais
# do domínio TRFUNCTION (DD07T, via RFC_READ_TABLE): 'W' = "Ordem customizing".
# CTS_WBO_CREATE_REQUEST teria por default nativo 'K' (Workbench) — é sobreposto aqui.
TRANSPORT_CATEGORY_CUSTOMIZING = "W"

# Valor de STATE (texto humano devolvido por CTS_WBO_SELECT_REQUESTS /
# CTS_WBO_GET_REQUEST_DETAILS) que identifica uma Request aberta/modificável.
# Confirmado empiricamente em teste real: TRSTATUS aceita apenas este texto — os códigos
# clássicos de E070-TRSTATUS ('D', 'd', etc.) fazem no-op silencioso nesta RFC.
TRANSPORT_STATE_OPEN = "CHANGEABLE"

# Apenas DEV pode receber escrita nesta primeira fase (mesma restrição de ambiente
# aplicada em sap_rfc.pfcg_role_create_service).
ALLOWED_TRANSPORT_ENVIRONMENTS = ("DEV",)

TRANSPORT_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "CTS_WBO_SELECT_REQUESTS",
    "CTS_WBO_CREATE_REQUEST",
    "CTS_WBO_GET_REQUEST_DETAILS",
    "RFC_READ_TABLE",
)
# E070 (tasks filhas de uma request, via STRKORR), E071 (lista de objetos, ex.: R3TR AGR/PROF)
# e E071K (chaves de tabelas de customizing, ex.: AGR_DEFINE/AGR_TCODES/AGR_1251) — as mesmas
# tabelas usadas pela SE01/SE09 para mostrar "o que está atribuído a esta request". Confirmado
# empiricamente como legível pelo utilizador RFC (devolve dados reais de outros objetos).
TRANSPORT_ALLOWED_TABLES: tuple[str, ...] = ("E070", "E071", "E071K")


def _assert_transport_environment_allowed(environment: str) -> str:
    env = str(environment or "").strip().upper()
    if env not in ALLOWED_TRANSPORT_ENVIRONMENTS:
        raise ValueError(
            f"Operações de transporte via RFC só são permitidas em DEV nesta fase. "
            f"Ambiente '{env}' está bloqueado (QAD e PRD não são permitidos para escrita)."
        )
    return env


def _error_result(environment: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERROR",
        "environment": environment,
        "error_type": error_type,
        "message": message,
    }
    if details:
        payload["details"] = details
    return payload


def validate_request_description(description: str) -> str:
    normalized = str(description or "").strip()
    if not normalized:
        raise ValueError("Informe uma descrição para a Request de transporte.")
    if len(normalized) > MAX_REQUEST_DESCRIPTION_LENGTH:
        raise ValueError(
            f"A descrição da Request excede o limite real de {MAX_REQUEST_DESCRIPTION_LENGTH} "
            "caracteres (E07T-AS4TEXT). Não é truncada automaticamente — reduza o texto."
        )
    return normalized


def search_open_transport_requests(environment: str) -> dict[str, Any]:
    """Read-only: lista as Requests abertas/utilizáveis pelo utilizador RFC em `environment`.

    Reimplementação via RFC (CTS_WBO_SELECT_REQUESTS) do critério funcional usado por
    Processos/pesquisar_request.py (E070 filtrado por AS4USER=utilizador atual e
    TRSTATUS='D'=aberta, via SE16H): aqui filtra-se por OWNER=utilizador RFC e
    TRSTATUS='CHANGEABLE' (valor real confirmado empiricamente, não o código E070 'D').
    Nunca escreve em SAP.
    """
    try:
        env = _assert_transport_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), "ENVIRONMENT_BLOCKED", str(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    guard = make_write_guard(TRANSPORT_ALLOWED_FUNCTIONS, TRANSPORT_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")

        guard.assert_function_allowed("CTS_WBO_SELECT_REQUESTS")
        result = connection.call(
            "CTS_WBO_SELECT_REQUESTS",
            OWNER=params["user"],
            TRSTATUS=TRANSPORT_STATE_OPEN,
        )
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    exception_text = str(result.get("EXCEPTION") or "").strip()
    if exception_text:
        return _error_result(env, "CTS_WBO_SELECT_REQUESTS_EXCEPTION", exception_text)

    rows = result.get("REQUESTS") or []
    requests = [
        {
            "request": str(row.get("REQUEST") or "").strip(),
            "description": str(row.get("DESCRIPTION") or "").strip(),
            "owner": str(row.get("OWNER") or "").strip(),
            "trtype": str(row.get("TRTYPE") or "").strip(),
            "target_system": str(row.get("TARSYSTEM") or "").strip(),
            "state": str(row.get("STATE") or "").strip(),
        }
        for row in rows
        if str(row.get("REQUEST") or "").strip()
    ]
    requests.sort(key=lambda item: item["request"])

    return {
        "ok": True,
        "status": "SEARCH_READY",
        "environment": env,
        "owner": params["user"],
        "requests": requests,
        "requests_count": len(requests),
    }


def validate_transport_request(environment: str, request_number: str) -> dict[str, Any]:
    """Read-only: (re)valida uma Request específica — existe / continua aberta / é utilizável.

    Usado tanto quando o utilizador seleciona uma Request existente na pré-visualização
    como, novamente, imediatamente antes da escrita real (proteção contra condição de
    corrida entre a pré-visualização e a confirmação). Nunca escreve em SAP.
    """
    try:
        env = _assert_transport_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), "ENVIRONMENT_BLOCKED", str(exc))

    request_clean = str(request_number or "").strip().upper()
    if not request_clean:
        return _error_result(env, "INVALID_INPUT", "Informe o número da Request de transporte.")

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    guard = make_write_guard(TRANSPORT_ALLOWED_FUNCTIONS, TRANSPORT_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")

        guard.assert_function_allowed("CTS_WBO_GET_REQUEST_DETAILS")
        result = connection.call("CTS_WBO_GET_REQUEST_DETAILS", REQUESTID=request_clean)
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    exception_text = str(result.get("EXCEPTION") or "").strip()
    detail = result.get("REQUEST_DETAIL") or {}
    request_found = str(detail.get("REQUEST") or "").strip()

    if exception_text or not request_found:
        return _error_result(
            env,
            "REQUEST_NOT_FOUND",
            f"A Request {request_clean} não existe ou não é acessível em {env}.",
            details=exception_text or None,
        )

    state = str(detail.get("STATE") or "").strip()
    owner = str(detail.get("OWNER") or "").strip()
    target_system = str(detail.get("TARSYSTEM") or "").strip()

    if state != TRANSPORT_STATE_OPEN:
        return _error_result(
            env,
            "REQUEST_NOT_OPEN",
            f"A Request {request_clean} não está aberta (estado atual: {state or 'desconhecido'}). "
            "Só é possível utilizar Requests com estado CHANGEABLE.",
        )

    if owner and owner.upper() != str(params["user"]).upper():
        return _error_result(
            env,
            "REQUEST_NOT_OWNED",
            f"A Request {request_clean} pertence a outro utilizador ({owner}), não a {params['user']}.",
        )

    return {
        "ok": True,
        "status": "REQUEST_VALID",
        "environment": env,
        "request": request_found,
        "description": str(detail.get("DESCRIPTION") or "").strip(),
        "trtype": str(detail.get("TRTYPE") or "").strip(),
        "target_system": target_system,
        "owner": owner,
        "state": state,
    }


def create_transport_request(environment: str, description: str) -> dict[str, Any]:
    """ESCRITA: cria uma nova Request de transporte via CTS_WBO_CREATE_REQUEST (Customizing).

    Só deve ser chamada a partir do passo real de confirmação de criação (nunca a partir
    de uma pré-visualização). Verificação pós-escrita com ligação nova e independente via
    validate_transport_request(), seguindo o mesmo princípio de create_pfcg_role_rfc: nunca
    assumir sucesso apenas pela ausência de exceção.
    """
    try:
        env = _assert_transport_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), "ENVIRONMENT_BLOCKED", str(exc))

    try:
        normalized_description = validate_request_description(description)
    except ValueError as exc:
        return _error_result(env, "INVALID_INPUT", str(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    guard = make_write_guard(TRANSPORT_ALLOWED_FUNCTIONS, TRANSPORT_ALLOWED_TABLES)
    connection = None
    call_error_type: str | None = None
    call_error_message: str | None = None
    call_error_details: str | None = None
    new_request = ""
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")

        guard.assert_function_allowed("CTS_WBO_CREATE_REQUEST")
        try:
            result = connection.call(
                "CTS_WBO_CREATE_REQUEST",
                CATEGORY=TRANSPORT_CATEGORY_CUSTOMIZING,
                DESCRIPTION=normalized_description,
                OWNER=params["user"],
            )
            exception_text = str(result.get("EXCEPTION") or "").strip()
            if exception_text:
                call_error_type = "CTS_WBO_CREATE_REQUEST_EXCEPTION"
                call_error_message = exception_text
            new_request = str(result.get("REQUEST") or "").strip()
        except Exception as exc:
            call_error_type, call_error_message = classify_rfc_error(exc)
            call_error_details = format_exception(exc)
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    if not new_request:
        return _error_result(
            env,
            call_error_type or "REQUEST_NOT_CREATED",
            call_error_message or "CTS_WBO_CREATE_REQUEST não devolveu um número de Request.",
            details=call_error_details,
        )

    verification = validate_transport_request(env, new_request)
    if not verification.get("ok"):
        verification["status"] = "PARTIAL_FAILURE"
        verification["request"] = new_request
        verification["message"] = (
            f"A Request {new_request} foi comunicada como criada pelo SAP, mas não foi possível "
            "confirmar o seu estado numa ligação nova e independente."
        )
        return verification

    return {
        "ok": True,
        "status": "REQUEST_CREATED",
        "environment": env,
        "request": new_request,
        "description": verification.get("description") or normalized_description,
        "trtype": verification.get("trtype"),
        "target_system": verification.get("target_system"),
        "state": verification.get("state"),
    }


def list_transport_request_objects(environment: str, request_number: str) -> dict[str, Any]:
    """Read-only: lista os objetos atribuídos a uma Request (e às suas tasks filhas).

    Reproduz o que SE01/SE09 mostram como "objetos da request": entradas em E071 (objetos de
    repositório, ex.: PGMID R3TR OBJECT AGR/PROF) e em E071K (chaves de tabelas de customizing,
    ex.: OBJNAME AGR_DEFINE/AGR_TCODES/AGR_1251 com a TABKEY correspondente). Percorre também as
    tasks filhas (E070 onde STRKORR = request), porque em muitos sistemas as entradas ficam
    registadas numa task do utilizador e não diretamente na request-mãe. Nunca escreve em SAP.
    """
    try:
        env = _assert_transport_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), "ENVIRONMENT_BLOCKED", str(exc))

    request_clean = str(request_number or "").strip().upper()
    if not request_clean:
        return _error_result(env, "INVALID_INPUT", "Informe o número da Request de transporte.")

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for_env(env)
    except Exception as exc:
        return _error_result(env, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    guard = make_write_guard(TRANSPORT_ALLOWED_FUNCTIONS, TRANSPORT_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")

        task_rows = read_table(
            connection,
            guard,
            table_name="E070",
            fields=["TRKORR"],
            options=make_option_eq("STRKORR", request_clean),
            rowcount=0,
        )
        korr_numbers = [request_clean] + sorted({row[0].strip() for row in task_rows if row[0].strip()})

        objects: list[dict[str, Any]] = []
        table_keys: list[dict[str, Any]] = []
        for korr in korr_numbers:
            for row in read_table(
                connection,
                guard,
                table_name="E071",
                fields=["TRKORR", "PGMID", "OBJECT", "OBJ_NAME", "OBJFUNC"],
                options=make_option_eq("TRKORR", korr),
                rowcount=0,
            ):
                objects.append(
                    {
                        "request_or_task": row[0],
                        "pgmid": row[1],
                        "object": row[2],
                        "obj_name": row[3],
                        "objfunc": row[4],
                    }
                )
            for row in read_table(
                connection,
                guard,
                table_name="E071K",
                fields=["TRKORR", "PGMID", "OBJECT", "OBJNAME", "TABKEY"],
                options=make_option_eq("TRKORR", korr),
                rowcount=0,
            ):
                table_keys.append(
                    {
                        "request_or_task": row[0],
                        "pgmid": row[1],
                        "object": row[2],
                        "objname": row[3],
                        "tabkey": row[4],
                    }
                )
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, error_type, message, details=format_exception(exc))
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

    return {
        "ok": True,
        "status": "OBJECTS_LISTED",
        "environment": env,
        "request": request_clean,
        "tasks_checked": korr_numbers,
        "objects": objects,
        "objects_count": len(objects),
        "table_keys": table_keys,
        "table_keys_count": len(table_keys),
    }
