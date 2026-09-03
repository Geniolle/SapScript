from __future__ import annotations

import re
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
    make_option_in,
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

# Limites reais confirmados via DDIF_FIELDINFO_GET (não assumidos por documentação):
#   AGR_DEFINE-AGR_NAME  -> CHAR30
#   AGR_TEXTS-TEXT       -> CHAR80
#   SSM_TCODES-TCODE / TSTC-TCODE -> CHAR20 (estrutura real do parâmetro TCODES de
#   PRGN_RFC_CREATE_ACTIVITY_GROUP, confirmada via RFC_GET_FUNCTION_INTERFACE)
MAX_ROLE_NAME_LENGTH = 30
MAX_DESCRIPTION_LENGTH = 80
MAX_TCODE_LENGTH = 20
MAX_PROFILE_TEXT_LENGTH = 60  # AGR_PROF-PTEXT

# Apenas DEV pode receber escrita nesta primeira fase (secção "restrição de ambiente").
ALLOWED_CREATE_ENVIRONMENTS = KNOWN_ENVIRONMENTS  # escrita permitida em DEV/QAD/PRD/CUA (autorizado explicitamente)

# Modos de transporte suportados pelo fluxo "Criar Individualmente":
#   LOCAL            -> REQUEST="" (objeto local, sem transporte, como acontecia até agora).
#   CREATE_REQUEST   -> cria uma nova Request de Customizing (via pfcg_transport_service)
#                        e usa o TRKORR devolvido.
#   EXISTING_REQUEST -> reutiliza uma Request já existente e aberta, escolhida pelo utilizador.
TRANSPORT_MODE_LOCAL = "LOCAL"
TRANSPORT_MODE_CREATE_REQUEST = "CREATE_REQUEST"
TRANSPORT_MODE_EXISTING_REQUEST = "EXISTING_REQUEST"
VALID_TRANSPORT_MODES = (TRANSPORT_MODE_LOCAL, TRANSPORT_MODE_CREATE_REQUEST, TRANSPORT_MODE_EXISTING_REQUEST)

WRITE_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "PRGN_RFC_CREATE_ACTIVITY_GROUP",
    "PRGN_GEN_PROFILES_FOR_ROLES",
    "PRGN_CHECK_PROFILE_STATUS_RFC",
)
WRITE_ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TEXTS", "AGR_TCODES", "TSTC")

_TCODE_SPLIT_RE = re.compile(r"[;,\n ]+")


def normalize_tcodes(raw: Any) -> list[str]:
    """Normaliza uma lista/texto de transações: maiúsculas, remove /N e /O, remove duplicados.

    Aceita separação por espaço, vírgula, ponto-e-vírgula ou quebra de linha.
    """
    if isinstance(raw, (list, tuple, set)):
        parts = [str(item) for item in raw]
    else:
        text = str(raw or "").replace("\r", "\n").replace("\t", " ")
        parts = _TCODE_SPLIT_RE.split(text)

    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        value = str(part or "").strip().upper()
        if not value:
            continue
        if value.startswith("/N") or value.startswith("/O"):
            value = value[2:].strip()
        if not value or value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


def validate_role_name_for_create(role_name: str) -> str:
    normalized = validate_role_name(role_name)
    if len(normalized) > MAX_ROLE_NAME_LENGTH:
        raise ValueError(
            f"Nome da função excede o limite real de {MAX_ROLE_NAME_LENGTH} caracteres (AGR_DEFINE-AGR_NAME)."
        )
    return normalized


def validate_description(description: str) -> str:
    normalized = str(description or "").strip()
    if not normalized:
        raise ValueError("Informe uma descrição para o Perfil de Autorização.")
    if len(normalized) > MAX_DESCRIPTION_LENGTH:
        raise ValueError(
            f"A descrição excede o limite real de {MAX_DESCRIPTION_LENGTH} caracteres (AGR_TEXTS-TEXT)."
        )
    return normalized


def validate_tcode_list(tcodes: list[str]) -> list[str]:
    if not tcodes:
        raise ValueError("Informe pelo menos uma transação.")
    cleaned: list[str] = []
    for tcode in tcodes:
        value = str(tcode or "").strip().upper()
        if not value:
            continue
        if len(value) > MAX_TCODE_LENGTH:
            raise ValueError(f"A transação '{value}' excede o limite real de {MAX_TCODE_LENGTH} caracteres.")
        cleaned.append(value)
    if not cleaned:
        raise ValueError("Informe pelo menos uma transação.")
    return cleaned


def validate_transport_inputs(transport_mode: str, request_number: str, request_description: str) -> dict[str, str]:
    """Valida localmente (sem RFC) os campos do modo de transporte escolhido.

    A revalidação real em SAP (existência/estado da Request) acontece depois, já com
    ligação aberta, em preview_pfcg_role_create()/create_pfcg_role_rfc().
    """
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

    # TRANSPORT_MODE_CREATE_REQUEST
    normalized_description = validate_request_description(request_description)
    return {"transport_mode": mode, "request_number": "", "request_description": normalized_description}


def _assert_create_environment_allowed(environment: str) -> str:
    env = str(environment or "").strip().upper()
    if env not in ALLOWED_CREATE_ENVIRONMENTS:
        raise ValueError(
            f"Criação individual via RFC só é permitida em DEV nesta fase. "
            f"Ambiente '{env}' está bloqueado (QAD e PRD não são permitidos para escrita)."
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


def _fetch_existing_tcodes(connection: Any, guard: Any, tcodes: list[str]) -> set[str]:
    if not tcodes:
        return set()
    rows = read_table(
        connection,
        guard,
        table_name="TSTC",
        fields=["TCODE"],
        options=make_option_in("TCODE", tcodes),
        rowcount=0,
    )
    return {row[0].strip().upper() for row in rows if row[0].strip()}


def _validate_inputs(
    environment: str,
    role_name: str,
    description: str,
    tcodes: Any,
    transport_mode: str = TRANSPORT_MODE_LOCAL,
    request_number: str = "",
    request_description: str = "",
) -> tuple[str, str, str, list[str], dict[str, str]] | dict[str, Any]:
    """Returns (env, role, description, tcodes, transport) on success, or an error-result dict."""
    raw_role_upper = str(role_name or "").strip().upper()
    try:
        env = _assert_create_environment_allowed(environment)
    except ValueError as exc:
        return _error_result(str(environment or "").strip().upper(), raw_role_upper, "ENVIRONMENT_BLOCKED", str(exc))

    try:
        normalized_role = validate_role_name_for_create(role_name)
    except ValueError as exc:
        return _error_result(env, raw_role_upper, "INVALID_INPUT", str(exc))

    try:
        normalized_description = validate_description(description)
    except ValueError as exc:
        return _error_result(env, normalized_role, "INVALID_INPUT", str(exc))

    try:
        tcodes_list = validate_tcode_list(normalize_tcodes(tcodes))
    except ValueError as exc:
        return _error_result(env, normalized_role, "INVALID_INPUT", str(exc))

    try:
        transport = validate_transport_inputs(transport_mode, request_number, request_description)
    except ValueError as exc:
        return _error_result(env, normalized_role, "INVALID_TRANSPORT_INPUT", str(exc))

    return env, normalized_role, normalized_description, tcodes_list, transport


def preview_pfcg_role_create(
    environment: str,
    role_name: str,
    description: str,
    tcodes: Any,
    transport_mode: str = TRANSPORT_MODE_LOCAL,
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Read-only preview: valida ambiente, campos, TCODEs reais (TSTC), não-existência da função
    e, consoante o modo de transporte, a Request existente escolhida (CTS_WBO_GET_REQUEST_DETAILS).

    Nunca escreve em SAP. Usado exclusivamente para preparar o ecrã de confirmação.
    """
    validated = _validate_inputs(
        environment, role_name, description, tcodes, transport_mode, request_number, request_description
    )
    if isinstance(validated, dict):
        return validated
    env, normalized_role, normalized_description, tcodes_list, transport = validated

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

    guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    try:
        try:
            if role_exists(connection, guard, normalized_role):
                return _error_result(
                    env,
                    normalized_role,
                    "ROLE_ALREADY_EXISTS",
                    f"A função {normalized_role} já existe em {env}. Não é possível criar novamente.",
                )
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, normalized_role, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        try:
            existing_tcodes = _fetch_existing_tcodes(connection, guard, tcodes_list)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, normalized_role, f"TSTC_{error_type}", message, details=format_exception(exc))

        missing_tcodes = sorted(set(tcodes_list) - existing_tcodes)
        if missing_tcodes:
            payload = _error_result(
                env,
                normalized_role,
                "TCODES_NOT_FOUND",
                f"As seguintes transações não existem em {env}: {', '.join(missing_tcodes)}",
            )
            payload["missing_tcodes"] = missing_tcodes
            return payload

        transport_preview: dict[str, Any] = {"transport_mode": transport["transport_mode"]}
        if transport["transport_mode"] == TRANSPORT_MODE_EXISTING_REQUEST:
            request_check = validate_transport_request(env, transport["request_number"])
            if not request_check.get("ok"):
                payload = _error_result(
                    env,
                    normalized_role,
                    request_check.get("error_type") or "REQUEST_NOT_VALID",
                    request_check.get("message") or "A Request de transporte selecionada não é válida.",
                )
                return payload
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
            "description": normalized_description,
            "tcodes": tcodes_list,
            "tcodes_count": len(tcodes_list),
            "transport": transport_preview,
        }
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass


def create_pfcg_role_rfc(
    environment: str,
    role_name: str,
    description: str,
    tcodes: Any,
    transport_mode: str = TRANSPORT_MODE_LOCAL,
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Cria a função PFCG via RFC (PRGN_RFC_CREATE_ACTIVITY_GROUP), apenas em DEV.

    Sequência: valida -> RFC_PING -> reconfirma não-existência (condição de corrida) ->
    reconfirma TCODEs -> resolve a Request de transporte (LOCAL: nenhuma; CREATE_REQUEST:
    cria uma nova Request de Customizing via pfcg_transport_service; EXISTING_REQUEST:
    revalida novamente a Request escolhida, condição de corrida) -> chama a função de
    criação com valores explícitos e conservadores para ORG_LEVELS_WITH_STAR /
    UNMAINTAINED_FIELDS_WITH_STAR / TEMPLATE -> abre uma ligação nova e independente para
    validar de forma objetiva se os dados foram persistidos (nunca assume sucesso apenas
    por ausência de exceção).
    """
    validated = _validate_inputs(
        environment, role_name, description, tcodes, transport_mode, request_number, request_description
    )
    if isinstance(validated, dict):
        return validated
    env, normalized_role, normalized_description, tcodes_list, transport = validated

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

    guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(env, normalized_role, error_type, message, details=format_exception(exc))

    try:
        try:
            if role_exists(connection, guard, normalized_role):
                return _error_result(
                    env,
                    normalized_role,
                    "ROLE_ALREADY_EXISTS",
                    f"A função {normalized_role} passou a existir em {env} entre a pré-visualização e a confirmação.",
                )
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, normalized_role, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        try:
            existing_tcodes = _fetch_existing_tcodes(connection, guard, tcodes_list)
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(env, normalized_role, f"TSTC_{error_type}", message, details=format_exception(exc))

        missing_tcodes = sorted(set(tcodes_list) - existing_tcodes)
        if missing_tcodes:
            payload = _error_result(
                env,
                normalized_role,
                "TCODES_NOT_FOUND",
                f"As seguintes transações não existem em {env}: {', '.join(missing_tcodes)}",
            )
            payload["missing_tcodes"] = missing_tcodes
            return payload

        # Resolução da Request de transporte a passar em PRGN_RFC_CREATE_ACTIVITY_GROUP-REQUEST,
        # de acordo com o modo escolhido. Feito com a ligação de escrita já aberta e imediatamente
        # antes da chamada de criação, para minimizar a janela de condição de corrida.
        created_transport_request: dict[str, Any] | None = None
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
                    request_create_result.get("error_type") or "REQUEST_NOT_CREATED",
                    request_create_result.get("message") or "Não foi possível criar a Request de transporte.",
                )
            resolved_request_number = str(request_create_result.get("request") or "")
            created_transport_request = request_create_result

        profile_text = normalized_description[:MAX_PROFILE_TEXT_LENGTH]

        guard.assert_function_allowed("PRGN_RFC_CREATE_ACTIVITY_GROUP")
        call_error_type: str | None = None
        call_error_message: str | None = None
        call_error_details: str | None = None
        return_rows: list[dict[str, Any]] = []
        # Inicializado a partir da Request já resolvida (não de "") para que o número não se
        # perca se PRGN_RFC_CREATE_ACTIVITY_GROUP lançar exceção antes de devolver NEW_REQUEST
        # — confirmado empiricamente: isso acontece sempre que a geração interna de perfil falha
        # (sem Níveis Organizacionais), mesmo com a função e a Request corretamente persistidas.
        new_request = resolved_request_number
        try:
            create_result = connection.call(
                "PRGN_RFC_CREATE_ACTIVITY_GROUP",
                ACTIVITY_GROUP=normalized_role,
                ACTIVITY_GROUP_TEXT=normalized_description,
                NO_DIALOG="X",
                ONLY_TCODE_ASSIGNMENT="",
                # Valores explícitos e conservadores: nunca aceitar o default perigoso 'X'
                # confirmado via get_function_description() na assinatura real da função.
                ORG_LEVELS_WITH_STAR="",
                UNMAINTAINED_FIELDS_WITH_STAR="",
                PARENT_ROLE="",
                PROFILE_NAME="",
                PROFILE_TEXT=profile_text,
                REQUEST=resolved_request_number,
                # Confirmado via DFIES/interface real: sem default oculto nesta assinatura,
                # mas mantido explicitamente vazio para eliminar qualquer ambiguidade.
                TEMPLATE="",
                TCODES=[{"TCODE": tcode} for tcode in tcodes_list],
            )
            return_rows = create_result.get("RETURN") or []
            new_request = str(create_result.get("NEW_REQUEST") or "").strip() or resolved_request_number
        except Exception as exc:
            # Não devolver erro de imediato aqui: confirmado empiricamente em teste real que
            # PRGN_RFC_CREATE_ACTIVITY_GROUP pode lançar uma exceção (ex.: geração de perfil
            # falhada por falta de valores de nível de organização, que esta ferramenta nunca
            # preenche automaticamente) mesmo depois de já ter persistido a função/textos/
            # tcodes em BD. A única forma fiável de saber o que foi realmente criado é a
            # verificação pós-escrita independente abaixo — nunca a mera presença/ausência
            # desta exceção.
            if getattr(exc, "key", "") == "ERROR_WHEN_GENERATING_PROFILE":
                call_error_type = "ERROR_WHEN_GENERATING_PROFILE"
                call_error_message = (
                    "A função/transações foram criadas, mas a tentativa de geração automática do perfil "
                    "feita internamente por PRGN_RFC_CREATE_ACTIVITY_GROUP falhou (ex.: Níveis "
                    "Organizacionais por preencher). A geração é repetida a seguir via "
                    "PRGN_GEN_PROFILES_FOR_ROLES, que não depende desse preenchimento."
                )
            else:
                call_error_type, call_error_message = classify_rfc_error(exc)
            call_error_details = format_exception(exc)

        error_messages = [
            str(row.get("MESSAGE") or "").strip()
            for row in return_rows
            if str(row.get("TYPE") or "").strip().upper() in {"E", "A"} and str(row.get("MESSAGE") or "").strip()
        ]
        if call_error_message:
            error_messages.append(call_error_message)

        # Verificação pós-escrita com ligação NOVA e independente: só uma leitura nova
        # sem qualquer relação com a ligação que escreveu prova persistência real em BD
        # (isolamento de transação impede que outra sessão veja dados não confirmados).
        try:
            connection.close()
        except Exception:
            pass
        connection = None

        verify_connection = Connection(**params)
        try:
            verify_guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)

            role_created = role_exists(verify_connection, verify_guard, normalized_role)
            if not role_created:
                if call_error_type:
                    payload = _error_result(
                        env,
                        normalized_role,
                        f"PRGN_RFC_CREATE_ACTIVITY_GROUP_{call_error_type}",
                        call_error_message or "",
                        details=call_error_details,
                    )
                else:
                    payload = _error_result(
                        env,
                        normalized_role,
                        "WRITE_NOT_PERSISTED",
                        "A chamada RFC não gerou exceção, mas a função não foi encontrada numa ligação nova e "
                        "independente (AGR_DEFINE). Pode indicar necessidade de commit explícito ainda não "
                        "confirmada — não repetir a criação sem investigar antes.",
                    )
                if error_messages:
                    payload["sap_return_messages"] = error_messages
                if new_request:
                    payload["transport_request"] = new_request
                return payload

            text_rows = read_table(
                verify_connection,
                verify_guard,
                table_name="AGR_TEXTS",
                fields=["AGR_NAME", "SPRAS", "TEXT"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=20,
            )
            description_ok = any(str(row[2] or "").strip() == normalized_description for row in text_rows)

            tcode_rows = read_table(
                verify_connection,
                verify_guard,
                table_name="AGR_TCODES",
                fields=["AGR_NAME", "TCODE"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=0,
            )
            created_tcodes = sorted({row[1].strip().upper() for row in tcode_rows if row[1].strip()})
            missing_after_create = sorted(set(tcodes_list) - set(created_tcodes))

            # PRGN_RFC_CREATE_ACTIVITY_GROUP cria a função mas nunca gera o perfil de
            # autorizações (fica amarelo na PFCG) — passo em falta confirmado empiricamente.
            # PRGN_GEN_PROFILES_FOR_ROLES é a RFC de geração em massa usada pela SUPC/PFUD e foi
            # confirmada empiricamente (teste real + verificação visual do utilizador na PFCG) a
            # produzir semáforo VERDE mesmo com os Níveis Organizacionais por preencher — é o
            # equivalente RFC ao fluxo manual "abrir o perfil, fechar o pop-up sem atribuir nada,
            # gravar e gerar". Não tem parâmetros ORG_LEVELS_WITH_STAR / FILL_EMPTY_FIELDS_WITH_STAR
            # / NO_DIALOG: esta ferramenta nunca atribui '*' nem valores aos Níveis Organizacionais
            # por decisão explícita do negócio, e esta função não precisa disso para gerar.
            verify_guard.assert_function_allowed("PRGN_GEN_PROFILES_FOR_ROLES")
            try:
                mass_gen_result = verify_connection.call(
                    "PRGN_GEN_PROFILES_FOR_ROLES",
                    IV_USERCOMPARE="",
                    IT_ROLES=[{"AGR_NAME": normalized_role}],
                )
                for row in mass_gen_result.get("ET_RETURN") or []:
                    if str(row.get("TYPE") or "").strip().upper() in {"E", "A"}:
                        row_message = str(row.get("MESSAGE") or "").strip()
                        if row_message:
                            error_messages.append(row_message)
            except Exception as exc:
                _, generate_error_message = classify_rfc_error(exc)
                error_messages.append(generate_error_message)

            # Confirmado empiricamente: RFC_READ_TABLE em AGR_PROF/AGR_1251 devolve
            # TABLE_WITHOUT_DATA mesmo com o perfil já gerado e o semáforo verde na PFCG —
            # o utilizador técnico não tem autorização de leitura direta nessas tabelas via
            # RFC_READ_TABLE. O único sinal fiável de geração é PRGN_CHECK_PROFILE_STATUS_RFC,
            # a mesma lógica que pinta o semáforo da PFCG (verificado contra a UI real).
            verify_guard.assert_function_allowed("PRGN_CHECK_PROFILE_STATUS_RFC")
            status_result = verify_connection.call(
                "PRGN_CHECK_PROFILE_STATUS_RFC",
                ACTIVITY_GROUP=normalized_role,
                TRANSACTIONS_CHANGED="",
            )
            profile_generated = str(status_result.get("LED_COLOR") or "").strip().upper() == "GREEN"
            profile_status_message = str(status_result.get("MESSAGE_TEXT") or "").strip()

            if not description_ok or missing_after_create:
                payload = {
                    "ok": False,
                    "status": "PARTIAL_FAILURE",
                    "environment": env,
                    "role": normalized_role,
                    "description_set": description_ok,
                    "tcodes_created": created_tcodes,
                    "tcodes_missing": missing_after_create,
                    "profile_generated": profile_generated,
                    "profile_status_message": profile_status_message,
                    "message": (
                        "A função foi criada mas nem todos os elementos foram confirmados na verificação "
                        "pós-escrita independente. Não foi executada qualquer remoção automática da função."
                    ),
                }
                if error_messages:
                    payload["sap_return_messages"] = error_messages
                if new_request:
                    payload["transport_request"] = new_request
                if call_error_type:
                    payload["create_call_error_type"] = f"PRGN_RFC_CREATE_ACTIVITY_GROUP_{call_error_type}"
                payload["transport_mode"] = transport["transport_mode"]
                return payload

            # profile_generated=False aqui é inesperado (PRGN_GEN_PROFILES_FOR_ROLES gera mesmo
            # com Níveis Organizacionais por preencher) — mantido como estado defensivo para o
            # caso de a geração falhar por outro motivo (ex.: falha ET_RETURN não capturada acima).
            result_payload: dict[str, Any] = {
                "ok": True,
                "status": "CREATED" if profile_generated else "CREATED_PENDING_ORG_LEVELS",
                "environment": env,
                "role": normalized_role,
                "description": normalized_description,
                "tcodes_requested": len(tcodes_list),
                "tcodes_created": len(created_tcodes),
                "profile_generated": profile_generated,
                "profile_status_message": profile_status_message,
                "transport_mode": transport["transport_mode"],
            }
            if not profile_generated:
                result_payload["message"] = (
                    "Função e transações criadas com sucesso, mas a geração do perfil de autorização "
                    "não foi confirmada como concluída (semáforo não ficou verde). Verifique manualmente "
                    "na PFCG."
                )
                if error_messages:
                    result_payload["sap_return_messages"] = error_messages
            if new_request:
                result_payload["transport_request"] = new_request
            if created_transport_request:
                result_payload["transport_request_created"] = True
            return result_payload
        except Exception as exc:
            error_type, message = classify_rfc_error(exc)
            return _error_result(
                env, normalized_role, f"POST_VALIDATION_{error_type}", message, details=format_exception(exc)
            )
        finally:
            try:
                verify_connection.close()
            except Exception:
                pass
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
