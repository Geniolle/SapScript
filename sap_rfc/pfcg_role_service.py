from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from sap_agent.safety import SafetyGuard


REQUIRED_ENV_VARS = [
    "SAP_PRD_USER",
    "SAP_PRD_PASSWD",
    "SAP_PRD_ASHOST",
    "SAP_PRD_SYSNR",
    "SAP_PRD_CLIENT",
]
ROLE_NAME_RE = re.compile(r"^[A-Z0-9_/\-:]+$")
DELIMITER = "|"
SYSTEM_NAME = "PRD"
ALLOWED_FUNCTIONS = ("RFC_PING", "RFC_READ_TABLE")
ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TEXTS")


def find_project_root() -> Path:
    explicit = str(os.getenv("SAP_SCRIPT_PROJECT_DIR", "") or "").strip()
    if explicit:
        path = Path(explicit).resolve()
        if (path / ".env.example").exists():
            return path

    current = Path(__file__).resolve().parent
    for candidate in [current, *current.parents]:
        if (candidate / ".env.example").exists():
            return candidate
    raise RuntimeError("Não foi possível localizar a raiz do projeto.")


def load_project_env(project_root: Path) -> None:
    from dotenv import load_dotenv

    load_dotenv(project_root / ".env", override=False)


def validate_role_name(role_name: str) -> str:
    normalized = str(role_name or "").strip().upper()
    if not normalized:
        raise ValueError("Entrada inválida: informe um nome de função/perfil PFCG.")
    if not ROLE_NAME_RE.fullmatch(normalized):
        raise ValueError("Entrada inválida: use apenas A-Z, 0-9, _, -, / ou :.")
    return normalized


def format_exception(exc: BaseException) -> str:
    parts: list[str] = [exc.__class__.__name__]
    for attr in ("key", "code", "message"):
        value = getattr(exc, attr, None)
        if value:
            parts.append(f"{attr}={value}")
    text = str(exc).strip()
    if text:
        parts.append(text)
    return " | ".join(parts)


def classify_import_error(exc: BaseException) -> tuple[str, str]:
    if isinstance(exc, ModuleNotFoundError) and getattr(exc, "name", "") == "pyrfc":
        return "PYRFC_UNAVAILABLE", "PyRFC não instalado."

    text = f"{exc.__class__.__name__} {exc}".lower()
    sdk_markers = [
        "sapnwrfc",
        "dll load failed",
        "cannot open shared object file",
        "library not found",
        "netweaver",
    ]
    if any(marker in text for marker in sdk_markers):
        return "SAP_NWRFC_SDK_UNAVAILABLE", "SAP NetWeaver RFC SDK não disponível/configurado."

    return "PYRFC_IMPORT_ERROR", "Falha ao carregar PyRFC ou o SAP NetWeaver RFC SDK."


def classify_rfc_error(exc: BaseException) -> tuple[str, str]:
    text = f"{exc.__class__.__name__} {getattr(exc, 'message', '')} {exc}".lower()
    name = exc.__class__.__name__.lower()

    if "timeout" in text or "timed out" in text:
        return "RFC_TIMEOUT", "Timeout na ligação RFC."
    if "communicationerror" in name or "hostname" in text or "host" in text or "service" in text:
        return "RFC_COMMUNICATION_ERROR", "Hostname/servidor SAP inacessível ou problema de rede."
    if "logonerror" in name or "logon" in text:
        if "client" in text:
            return "RFC_LOGON_CLIENT_ERROR", "Cliente SAP incorreto ou não acessível."
        if "password" in text or "name or password is incorrect" in text or "senha" in text:
            return "RFC_LOGON_CREDENTIAL_ERROR", "Credencial incorreta."
        if "locked" in text or "block" in text or "bloque" in text:
            return "RFC_LOGON_USER_LOCKED", "Utilizador SAP bloqueado."
        return "RFC_LOGON_ERROR", "Erro de autenticação/logon SAP."
    if is_authorization_error(exc):
        return "RFC_AUTHORIZATION_ERROR", "Falta de autorização RFC ou de leitura."
    if "externalruntimeerror" in name or "sapnwrfc" in text or "sdk" in text:
        return "SAP_NWRFC_SDK_UNAVAILABLE", "SAP NetWeaver RFC SDK não disponível/configurado."
    return "RFC_ERROR", "Falha RFC não classificada automaticamente."


def is_authorization_error(exc: BaseException) -> bool:
    text = f"{exc.__class__.__name__} {getattr(exc, 'message', '')} {exc}".lower()
    markers = [
        "authorization",
        "authorisation",
        "not authorized",
        "not authorised",
        "sem autorização",
        "não autorizado",
        "nao autorizado",
        "s_tabu_dis",
        "s_rfc",
    ]
    return any(marker in text for marker in markers)


def build_connection_params() -> dict[str, str]:
    missing = [name for name in REQUIRED_ENV_VARS if not os.getenv(name, "").strip()]
    if missing:
        raise RuntimeError(f"Variáveis obrigatórias ausentes: {', '.join(missing)}")

    return {
        "user": os.environ["SAP_PRD_USER"],
        "passwd": os.environ["SAP_PRD_PASSWD"],
        "ashost": os.environ["SAP_PRD_ASHOST"],
        "sysnr": os.environ["SAP_PRD_SYSNR"],
        "client": os.environ["SAP_PRD_CLIENT"],
        "lang": os.getenv("SAP_PRD_LANG", "PT").strip() or "PT",
    }


def make_read_only_guard() -> SafetyGuard:
    return SafetyGuard.build(
        allow_write_operations=False,
        allowed_functions=ALLOWED_FUNCTIONS,
        allowed_tables=ALLOWED_TABLES,
    )


def make_option_eq(field: str, value: str) -> list[dict[str, str]]:
    return [{"TEXT": f"{field} = '{value}'"}]


def parse_rfc_table_rows(result: dict[str, Any], expected_columns: int) -> list[list[str]]:
    rows = []
    for item in result.get("DATA", []) or []:
        wa = str(item.get("WA", "") or "")
        parts = [part.strip() for part in wa.split(DELIMITER)]
        if len(parts) < expected_columns:
            parts += [""] * (expected_columns - len(parts))
        rows.append(parts[:expected_columns])
    return rows


def normalize_spras(value: str) -> str | None:
    lang = value.strip().upper()
    if lang in {"P", "PT", "PTBR", "PT-PT", "PT-BR"}:
        return "PT"
    if lang in {"E", "EN", "ENUS", "EN-GB", "EN-US"}:
        return "EN"
    return lang or None


def choose_best_text(rows: list[dict[str, str]]) -> dict[str, str] | None:
    if not rows:
        return None

    scored: list[tuple[int, dict[str, str]]] = []
    for row in rows:
        lang = normalize_spras(row.get("SPRAS", "") or "") or ""
        score = 2
        if lang == "PT":
            score = 0
        elif lang == "EN":
            score = 1
        scored.append((score, row))

    scored.sort(key=lambda item: item[0])
    return scored[0][1]


def _error_result(role_name: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "role": role_name,
        "error_type": error_type,
        "message": message,
        "system": SYSTEM_NAME,
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _read_table(
    connection: Any,
    guard: SafetyGuard,
    *,
    table_name: str,
    fields: list[str],
    options: list[dict[str, str]],
    rowcount: int,
) -> list[list[str]]:
    guard.assert_table_allowed(table_name)
    guard.assert_function_allowed("RFC_READ_TABLE")
    result = connection.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table_name,
        DELIMITER=DELIMITER,
        FIELDS=[{"FIELDNAME": field} for field in fields],
        OPTIONS=options,
        ROWCOUNT=rowcount,
    )
    return parse_rfc_table_rows(dict(result or {}), expected_columns=len(fields))


def analyze_pfcg_role_prd(role_name: str) -> dict[str, Any]:
    normalized_role = str(role_name or "").strip().upper()

    try:
        normalized_role = validate_role_name(role_name)
    except ValueError as exc:
        return _error_result(normalized_role, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params()
    except Exception as exc:
        return _error_result(normalized_role, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(normalized_role, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard()
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
            define_rows = _read_table(
                connection,
                guard,
                table_name="AGR_DEFINE",
                fields=["AGR_NAME"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=5,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    normalized_role,
                    "AGR_DEFINE_AUTHORIZATION_ERROR",
                    "Não foi possível determinar se a função existe: sem autorização para consultar AGR_DEFINE.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(normalized_role, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        if not define_rows:
            return {
                "ok": True,
                "status": "NAO_EXISTE",
                "role": normalized_role,
                "description": None,
                "language": None,
                "system": SYSTEM_NAME,
                "client": params["client"],
            }

        description = None
        language = None
        warning = None
        try:
            text_rows_raw = _read_table(
                connection,
                guard,
                table_name="AGR_TEXTS",
                fields=["AGR_NAME", "SPRAS", "TEXT"],
                options=make_option_eq("AGR_NAME", normalized_role),
                rowcount=20,
            )
            text_rows = [
                {"AGR_NAME": row[0], "SPRAS": row[1], "TEXT": row[2]}
                for row in text_rows_raw
            ]
            best = choose_best_text(text_rows)
            if best and str(best.get("TEXT", "")).strip():
                description = str(best["TEXT"]).strip()
                language = normalize_spras(str(best.get("SPRAS", "") or ""))
        except Exception as exc:
            if is_authorization_error(exc):
                warning = "Sem autorização para consultar AGR_TEXTS."
            else:
                warning = classify_rfc_error(exc)[1]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "EXISTE",
            "role": normalized_role,
            "description": description,
            "language": language,
            "system": SYSTEM_NAME,
            "client": params["client"],
        }
        if warning:
            payload["warning"] = warning
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
