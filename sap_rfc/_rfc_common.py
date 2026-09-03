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
KNOWN_ENVIRONMENTS = ("DEV", "QAD", "PRD", "CUA")


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


def build_connection_params_for_env(environment: str) -> dict[str, str]:
    """Generic multi-environment counterpart of `build_connection_params()` (PRD-only).

    Reads `SAP_{ENV}_USER/PASSWD/ASHOST/SYSNR/CLIENT/LANG` for ENV in {DEV, QAD, PRD}.
    """
    env = str(environment or "").strip().upper()
    if env not in KNOWN_ENVIRONMENTS:
        raise ValueError(f"Ambiente desconhecido: {environment}")

    required = [f"SAP_{env}_USER", f"SAP_{env}_PASSWD", f"SAP_{env}_ASHOST", f"SAP_{env}_SYSNR", f"SAP_{env}_CLIENT"]
    missing = [name for name in required if not os.getenv(name, "").strip()]
    if missing:
        raise RuntimeError(f"Variáveis obrigatórias ausentes para {env}: {', '.join(missing)}")

    return {
        "user": os.environ[f"SAP_{env}_USER"],
        "passwd": os.environ[f"SAP_{env}_PASSWD"],
        "ashost": os.environ[f"SAP_{env}_ASHOST"],
        "sysnr": os.environ[f"SAP_{env}_SYSNR"],
        "client": os.environ[f"SAP_{env}_CLIENT"],
        "lang": os.getenv(f"SAP_{env}_LANG", "PT").strip() or "PT",
    }


def build_connection_params_for(environment: str | None = None) -> dict[str, str]:
    """Params de ligacao para o ambiente pedido.

    - Vazio/None ou 'PRD' -> `build_connection_params()` (caminho PRD historico).
    - DEV/QAD/CUA -> `build_connection_params_for_env(env)` (SAP_{ENV}_*).
    """
    env = str(environment or "").strip().upper() or "PRD"
    if env == "PRD":
        return build_connection_params()
    return build_connection_params_for_env(env)


def resolve_target_env(default: str = "PRD") -> str:
    """Ambiente-alvo lido de PFCG_TARGET_ENV (posto pelo worker), validado."""
    env = os.getenv("PFCG_TARGET_ENV", "").strip().upper() or default
    return env if env in KNOWN_ENVIRONMENTS else default


def make_read_only_guard(allowed_tables: tuple[str, ...]) -> SafetyGuard:
    return SafetyGuard.build(
        allow_write_operations=False,
        allowed_functions=("RFC_PING", "RFC_READ_TABLE"),
        allowed_tables=allowed_tables,
    )


def make_write_guard(allowed_functions: tuple[str, ...], allowed_tables: tuple[str, ...]) -> SafetyGuard:
    """Guard for RFC flows that legitimately need to call a whitelisted write function.

    Unlike `make_read_only_guard`, `allow_write_operations=True` here — but this only
    lifts the mutation-keyword block; the explicit `allowed_functions` whitelist still
    fully controls exactly which RFC function names may be called.
    """
    return SafetyGuard.build(
        allow_write_operations=True,
        allowed_functions=allowed_functions,
        allowed_tables=allowed_tables,
    )


def make_option_eq(field: str, value: str) -> list[dict[str, str]]:
    return [{"TEXT": f"{field} = '{value}'"}]


def make_option_in(field: str, values: list[str]) -> list[dict[str, str]]:
    """Build RFC_READ_TABLE OPTIONS rows for `field IN (values)` using OR-joined equality clauses.

    Each value produces its own OPTIONS row, so a single condition never risks exceeding the
    72-character line limit of the underlying function module.
    """
    options: list[dict[str, str]] = []
    for index, value in enumerate(values):
        prefix = "OR " if index > 0 else ""
        options.append({"TEXT": f"{prefix}{field} = '{value}'"})
    return options


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
        lang = normalize_spras(row.get("SPRAS", "") or row.get("SPRSL", "") or "") or ""
        score = 2
        if lang == "PT":
            score = 0
        elif lang == "EN":
            score = 1
        scored.append((score, row))

    scored.sort(key=lambda item: item[0])
    return scored[0][1]


def read_table(
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


def role_exists(connection: Any, guard: SafetyGuard, role_name: str) -> bool:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_DEFINE",
        fields=["AGR_NAME"],
        options=make_option_eq("AGR_NAME", role_name),
        rowcount=1,
    )
    return bool(rows)


def fetch_composite_members(connection: Any, guard: SafetyGuard, role_name: str) -> list[str]:
    """Return the CHILD_AGR members if `role_name` is a composite role (Sammelrolle), else []."""
    rows = read_table(
        connection,
        guard,
        table_name="AGR_AGRS",
        fields=["AGR_NAME", "CHILD_AGR"],
        options=make_option_eq("AGR_NAME", role_name),
        rowcount=0,
    )
    return sorted({row[1].strip() for row in rows if row[1].strip()})
