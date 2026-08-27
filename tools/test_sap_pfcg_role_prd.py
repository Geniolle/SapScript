from __future__ import annotations

import argparse
import os
import re
import sys
import traceback
from pathlib import Path
from typing import Any


REQUIRED_ENV_VARS = [
    "SAP_PRD_USER",
    "SAP_PRD_PASSWD",
    "SAP_PRD_ASHOST",
    "SAP_PRD_SYSNR",
    "SAP_PRD_CLIENT",
]
ROLE_NAME_RE = re.compile(r"^[A-Z0-9_/\-:]+$")
DELIMITER = "|"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Diagnóstico read-only de existência de função PFCG no SAP PRD."
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Mostra traceback técnico completo em caso de erro.",
    )
    return parser.parse_args()


def print_header() -> None:
    print("=" * 60)
    print(" SAP PRD - ANÁLISE PFCG")
    print("=" * 60)
    print()


def debug_trace(debug: bool) -> None:
    if debug:
        print()
        traceback.print_exc()


def find_project_root() -> Path:
    current = Path(__file__).resolve().parent
    for candidate in [current, *current.parents]:
        if (candidate / ".env.example").exists():
            return candidate
    raise RuntimeError("Não foi possível localizar a raiz do projeto a partir de tools/.")


def load_project_env(project_root: Path, debug: bool) -> int:
    try:
        from dotenv import load_dotenv
    except Exception:
        print("❌ python-dotenv não está disponível. Instale a dependência para carregar o .env.")
        debug_trace(debug)
        return 2

    env_path = project_root / ".env"
    load_dotenv(env_path, override=False)
    return 0


def env_status(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        return "AUSENTE"
    if name == "SAP_PRD_PASSWD":
        return "******** / configurado"
    return value


def print_prevalidation() -> list[str]:
    lang = os.getenv("SAP_PRD_LANG", "PT").strip() or "PT"
    print("Pré-validação:")
    print(f"  SAP_PRD_USER    : {env_status('SAP_PRD_USER')}")
    print(f"  SAP_PRD_PASSWD  : {env_status('SAP_PRD_PASSWD')}")
    print(f"  SAP_PRD_ASHOST  : {env_status('SAP_PRD_ASHOST')}")
    print(f"  SAP_PRD_SYSNR   : {env_status('SAP_PRD_SYSNR')}")
    print(f"  SAP_PRD_CLIENT  : {env_status('SAP_PRD_CLIENT')}")
    print(f"  SAP_PRD_LANG    : {lang}")
    print()
    return [name for name in REQUIRED_ENV_VARS if not os.getenv(name, "").strip()]


def print_runtime_config() -> None:
    lang = os.getenv("SAP_PRD_LANG", "PT").strip() or "PT"
    print("Configuração:")
    print(f"  User    : {os.getenv('SAP_PRD_USER', '').strip() or 'AUSENTE'}")
    print(f"  Host    : {os.getenv('SAP_PRD_ASHOST', '').strip() or 'AUSENTE'}")
    print(f"  Sysnr   : {os.getenv('SAP_PRD_SYSNR', '').strip() or 'AUSENTE'}")
    print(f"  Client  : {os.getenv('SAP_PRD_CLIENT', '').strip() or 'AUSENTE'}")
    print(f"  Lang    : {lang}")
    print("  Password: ********")
    print()


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


def classify_import_error(exc: BaseException) -> str:
    if isinstance(exc, ModuleNotFoundError) and getattr(exc, "name", "") == "pyrfc":
        return "PyRFC não instalado."

    text = f"{exc.__class__.__name__} {exc}".lower()
    sdk_markers = [
        "sapnwrfc",
        "dll load failed",
        "cannot open shared object file",
        "library not found",
        "netweaver",
    ]
    if any(marker in text for marker in sdk_markers):
        return "SAP NetWeaver RFC SDK não disponível/configurado."

    return "Falha ao carregar PyRFC ou o SAP NetWeaver RFC SDK."


def classify_rfc_error(exc: BaseException) -> str:
    text = f"{exc.__class__.__name__} {getattr(exc, 'message', '')} {exc}".lower()
    name = exc.__class__.__name__.lower()

    if "timeout" in text or "timed out" in text:
        return "Timeout na ligação RFC."
    if "communicationerror" in name or "hostname" in text or "host" in text or "service" in text:
        return "Hostname/servidor SAP inacessível ou problema de rede."
    if "logonerror" in name or "logon" in text:
        if "client" in text:
            return "Cliente SAP incorreto ou não acessível."
        if "password" in text or "name or password is incorrect" in text or "senha" in text:
            return "Credencial incorreta."
        if "locked" in text or "block" in text or "bloque" in text:
            return "Utilizador SAP bloqueado."
        return "Erro de autenticação/logon SAP."
    if "authorization" in text or "authorisation" in text or "not authorized" in text or "not authorised" in text:
        return "Falta de autorização RFC ou de leitura."
    if "externalruntimeerror" in name or "sapnwrfc" in text or "sdk" in text:
        return "SAP NetWeaver RFC SDK não disponível/configurado."
    return "Falha RFC não classificada automaticamente."


def build_connection_params() -> dict[str, str]:
    return {
        "user": os.environ["SAP_PRD_USER"],
        "passwd": os.environ["SAP_PRD_PASSWD"],
        "ashost": os.environ["SAP_PRD_ASHOST"],
        "sysnr": os.environ["SAP_PRD_SYSNR"],
        "client": os.environ["SAP_PRD_CLIENT"],
        "lang": os.getenv("SAP_PRD_LANG", "PT").strip() or "PT",
    }


def import_pyrfc(debug: bool) -> tuple[Any | None, int]:
    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        print("[1/3] Conexão RFC")
        print(f"❌ {classify_import_error(exc)}")
        if debug:
            print(format_exception(exc))
            debug_trace(debug)
        print()
        print("RESULTADO")
        print("-" * 60)
        print("Status     : ERRO")
        print("Detalhe    : Camada RFC indisponível.")
        print("-" * 60)
        return None, 2

    return Connection, 0


def prompt_role_name() -> tuple[str | None, int]:
    raw = input("Nome da função/perfil PFCG: ")
    role_name = raw.strip().upper()

    if not role_name:
        print("❌ Entrada inválida: informe um nome de função/perfil PFCG.")
        return None, 2

    if not ROLE_NAME_RE.fullmatch(role_name):
        print("❌ Entrada inválida: use apenas A-Z, 0-9, _, -, / ou :.")
        return None, 2

    return role_name, 0


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


def test_rfc_ping(Connection: Any, debug: bool) -> tuple[Any | None, int]:
    print("[1/3] Conexão RFC")
    try:
        connection = Connection(**build_connection_params())
        connection.call("RFC_PING")
    except Exception as exc:
        print(f"❌ {classify_rfc_error(exc)}")
        print(format_exception(exc))
        debug_trace(debug)
        print()
        print("RESULTADO")
        print("-" * 60)
        print("Status     : ERRO")
        print("Detalhe    : Não foi possível estabelecer a conexão RFC com SAP PRD.")
        print(f"Função     : {role_name_or_dash(None)}")
        print("Sistema    : PRD")
        print(f"Cliente    : {os.getenv('SAP_PRD_CLIENT', '').strip() or 'AUSENTE'}")
        print("-" * 60)
        return None, 2

    print("✅ Conectado ao SAP PRD.")
    print()
    return connection, 0


def role_name_or_dash(role_name: str | None) -> str:
    return role_name or "-"


def read_agr_define(connection: Any, role_name: str, debug: bool) -> tuple[bool | None, str, int]:
    print("[2/3] AGR_DEFINE")
    try:
        result = connection.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="AGR_DEFINE",
            DELIMITER=DELIMITER,
            FIELDS=[{"FIELDNAME": "AGR_NAME"}],
            OPTIONS=make_option_eq("AGR_NAME", role_name),
            ROWCOUNT=5,
        )
    except Exception as exc:
        detail = "sem autorização para consultar AGR_DEFINE." if is_authorization_error(exc) else classify_rfc_error(exc)
        print("⚠️ Não foi possível determinar se a função existe:")
        print(f"⚠️ {detail}")
        print(format_exception(exc))
        debug_trace(debug)
        print()
        return None, detail, 3

    rows = parse_rfc_table_rows(result, expected_columns=1)
    if not rows:
        print("❌ Função não encontrada.")
        print()
        return False, "", 0

    print("✅ Função encontrada.")
    print()
    return True, rows[0][0] or role_name, 0


def normalize_spras(value: str) -> str:
    lang = value.strip().upper()
    if lang in {"P", "PT", "PTBR", "PT-PT", "PT-BR"}:
        return "PT"
    if lang in {"E", "EN", "ENUS", "EN-GB", "EN-US"}:
        return "EN"
    return lang or "-"


def choose_best_text(rows: list[dict[str, str]]) -> dict[str, str] | None:
    if not rows:
        return None

    scored: list[tuple[int, dict[str, str]]] = []
    for row in rows:
        lang = normalize_spras(row.get("SPRAS", ""))
        score = 2
        if lang == "PT":
            score = 0
        elif lang == "EN":
            score = 1
        scored.append((score, row))

    scored.sort(key=lambda item: item[0])
    return scored[0][1]


def read_agr_texts(connection: Any, role_name: str, debug: bool) -> tuple[dict[str, str] | None, str]:
    print("[3/3] AGR_TEXTS")
    try:
        result = connection.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="AGR_TEXTS",
            DELIMITER=DELIMITER,
            FIELDS=[
                {"FIELDNAME": "AGR_NAME"},
                {"FIELDNAME": "SPRAS"},
                {"FIELDNAME": "TEXT"},
            ],
            OPTIONS=make_option_eq("AGR_NAME", role_name),
            ROWCOUNT=20,
        )
    except Exception as exc:
        detail = "sem autorização para consultar AGR_TEXTS." if is_authorization_error(exc) else classify_rfc_error(exc)
        print("⚠️ Não foi possível consultar a descrição da função.")
        print(f"⚠️ {detail}")
        print(format_exception(exc))
        debug_trace(debug)
        print()
        return None, detail

    parsed_rows = []
    for parts in parse_rfc_table_rows(result, expected_columns=3):
        parsed_rows.append(
            {
                "AGR_NAME": parts[0],
                "SPRAS": parts[1],
                "TEXT": parts[2],
            }
        )

    best = choose_best_text(parsed_rows)
    if best is None or not best.get("TEXT", "").strip():
        print("⚠️ Nenhuma descrição disponível.")
        print()
        return None, "Descrição não disponível."

    print("✅ Descrição encontrada.")
    print()
    return best, ""


def print_result_exists(role_name: str, description: dict[str, str] | None) -> None:
    print("RESULTADO")
    print("-" * 60)
    print("Status     : EXISTE")
    print(f"Função     : {role_name}")
    if description and description.get("TEXT", "").strip():
        print(f"Descrição  : {description['TEXT'].strip()}")
        print(f"Idioma     : {normalize_spras(description.get('SPRAS', ''))}")
    print("Sistema    : PRD")
    print(f"Cliente    : {os.getenv('SAP_PRD_CLIENT', '').strip() or 'AUSENTE'}")
    print("-" * 60)


def print_result_not_found(role_name: str) -> None:
    print("RESULTADO")
    print("-" * 60)
    print("Status     : NÃO EXISTE")
    print(f"Função     : {role_name}")
    print("Sistema    : PRD")
    print(f"Cliente    : {os.getenv('SAP_PRD_CLIENT', '').strip() or 'AUSENTE'}")
    print("-" * 60)


def print_result_unknown(role_name: str, detail: str) -> None:
    print("RESULTADO")
    print("-" * 60)
    print("Status     : INDETERMINADO")
    print(f"Função     : {role_name}")
    print(f"Detalhe    : {detail}")
    print("Sistema    : PRD")
    print(f"Cliente    : {os.getenv('SAP_PRD_CLIENT', '').strip() or 'AUSENTE'}")
    print("-" * 60)


def safe_close(connection: Any) -> None:
    try:
        if connection is not None:
            connection.close()
    except Exception:
        pass


def main() -> int:
    args = parse_args()
    print_header()

    try:
        project_root = find_project_root()
    except Exception as exc:
        print(f"❌ {exc}")
        debug_trace(args.debug)
        return 2

    load_status = load_project_env(project_root, args.debug)
    if load_status != 0:
        return load_status

    missing_vars = print_prevalidation()
    if missing_vars:
        print("❌ Variáveis obrigatórias ausentes:")
        for name in missing_vars:
            print(f"  - {name}")
        print()
        print("RESULTADO")
        print("-" * 60)
        print("Status     : ERRO")
        print("Detalhe    : Configuração incompleta.")
        print("-" * 60)
        return 2

    print_runtime_config()

    role_name, role_status = prompt_role_name()
    if role_status != 0 or role_name is None:
        print()
        print_result_unknown("-", "Entrada inválida.")
        return role_status

    print("Função analisada:")
    print(f"  {role_name}")
    print()

    Connection, status = import_pyrfc(args.debug)
    if status != 0 or Connection is None:
        return status

    connection = None
    try:
        connection, ping_status = test_rfc_ping(Connection, args.debug)
        if ping_status != 0 or connection is None:
            return ping_status

        exists, define_detail, define_status = read_agr_define(connection, role_name, args.debug)
        if define_status != 0 or exists is None:
            print_result_unknown(role_name, define_detail)
            return define_status or 3

        if not exists:
            print("[3/3] AGR_TEXTS")
            print("ℹ️ Consulta de descrição ignorada porque a função não existe.")
            print()
            print_result_not_found(role_name)
            return 0

        description, _ = read_agr_texts(connection, role_name, args.debug)
        print_result_exists(role_name, description)
        return 0
    finally:
        safe_close(connection)


if __name__ == "__main__":
    sys.exit(main())
