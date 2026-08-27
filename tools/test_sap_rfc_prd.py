from __future__ import annotations

import argparse
import os
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Diagnóstico read-only da conexão RFC com SAP PRD."
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Mostra traceback técnico completo em caso de erro.",
    )
    return parser.parse_args()


def print_header() -> None:
    print("=" * 60)
    print(" SAP RFC PRD - diagnóstico")
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
    print("[1/3] PyRFC")
    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        print(f"❌ {classify_import_error(exc)}")
        if debug:
            print(format_exception(exc))
            debug_trace(debug)
        print()
        print("RESULTADO:")
        print("❌ O teste não pôde continuar porque a camada RFC não está disponível.")
        return None, 2

    print("✅ PyRFC disponível.")
    print()
    return Connection, 0


def test_rfc_ping(Connection: Any, debug: bool) -> tuple[Any | None, bool, int]:
    print("[2/3] RFC_PING")
    try:
        connection = Connection(**build_connection_params())
        connection.call("RFC_PING")
    except Exception as exc:
        print(f"❌ {classify_rfc_error(exc)}")
        print(format_exception(exc))
        debug_trace(debug)
        print()
        print("RESULTADO:")
        print("❌ Não foi possível estabelecer a conexão RFC com SAP PRD.")
        return None, False, 2

    print("✅ RFC_PING executado com sucesso.")
    print("✅ Conexão SAP PRD estabelecida.")
    print()
    return connection, True, 0


def test_agr_define(connection: Any, debug: bool) -> tuple[bool, int]:
    print("[3/3] AGR_DEFINE")
    try:
        connection.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="AGR_DEFINE",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "AGR_NAME"}],
            ROWCOUNT=1,
        )
    except Exception as exc:
        print("⚠️ RFC conectado, mas leitura AGR_DEFINE falhou.")
        print(f"⚠️ {classify_rfc_error(exc)}")
        print(format_exception(exc))
        debug_trace(debug)
        print()
        print("RESULTADO:")
        print("⚠️ SAP PRD conectado por RFC, mas o utilizador não possui autorização")
        print("para ler AGR_DEFINE.")
        return False, 3

    print("✅ RFC_READ_TABLE autorizado.")
    print("✅ AGR_DEFINE acessível.")
    print()
    print("RESULTADO:")
    print("✅ SAP PRD RFC pronto para consultas do Agente Salsa IT.")
    return True, 0


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
        print("RESULTADO:")
        print("❌ Configuração incompleta. Preencha as variáveis acima e repita o teste.")
        return 2

    print_runtime_config()

    Connection, status = import_pyrfc(args.debug)
    if status != 0 or Connection is None:
        return status

    connection = None
    try:
        connection, ping_ok, ping_status = test_rfc_ping(Connection, args.debug)
        if ping_status != 0 or not ping_ok or connection is None:
            return ping_status

        _, read_status = test_agr_define(connection, args.debug)
        return read_status
    finally:
        safe_close(connection)


if __name__ == "__main__":
    sys.exit(main())
