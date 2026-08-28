"""Ligação RFC ao SAP (context manager), READ-ONLY.

Credenciais lidas exclusivamente do `.env` do projeto. Nunca hardcoded,
nunca escritas em log.
"""

from __future__ import annotations

import logging
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from .config import ENV_PREFIXES, REQUIRED_ENV_SUFFIXES
from .security import safe_rfc_call

logger = logging.getLogger(__name__)


class SapConnectionError(RuntimeError):
    """Falha de configuração ou de ligação RFC."""


def find_project_root(start: Path | None = None) -> Path:
    """Sobe na árvore de directórios até encontrar `.env` ou `.env.example`."""
    explicit = str(os.getenv("SAP_SCRIPT_PROJECT_DIR", "") or "").strip()
    if explicit and Path(explicit).exists():
        return Path(explicit).resolve()

    current = (start or Path(__file__)).resolve().parent
    for candidate in [current, *current.parents]:
        if (candidate / ".env").exists() or (candidate / ".env.example").exists():
            return candidate
    raise SapConnectionError("Não foi possível localizar a raiz do projeto (.env).")


def load_env(project_root: Path | None = None) -> Path:
    """Carrega o `.env` do projeto sem sobrepor variáveis já definidas."""
    from dotenv import load_dotenv

    root = project_root or find_project_root()
    env_path = root / ".env"
    load_dotenv(env_path, override=False)
    logger.info("Ficheiro .env carregado: %s", env_path)
    return root


def _prefix_is_complete(prefix: str) -> bool:
    return all(os.getenv(f"{prefix}{sfx}", "").strip() for sfx in REQUIRED_ENV_SUFFIXES)


def resolve_env_prefix() -> str:
    """Devolve o primeiro prefixo de `.env` totalmente preenchido.

    Ordem: `SAP_R3_` (oficial), depois `SAP_DEV_` (fallback, mesmo host).
    """
    for prefix in ENV_PREFIXES:
        if _prefix_is_complete(prefix):
            if prefix != ENV_PREFIXES[0]:
                logger.warning(
                    "Variáveis %s* não encontradas; a usar fallback %s* "
                    "(host %s). Recomenda-se renomear no .env para %s*.",
                    ENV_PREFIXES[0],
                    prefix,
                    os.getenv(f"{prefix}ASHOST", "?"),
                    ENV_PREFIXES[0],
                )
            return prefix
    esperado = ", ".join(f"{ENV_PREFIXES[0]}{s}" for s in REQUIRED_ENV_SUFFIXES)
    raise SapConnectionError(
        f"Configuração RFC incompleta. Preencha no .env: {esperado}."
    )


def require_prefix(prefix: str, *, purpose: str) -> str:
    """Exige que `{prefix}*` esteja COMPLETO no ambiente, sem qualquer fallback.

    Usado por fases que têm de correr num sistema específico (ex.: a
    reconciliação Payroll×REGU só faz sentido no R/3 real). Se faltar algum
    parâmetro, aborta com mensagem clara — nunca cai silenciosamente noutro
    prefixo (ex.: `SAP_DEV_*`).
    """
    if _prefix_is_complete(prefix):
        return prefix
    missing = [f"{prefix}{s}" for s in REQUIRED_ENV_SUFFIXES
               if not os.getenv(f"{prefix}{s}", "").strip()]
    raise SapConnectionError(
        f"{prefix}* connection parameters required for {purpose}. "
        f"Ausentes no .env: {', '.join(missing)}. "
        f"Sem fallback para outro sistema nesta operação."
    )


def build_connection_params(prefix: str | None = None) -> dict[str, str]:
    """Monta os parâmetros para `pyrfc.Connection`. A password nunca é logada."""
    pref = prefix or resolve_env_prefix()
    missing = [f"{pref}{s}" for s in REQUIRED_ENV_SUFFIXES if not os.getenv(f"{pref}{s}", "").strip()]
    if missing:
        raise SapConnectionError(f"Variáveis obrigatórias ausentes: {', '.join(missing)}")

    params = {
        "user": os.environ[f"{pref}USER"],
        "passwd": os.environ[f"{pref}PASSWD"],
        "ashost": os.environ[f"{pref}ASHOST"],
        "sysnr": os.environ[f"{pref}SYSNR"],
        "client": os.environ[f"{pref}CLIENT"],
        "lang": os.getenv(f"{pref}LANG", "PT").strip() or "PT",
    }
    return params


def safe_connection_summary(params: dict[str, str]) -> dict[str, str]:
    """Versão dos parâmetros segura para logs / relatórios (sem password)."""
    return {
        "user": params.get("user", ""),
        "ashost": params.get("ashost", ""),
        "sysnr": params.get("sysnr", ""),
        "client": params.get("client", ""),
        "lang": params.get("lang", ""),
        "passwd": "********",
    }


@contextmanager
def sap_connection(prefix: str | None = None) -> Iterator[Any]:
    """Context manager que abre, testa (RFC_PING) e fecha a ligação RFC.

    Uso::

        with sap_connection() as conn:
            ...
    """
    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:  # pragma: no cover - depende do SDK local
        raise SapConnectionError(
            "PyRFC indisponível. Instale o SAP NetWeaver RFC SDK e o pacote pyrfc."
        ) from exc

    params = build_connection_params(prefix)
    logger.info("A ligar a SAP RFC: %s", safe_connection_summary(params))

    connection = None
    try:
        connection = Connection(**params)
        safe_rfc_call(connection, "RFC_PING")
        logger.info("SAP RFC conectado (RFC_PING OK).")
        yield connection
    finally:
        try:
            if connection is not None:
                connection.close()
                logger.info("Ligação RFC fechada.")
        except Exception:  # pragma: no cover
            logger.debug("Falha (ignorada) ao fechar a ligação RFC.", exc_info=True)
