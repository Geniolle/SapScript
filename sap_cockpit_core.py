###################################################################################
# SAP Cockpit - funcoes puras partilhadas entre o Terminal (SAP Cockpit.py) e a Web
#
# Sem efeitos no import: nao abre SAP, nao le .env, nao imprime nada.
###################################################################################

from __future__ import annotations

import inspect
from typing import Any, Callable


class ArgumentosExecInvalidos(ValueError):
    """Os argumentos calculados nao encaixam na assinatura de executar()."""


def parse_env_line(linha: str):
    """
    Faz parse simples de uma linha KEY=VALUE.
    Ignora linhas vazias e comentários.
    Mantém o valor como texto literal, removendo apenas aspas externas simples/duplas.
    """
    if not linha:
        return None, None

    linha = linha.strip()
    if not linha or linha.startswith("#"):
        return None, None

    if "=" not in linha:
        return None, None

    chave, valor = linha.split("=", 1)
    chave = chave.strip()
    valor = valor.strip()

    if not chave:
        return None, None

    if len(valor) >= 2 and (
        (valor.startswith('"') and valor.endswith('"'))
        or (valor.startswith("'") and valor.endswith("'"))
    ):
        valor = valor[1:-1]

    return chave, valor


def is_scripting_disabled_error(exc: Exception) -> bool:
    parts = [str(exc)]
    try:
        if hasattr(exc, "excepinfo") and exc.excepinfo:
            parts.append(" ".join([str(p) for p in exc.excepinfo if p]))
    except Exception:
        pass

    hay = " ".join(parts).lower()
    has_scripting = any(k in hay for k in ["scripting", "sapgui", "script"])
    has_disabled = any(
        k in hay
        for k in [
            "disabled",
            "not enabled",
            "not active",
            "inactive",
            "desativ",
            "inativ",
            "não está ativo",
            "nao esta ativo",
        ]
    )
    return has_scripting and has_disabled


def validar_argumentos_exec(
    exec_fn: Callable[..., Any],
    ambiente_cockpit: str,
    kwargs: dict[str, Any],
    processo: str = "",
) -> None:
    """
    Garante que executar(ambiente, **kwargs) encaixa na assinatura, antes de a chamar.
    Substitui o padrão `try: exec_fn(..., **kwargs) except TypeError: exec_fn(ambiente)`,
    que executava o script duas vezes quando o TypeError vinha de dentro do próprio script.
    """
    try:
        inspect.signature(exec_fn).bind(ambiente_cockpit, **kwargs)
    except TypeError as exc:
        raise ArgumentosExecInvalidos(
            f"Argumentos incompativeis com executar() em '{processo}': {exc}"
        ) from exc
