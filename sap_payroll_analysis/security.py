"""Camada de segurança READ-ONLY.

Objectivos:

* Nenhuma função RFC fora de `ALLOWED_RFC_FUNCTIONS` pode ser chamada.
* Nenhuma tabela fora de `READ_ONLY_TABLE_WHITELIST` pode ser lida.
* Qualquer nome de função com semântica de escrita é bloqueado por defesa
  em profundidade, mesmo que (por engano) entre na whitelist.

Toda a comunicação com o SAP passa obrigatoriamente por `safe_rfc_call`.
"""

from __future__ import annotations

import logging
from typing import Any

from .config import ALLOWED_RFC_FUNCTIONS, READ_ONLY_TABLE_WHITELIST

logger = logging.getLogger(__name__)

# Palavras que denunciam uma operação de escrita/execução. Usadas como
# segundo filtro, independente da whitelist.
_MUTATION_TOKENS: tuple[str, ...] = (
    "CREATE",
    "CHANGE",
    "UPDATE",
    "INSERT",
    "MODIFY",
    "DELETE",
    "REMOVE",
    "POST",
    "COMMIT",
    "ROLLBACK",
    "SAVE",
    "WRITE",
    "SET_",
    "_SET",
    "ENQUEUE",
    "DEQUEUE",
    "LOCK",
    "SUBMIT",
    "EXECUTE",
    "EXEC_",
    "RUN_",
    "_RUN",
    "START",
    "SCHEDULE",
    "TRANSFER",
    "BDC_",
    "CALL_TRANSACTION",
    "MAINTAIN",
    "UPLOAD",
    "ACTIVATE",
    "GENERATE",
)


class SecurityError(RuntimeError):
    """Levantada quando uma chamada viola a política read-only."""


def assert_function_allowed(function_name: str) -> None:
    name = str(function_name or "").strip().upper()
    if not name:
        raise SecurityError("Nome de função RFC vazio.")
    if name not in {f.upper() for f in ALLOWED_RFC_FUNCTIONS}:
        raise SecurityError(
            f"Função RFC '{function_name}' não está na whitelist "
            f"ALLOWED_RFC_FUNCTIONS. Chamada bloqueada."
        )
    hit = next((tok for tok in _MUTATION_TOKENS if tok in name), None)
    if hit is not None:
        raise SecurityError(
            f"Função RFC '{function_name}' contém o token de escrita '{hit}'. "
            f"Bloqueada pelo guarda read-only."
        )


def assert_table_allowed(table_name: str) -> None:
    name = str(table_name or "").strip().upper()
    if not name:
        raise SecurityError("Nome de tabela vazio.")
    if name not in {t.upper() for t in READ_ONLY_TABLE_WHITELIST}:
        raise SecurityError(
            f"Tabela '{table_name}' não está em READ_ONLY_TABLE_WHITELIST. "
            f"Para a ler, adicione-a explicitamente no código (config.py)."
        )


def safe_rfc_call(connection: Any, function_name: str, **kwargs: Any) -> dict[str, Any]:
    """Único ponto de entrada para invocar funções RFC.

    Valida a whitelist de funções, bloqueia tokens de escrita, valida a
    tabela quando aplicável (`QUERY_TABLE`) e só então delega em
    `connection.call`.
    """
    assert_function_allowed(function_name)

    query_table = kwargs.get("QUERY_TABLE")
    if query_table:
        assert_table_allowed(str(query_table))

    logger.debug(
        "RFC %s %s",
        function_name,
        {k: v for k, v in kwargs.items() if k not in {"OPTIONS"}},
    )
    result = connection.call(function_name, **kwargs)
    return dict(result or {})
