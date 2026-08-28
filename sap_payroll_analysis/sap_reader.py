"""Leitor genérico RFC_READ_TABLE, com paginação e parsing robusto.

Nunca assume que uma tabela cabe numa única chamada: pagina com
ROWSKIPS/ROWCOUNT até esgotar os registos.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Sequence

from .security import assert_table_allowed, safe_rfc_call

logger = logging.getLogger(__name__)

DELIMITER = "|"
DEFAULT_PAGE_SIZE = 5000
MAX_PAGES = 2000  # trava de segurança contra loops infinitos


class RfcReadError(RuntimeError):
    """Erro ao ler uma tabela via RFC_READ_TABLE."""

    def __init__(self, message: str, *, kind: str = "RFC_ERROR", table: str = "") -> None:
        super().__init__(message)
        self.kind = kind
        self.table = table


class TableNotAvailable(RfcReadError):
    pass


class FieldNotValid(RfcReadError):
    pass


class NotAuthorized(RfcReadError):
    pass


class BufferExceeded(RfcReadError):
    pass


class NoData(RfcReadError):
    """RFC_READ_TABLE devolveu TABLE_WITHOUT_DATA (nenhuma linha corresponde)."""


_ERROR_MARKERS: tuple[tuple[str, type[RfcReadError], str], ...] = (
    ("TABLE_WITHOUT_DATA", NoData, "TABLE_WITHOUT_DATA"),
    ("TABLE_NOT_AVAILABLE", TableNotAvailable, "TABLE_NOT_AVAILABLE"),
    ("NOT_AVAILABLE", TableNotAvailable, "TABLE_NOT_AVAILABLE"),
    ("FIELD_NOT_VALID", FieldNotValid, "FIELD_NOT_VALID"),
    ("DATA_BUFFER_EXCEEDED", BufferExceeded, "DATA_BUFFER_EXCEEDED"),
    ("SAPSQL_DATA_LOSS", RfcReadError, "DATA_LOSS"),
    ("DATA WAS LOST WHILE COPYING", RfcReadError, "DATA_LOSS"),
    ("NOT_AUTHORIZED", NotAuthorized, "NOT_AUTHORIZED"),
    ("AUTHORIZATION", NotAuthorized, "NOT_AUTHORIZED"),
    ("AUTHORISATION", NotAuthorized, "NOT_AUTHORIZED"),
    ("S_TABU", NotAuthorized, "NOT_AUTHORIZED"),
)


def _classify(exc: BaseException, table: str) -> RfcReadError:
    text = f"{exc.__class__.__name__} {getattr(exc, 'key', '')} {getattr(exc, 'message', '')} {exc}".upper()
    for marker, cls, kind in _ERROR_MARKERS:
        if marker in text:
            return cls(f"{table}: {kind} ({exc})", kind=kind, table=table)
    if "ABAP_APPLICATION_FAILURE" in text or "ABAP_RUNTIME_FAILURE" in text:
        return RfcReadError(f"{table}: {exc}", kind="ABAP_FAILURE", table=table)
    return RfcReadError(f"{table}: {exc}", kind="RFC_ERROR", table=table)


# ---------------------------------------------------------------------------
# Construção de OPTIONS (WHERE) respeitando o limite de 72 chars por linha.
# ---------------------------------------------------------------------------

def opt_eq(field: str, value: str) -> list[dict[str, str]]:
    return [{"TEXT": f"{field} = '{_escape(value)}'"}]


def opt_in(field: str, values: Iterable[str]) -> list[dict[str, str]]:
    """`field IN (...)` como igualdades ligadas por OR, uma por linha."""
    rows: list[dict[str, str]] = []
    for index, value in enumerate(values):
        prefix = "OR " if index else ""
        rows.append({"TEXT": f"{prefix}{field} = '{_escape(value)}'"})
    return rows or [{"TEXT": "1 = 2"}]  # conjunto vazio -> nunca devolve linhas


def opt_and(*groups: Sequence[dict[str, str]]) -> list[dict[str, str]]:
    """Liga grupos de condições com AND. Cada grupo é envolvido em parênteses."""
    combined: list[dict[str, str]] = []
    real_groups = [list(g) for g in groups if g]
    for gi, group in enumerate(real_groups):
        if gi:
            combined.append({"TEXT": "AND"})
        if len(group) > 1:
            combined.append({"TEXT": "("})
        combined.extend(dict(row) for row in group)
        if len(group) > 1:
            combined.append({"TEXT": ")"})
    return combined


def _escape(value: str) -> str:
    return str(value).replace("'", "''")


# ---------------------------------------------------------------------------
# Parsing de valores SAP.
# ---------------------------------------------------------------------------

def sap_str_to_decimal(raw: str) -> Decimal:
    """Converte a representação textual de um número SAP em `Decimal`.

    Trata: espaços, separador de milhares, sinal à direita (`123,45-`),
    sinal à esquerda, parênteses de negativo e vazio (-> 0).
    """
    text = str(raw or "").strip()
    if not text or text in {"-", "*"}:
        return Decimal("0")

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    if text.endswith("-"):
        negative = True
        text = text[:-1].strip()
    if text.startswith("-"):
        negative = True
        text = text[1:].strip()
    if text.startswith("+"):
        text = text[1:].strip()

    # Normaliza separadores: mantém o último como decimal.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # vírgula é decimal se aparece só uma vez e com <=2 casas à direita
        if text.count(",") == 1 and len(text.split(",")[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    text = text.replace(" ", "")
    try:
        value = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Valor SAP não numérico: {raw!r}") from exc
    return -value if negative else value


def normalize_sign(amount: Decimal, debit_credit: str) -> Decimal:
    """Normaliza para a convenção: Débito = positivo, Crédito = negativo.

    `debit_credit` aceita o indicador SHKZG / DRCRK ('S'/'H') ou 'D'/'C'.
    Se vazio, devolve o valor tal como está (assume já com sinal).
    """
    flag = str(debit_credit or "").strip().upper()
    magnitude = abs(amount)
    if flag in {"S", "D"}:
        return magnitude
    if flag in {"H", "C"}:
        return -magnitude
    return amount


# ---------------------------------------------------------------------------
# Leitura.
# ---------------------------------------------------------------------------

@dataclass
class ReadResult:
    table: str
    fields: list[str]
    rows: list[dict[str, str]]
    field_meta: list[dict[str, Any]]
    pages: int
    truncated: bool = False

    def __len__(self) -> int:  # pragma: no cover - conveniência
        return len(self.rows)


def read_table(
    connection: Any,
    table: str,
    *,
    fields: Sequence[str] | None = None,
    options: Sequence[dict[str, str]] | None = None,
    page_size: int = DEFAULT_PAGE_SIZE,
    max_rows: int | None = None,
) -> ReadResult:
    """Lê `table` via RFC_READ_TABLE com paginação automática.

    Levanta subclasses de `RfcReadError` para os erros conhecidos.
    """
    assert_table_allowed(table)
    field_list = list(fields or [])
    fields_payload = [{"FIELDNAME": f} for f in field_list]
    options_payload = [dict(o) for o in (options or [])]

    all_rows: list[dict[str, str]] = []
    field_meta: list[dict[str, Any]] = []
    sap_fields: list[str] = list(field_list)
    skip = 0
    pages = 0
    truncated = False

    while True:
        if pages >= MAX_PAGES:
            logger.warning("%s: atingido MAX_PAGES=%s, a interromper paginação.", table, MAX_PAGES)
            truncated = True
            break

        want = page_size
        if max_rows is not None:
            want = min(page_size, max_rows - len(all_rows))
            if want <= 0:
                truncated = True
                break

        try:
            result = safe_rfc_call(
                connection,
                "RFC_READ_TABLE",
                QUERY_TABLE=table,
                DELIMITER=DELIMITER,
                FIELDS=fields_payload,
                OPTIONS=options_payload,
                ROWSKIPS=skip,
                ROWCOUNT=want,
            )
        except Exception as exc:  # noqa: BLE001 - reclassificado abaixo
            err = _classify(exc, table)
            if isinstance(err, NoData):
                # "sem linhas" não é erro: devolve o que já houver (normalmente vazio).
                break
            raise err from exc

        pages += 1
        meta = result.get("FIELDS", []) or []
        if meta and not field_meta:
            field_meta = [dict(m) for m in meta]
            if not sap_fields:
                sap_fields = [str(m.get("FIELDNAME", "")) for m in meta]

        data = result.get("DATA", []) or []
        for item in data:
            wa = str(item.get("WA", "") or "")
            parts = wa.split(DELIMITER)
            row = {
                name: (parts[i].strip() if i < len(parts) else "")
                for i, name in enumerate(sap_fields)
            }
            all_rows.append(row)

        logger.debug("%s: página %s, +%s linhas (skip=%s)", table, pages, len(data), skip)

        if len(data) < want:
            break
        skip += want

    return ReadResult(
        table=table,
        fields=sap_fields,
        rows=all_rows,
        field_meta=field_meta,
        pages=pages,
        truncated=truncated,
    )


def count_rows(connection: Any, table: str, options: Sequence[dict[str, str]] | None = None) -> int:
    """Conta linhas lendo apenas a 1ª coluna (barato). -1 se indeterminado."""
    try:
        res = read_table(connection, table, fields=None, options=options, page_size=DEFAULT_PAGE_SIZE)
        return len(res.rows)
    except RfcReadError:
        return -1
