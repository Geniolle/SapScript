"""Descoberta de metadados (DDIC), somente leitura.

Estratégia:

1. Metadados via `DD03L` (+ `DD04T` para textos) — estreito e estável,
   funciona mesmo para tabelas largas (ACDOCA/BSEG/PPDIT).
2. Fallback: `RFC_READ_TABLE` sem `FIELDS` (traz `FIELDS` com nome/tipo/tam.).
3. Sonda de autorização com leitura mínima (1 linha, poucos campos).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Sequence

from .models import FieldGuess, FieldInfo, TableDiag
from .sap_reader import (
    BufferExceeded,
    NotAuthorized,
    RfcReadError,
    TableNotAvailable,
    opt_eq,
    read_table,
)

logger = logging.getLogger(__name__)


# Palavras-chave por conceito, da mais forte para a mais fraca.
CONCEPT_KEYWORDS: dict[str, list[str]] = {
    "empresa": ["BUKRS", "RBUKRS", "COMP_CODE", "BUKRS_", "RCOMP"],
    "conta": ["HKONT", "RACCT", "SAKNR", "SAKONR", "KONTO", "KONT", "ALTKT", "GLACCOUNT"],
    "valor": [
        "WRBTR", "DMBTR", "WSL", "HSL", "TSL", "KSL",
        "BETRG", "BETRA", "WOGBTR", "AMOUNT", "MBETR", "WERT",
    ],
    "moeda": ["WAERS", "RWCUR", "RHCUR", "PSWSL", "HWAER", "WAER", "CURRENCY", "RTCUR"],
    "posting_run": [
        "RUNID", "RUN_ID", "RUNNO", "RUNNR", "LAUFD", "LAUFI", "LINE_ID",
        "EVNUM", "DOC_RUN", "PAYROLL_RUN", "SIMU", "PABRJ", "PABRP",
    ],
    "documento": ["BELNR", "DOC_NUMBER", "DOCNR", "DOCID", "AWKEY", "DOCUMENT", "DOCLN"],
    "item": ["BUZEI", "DOCLN", "LINE", "LINE_ID", "POSNR", "ITEM", "SEQNO", "ZEILE"],
    "exercicio": ["GJAHR", "RYEAR", "BDATJ", "FISCYEAR", "PABRJ"],
    "periodo": ["POPER", "MONAT", "PERIO", "PABRP", "BUPER", "FISCPER"],
    "debito_credito": ["SHKZG", "DRCRK", "DEBCRED", "SOLLHABEN", "SHKZ"],
    "data_lancamento": ["BUDAT", "PSTNG_DATE", "BLDAT", "CPUDT"],
    "wage_type": ["LGART", "WAGETYPE", "LGA", "CUMTY"],
    "pernr": ["PERNR", "PERSNO", "EMPLOYEE"],
    "conta_simbolica": ["SYMKO", "KTOSL", "SYMBOLIC", "KONTH"],
}


def describe_table(connection: Any, table: str, *, sample: int = 0) -> TableDiag:
    """Devolve `TableDiag` com existência, autorização, campos e (opcional) amostra.

    Estratégia (mais fiável para tabelas largas como ACDOCA/BSEG/PPDIT):

    1. Metadados via `DD03L` (+ `DD04T`), que é estreito e estável.
    2. Se DD03L não devolver nada, tentar `RFC_READ_TABLE` sem `FIELDS`.
    3. Sonda de autorização: uma leitura mínima (1 linha, 1 campo) para
       distinguir "sem autorização" de "linha larga".
    """
    diag = TableDiag(table=table.upper())

    # 1) metadados via DDIC
    try:
        diag.fields = _fields_from_ddic(connection, table)
        if diag.fields:
            diag.exists = True
    except NotAuthorized:
        diag.note = "Sem autorização para DD03L."
    except RfcReadError as exc:
        diag.note = f"DD03L: {exc.kind}"

    # 2) fallback: metadados directamente do RFC_READ_TABLE
    if not diag.fields:
        try:
            res = read_table(connection, table, fields=None, options=None, max_rows=1)
            diag.exists = True
            diag.authorized = True
            diag.fields = _fields_from_meta(res.field_meta)
            if sample and diag.fields:
                diag.sample_rows = _sample(connection, table, diag.field_names(), sample)
            return diag
        except TableNotAvailable:
            diag.exists = False
            diag.note = "TABLE_NOT_AVAILABLE"
            return diag
        except NotAuthorized:
            diag.note = diag.note or "Sem autorização (RFC_READ_TABLE)."
            return diag
        except RfcReadError as exc:
            diag.note = diag.note or f"{exc.kind}"
            return diag

    # 3) sonda de autorização (leitura mínima)
    probe = _auth_probe(connection, table, diag.field_names())
    diag.authorized = probe.authorized
    if probe.note and not diag.note:
        diag.note = probe.note
    if probe.exists is False:
        diag.exists = False
    if sample and diag.authorized and diag.fields:
        diag.sample_rows = _sample(connection, table, diag.field_names(), sample)
    if not diag.note:
        diag.note = "OK"
    return diag


@dataclass
class _Probe:
    authorized: bool
    exists: bool | None = None
    note: str = ""


def _auth_probe(connection: Any, table: str, field_names: Sequence[str]) -> "_Probe":
    """Lê 1 linha com poucos campos estreitos para aferir autorização real."""
    narrow = [f for f in field_names if f and not f.startswith(".")][:3] or None
    try:
        read_table(connection, table, fields=narrow, max_rows=1)
        return _Probe(authorized=True)
    except TableNotAvailable:
        return _Probe(authorized=False, exists=False, note="TABLE_NOT_AVAILABLE")
    except NotAuthorized:
        return _Probe(authorized=False, note="Sem autorização para leitura de dados.")
    except BufferExceeded:
        return _Probe(authorized=True, note="Linha larga: ler sempre com FIELDS explícitos.")
    except RfcReadError as exc:
        return _Probe(authorized=False, note=f"Sonda falhou: {exc.kind}")


def _fields_from_meta(meta: Sequence[dict[str, Any]]) -> list[FieldInfo]:
    out: list[FieldInfo] = []
    for m in meta:
        out.append(
            FieldInfo(
                name=str(m.get("FIELDNAME", "")).strip(),
                datatype=str(m.get("TYPE", "")).strip(),
                length=_to_int(m.get("LENGTH")),
                offset=_to_int(m.get("OFFSET")),
                description=str(m.get("FIELDTEXT", "")).strip(),
            )
        )
    return out


def _fields_from_ddic(connection: Any, table: str, *, with_texts: bool = True) -> list[FieldInfo]:
    rows = read_table(
        connection,
        "DD03L",
        fields=["TABNAME", "FIELDNAME", "POSITION", "LENG", "DATATYPE", "ROLLNAME"],
        options=opt_eq("TABNAME", table.upper()),
    ).rows
    rows = [r for r in rows if r.get("FIELDNAME") and not r["FIELDNAME"].startswith(".")]
    rows.sort(key=lambda r: _to_int(r.get("POSITION")))

    # Textos são "nice to have": evita dezenas de chamadas em tabelas largas.
    if with_texts and 0 < len(rows) <= 120:
        texts = _rollname_texts(connection, {r.get("ROLLNAME", "") for r in rows if r.get("ROLLNAME")})
    else:
        texts = {}
    out: list[FieldInfo] = []
    for r in rows:
        out.append(
            FieldInfo(
                name=r["FIELDNAME"].strip(),
                datatype=r.get("DATATYPE", "").strip(),
                length=_to_int(r.get("LENG")),
                description=texts.get(r.get("ROLLNAME", "").strip(), ""),
            )
        )
    return out


def _rollname_texts(connection: Any, rollnames: set[str]) -> dict[str, str]:
    names = sorted(n for n in rollnames if n)
    if not names:
        return {}
    out: dict[str, str] = {}
    # DD04T em blocos para não estourar OPTIONS
    for start in range(0, len(names), 40):
        chunk = names[start : start + 40]
        cond = []
        for i, n in enumerate(chunk):
            cond.append({"TEXT": ("OR " if i else "") + f"ROLLNAME = '{n}'"})
        try:
            rows = read_table(
                connection,
                "DD04T",
                fields=["ROLLNAME", "DDLANGUAGE", "DDTEXT"],
                options=cond,
            ).rows
        except RfcReadError:
            break
        for r in rows:
            lang = r.get("DDLANGUAGE", "")
            key = r.get("ROLLNAME", "").strip()
            if key and (key not in out or lang in {"P", "PT"}):
                out[key] = r.get("DDTEXT", "").strip()
    return out


def _sample(connection: Any, table: str, fields: list[str], n: int) -> list[dict[str, str]]:
    # limita a largura da amostra: no máximo ~20 campos "interessantes"
    picked = fields[:20]
    try:
        return read_table(connection, table, fields=picked, max_rows=n).rows
    except RfcReadError:
        return []


def guess_fields(diag: TableDiag, concepts: Sequence[str] | None = None) -> dict[str, FieldGuess]:
    """Para cada conceito, ordena os campos da tabela por afinidade de nome."""
    available = diag.field_names()
    upper_map = {f.upper(): f for f in available}
    wanted = list(concepts) if concepts else list(CONCEPT_KEYWORDS)
    out: dict[str, FieldGuess] = {}

    for concept in wanted:
        keywords = CONCEPT_KEYWORDS.get(concept, [concept.upper()])
        scored: list[tuple[int, str]] = []
        for fld_upper, fld in upper_map.items():
            best = _score(fld_upper, keywords)
            if best > 0:
                scored.append((best, fld))
        scored.sort(key=lambda t: (-t[0], t[1]))
        candidates = [f for _, f in scored]
        out[concept] = FieldGuess(
            concept=concept,
            candidates=candidates,
            chosen=candidates[0] if candidates else None,
        )
    return out


def _score(field_upper: str, keywords: Sequence[str]) -> int:
    for rank, kw in enumerate(keywords):
        if field_upper == kw:
            return 1000 - rank
        if field_upper.startswith(kw) or field_upper.endswith(kw):
            return 500 - rank
        if kw in field_upper:
            return 200 - rank
    return 0


def _to_int(value: Any) -> int:
    try:
        return int(str(value).strip() or "0")
    except (TypeError, ValueError):
        return 0
