"""Fase 4.2 — localizar a origem técnica do delta entre

    SUM( PPOIX.BETRG  onde TSLIN in {LINUMs que alimentam a linha alvo} )

e

    PPDIT.WRBTR  da linha (DOCNUM, DOCLIN) alvo

para a conta colectora do net-pay (23120000), empresa 1010, run 0000001298.

100 % READ-ONLY (RFC_READ_TABLE via `safe_rfc_call`). NÃO lê o cluster PCL2,
NÃO chama `PYXX_READ_PAYROLL_RESULT` / `CU_READ_RGDIR` / `HR_GET_PAYROLL_RESULTS`,
NÃO executa Payroll nem posting. Só observação.

O objectivo não é encontrar "valores que somem o delta": é reconstruir a cadeia
com checkpoints persistidos e classificar EM QUE ESTÁGIO o valor muda, marcando
cada conclusão com [PROVED] / [OBSERVED] / [CANDIDATE] / [HYPOTHESIS] /
[UNEXPLAINED].
"""

from __future__ import annotations

import csv
import json
import logging
from dataclasses import dataclass, field
from decimal import Decimal
from itertools import combinations
from pathlib import Path
from typing import Any, Iterable

from .config import AnalysisParams
from .sap_reader import (
    NoData,
    RfcReadError,
    opt_and,
    opt_eq,
    opt_in,
    read_table,
    sap_str_to_decimal,
)
from .wagetype_trace import _PPOIX_FIELDS, _PPOIX_FIELDS_MIN, _read_ppoix_options, _signed

logger = logging.getLogger(__name__)

_ZERO = Decimal("0")
_CENT = Decimal("0.01")

#: Tabelas de posting HR que têm um campo monetário (confirmado por DDIC neste
#: sistema). Fora destas três não há valor persistido na cadeia de posting.
MONEY_TABLES_IN_CHAIN = ("PPOIX", "PPDIT", "PPDST")

#: Tabelas da cadeia de posting HR ECC inspeccionadas (lista explícita — NÃO
#: varredura ampla de prefixos, que apanharia tabelas S/4 SCM 'PPO*/PPD*' sem
#: relação). Só se verifica se cada uma tem um campo de MONTANTE (CURR).
_POSTING_CHAIN_TABLES = (
    "PPOIX", "PPOPX", "PPDIX", "PPDIT", "PPDST", "PPDSH", "PPDHD", "PPDMSG",
    "PEVST", "PEVAT", "PEVSH", "T52OKT", "T52OKK", "T52OKP",
)

#: Só CURR conta como "montante" (DEC/QUAN apanham quantidades/rácios).
_MONEY_DDIC_TYPES = {"CURR"}
_MONEY_NAME_HINTS = ("BETRG", "WRBTR", "DMBTR", "MBETR")

#: Classificações finais possíveis para a origem do delta.
DELTA_ORIGIN_CLASSES = (
    "PROVED_AT_PPOPX",
    "PROVED_AT_INTERMEDIATE_STAGE",
    "PROVED_BETWEEN_PPOIX_AND_PPDIT",
    "PROVED_PREVIOUS_RUN_NETTING",
    "PROVED_OTHER_RULE",
    "PARTIALLY_EXPLAINED",
    "UNEXPLAINED",
)


def _q(v: Any) -> Decimal:
    try:
        return Decimal(str(v))
    except Exception:  # noqa: BLE001
        return _ZERO


def _sum_betrg(rows: Iterable[dict[str, str]]) -> Decimal:
    return sum((_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")) for r in rows), _ZERO)


def _is_zero_tslin(tslin: str) -> bool:
    return str(tslin or "").strip("0") == ""


def _table_field_names(connection: Any, params: AnalysisParams, table: str) -> list[str]:
    """Nomes de campo de uma tabela via DD03L (com filtro TABNAME reaplicado no
    cliente — defensivo contra RFC_READ_TABLE que devolva a mais)."""
    try:
        rows = read_table(connection, "DD03L",
                          fields=["TABNAME", "FIELDNAME", "POSITION"],
                          options=opt_and(opt_eq("TABNAME", table.upper())),
                          page_size=200_000).rows
    except (RfcReadError, NoData):
        return []
    rows = [r for r in rows
            if r.get("TABNAME") == table.upper()
            and r.get("FIELDNAME") and not r["FIELDNAME"].startswith(".")]
    rows.sort(key=lambda r: int(r.get("POSITION") or 0) if str(r.get("POSITION") or "0").isdigit() else 0)
    return [r["FIELDNAME"].strip() for r in rows]


# ---------------------------------------------------------------------------
# Modelos
# ---------------------------------------------------------------------------

@dataclass
class PostingCheckpoint:
    """Um ponto observável (ou provadamente não-observável) da cadeia."""

    stage: str = ""            # ordem lógica: "1-PPOIX", "2-PPOPX", "3-PPDST", "4-PPDIT"
    source: str = ""           # tabela / origem
    key: str = ""              # chave textual do checkpoint
    row_count: int = 0
    amount: Decimal | None = None   # None => estágio sem valor persistido
    currency: str = "EUR"
    derivation: str = ""       # como foi obtido
    evidence: str = ""         # [PROVED]/[OBSERVED]/...

    def as_dict(self) -> dict[str, Any]:
        return {
            "stage": self.stage,
            "source": self.source,
            "key": self.key,
            "row_count": self.row_count,
            "amount": None if self.amount is None else str(self.amount),
            "currency": self.currency,
            "derivation": self.derivation,
            "evidence": self.evidence,
        }


@dataclass
class PostingDeltaTrace:
    run_id: str = ""
    company: str = ""
    account: str = ""
    docnum: str = ""
    doclin: str = ""

    feeder_linums: list[str] = field(default_factory=list)
    ppoix_rows: int = 0
    ppoix_sum: Decimal = _ZERO
    ppdit_wrbtr: Decimal = _ZERO
    delta: Decimal = _ZERO

    checkpoints: list[PostingCheckpoint] = field(default_factory=list)
    intermediates: dict[str, Any] = field(default_factory=dict)
    ppopx: dict[str, Any] = field(default_factory=dict)
    reconciliation: dict[str, Any] = field(default_factory=dict)
    zero_tslin: dict[str, Any] = field(default_factory=dict)
    tslin_pairs: list[dict[str, Any]] = field(default_factory=list)
    seqno_usage: dict[str, Any] = field(default_factory=dict)
    previous_runs: dict[str, Any] = field(default_factory=dict)
    previous_postings: dict[str, Any] = field(default_factory=dict)
    lgart_breakdown: dict[str, Any] = field(default_factory=dict)
    pernr_breakdown: dict[str, Any] = field(default_factory=dict)
    momag: dict[str, Any] = field(default_factory=dict)
    delta_candidates: dict[str, Any] = field(default_factory=dict)
    first_divergence: dict[str, Any] = field(default_factory=dict)
    classification: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def warn(self, m: str) -> None:
        if m not in self.warnings:
            self.warnings.append(m)
            logger.warning(m)

    def as_dict(self) -> dict[str, Any]:
        return {
            "input": {
                "run": self.run_id, "company": self.company, "account": self.account,
                "docnum": self.docnum, "doclin": self.doclin,
            },
            "feeder_linums": self.feeder_linums,
            "ppoix_rows": self.ppoix_rows,
            "ppoix_sum": str(self.ppoix_sum),
            "ppdit_wrbtr": str(self.ppdit_wrbtr),
            "delta": str(self.delta),
            "checkpoints": [c.as_dict() for c in self.checkpoints],
            "intermediates": self.intermediates,
            "ppopx": self.ppopx,
            "reconciliation": self.reconciliation,
            "zero_tslin": self.zero_tslin,
            "tslin_pairs": self.tslin_pairs,
            "seqno_usage": self.seqno_usage,
            "previous_runs": self.previous_runs,
            "previous_postings": self.previous_postings,
            "lgart_breakdown": self.lgart_breakdown,
            "pernr_breakdown": self.pernr_breakdown,
            "momag": self.momag,
            "delta_candidates": self.delta_candidates,
            "first_divergence": self.first_divergence,
            "classification": self.classification,
            "warnings": self.warnings,
        }


# ---------------------------------------------------------------------------
# Leitores auxiliares
# ---------------------------------------------------------------------------

def _ppdix_rows(connection: Any, params: AnalysisParams, run: str) -> list[dict[str, str]]:
    return read_table(connection, "PPDIX", fields=["RUNID", "LINUM", "DOCNUM", "DOCLIN"],
                      options=opt_and(opt_eq("RUNID", run)), page_size=params.page_size).rows


def _feeders_for_line(connection: Any, params: AnalysisParams, run: str,
                      docnum: str, doclin: str) -> tuple[list[str], list[dict[str, str]]]:
    """LINUMs (PPDIX) que alimentam (DOCNUM, DOCLIN), incluindo *qualquer* RUNID.

    O filtro por RUNID é aplicado a seguir; devolve também as linhas cruas para
    inspeccionar EVTYP / runs cruzados.
    """
    rows = read_table(connection, "PPDIX", fields=["RUNID", "EVTYP", "LINUM", "DOCNUM", "DOCLIN"],
                      options=opt_and(opt_eq("DOCNUM", docnum)), page_size=params.page_size).rows
    hits = [r for r in rows if r.get("DOCNUM") == docnum and r["DOCLIN"] == doclin]
    linums = sorted({r["LINUM"] for r in hits if r["RUNID"] == run})
    return linums, hits


def _ppdit_line(connection: Any, params: AnalysisParams, docnum: str,
                doclin: str) -> dict[str, str] | None:
    try:
        rows = read_table(connection, "PPDIT",
                          fields=["DOCNUM", "DOCLIN", "BUKRS", "HKONT", "KTOSL", "WRBTR",
                                  "WAERS", "NEG_POSTNG", "ITTYP"],
                          options=opt_and(opt_eq("DOCNUM", docnum)), page_size=params.page_size).rows
    except NoData:
        return None
    for r in rows:
        if r.get("DOCNUM") == docnum and r["DOCLIN"] == doclin:
            return r
    return None


def _read_run_ppoix(connection: Any, params: AnalysisParams, run: str) -> list[dict[str, str]]:
    """Todos os PPOIX de um run (com recurso a campos reduzidos se DATA_LOSS).

    O filtro RUNID é reaplicado no cliente (defensivo: o fallback de campos
    reduzidos e alguns quirks de RFC_READ_TABLE podem devolver a mais).
    """
    rows = _read_ppoix_options(connection, params, opt_and(opt_eq("RUNID", run)),
                               max_rows=1_000_000)
    return [r for r in rows if not r.get("RUNID") or r.get("RUNID") == run]


# ---------------------------------------------------------------------------
# 5 / 6 — PPOPX e estruturas intermédias
# ---------------------------------------------------------------------------

def discover_posting_intermediates(connection: Any, params: AnalysisParams) -> dict[str, Any]:
    """DDIC: para cada tabela PPO*/PPD*/PEV*/HRPP* transparente, diz se tem um
    campo monetário. Documenta que NÃO existe checkpoint com valor entre PPOIX e
    PPDIT (excepto PPDST, tratado à parte)."""
    out: dict[str, Any] = {"inspected_tables": list(_POSTING_CHAIN_TABLES), "tables": {}}
    names = list(_POSTING_CHAIN_TABLES)
    nameset = set(names)
    d3: list[dict[str, str]] = []
    for start in range(0, len(names), 20):
        chunk = names[start:start + 20]
        try:
            d3 += read_table(connection, "DD03L",
                             fields=["TABNAME", "FIELDNAME", "DATATYPE"],
                             options=opt_in("TABNAME", chunk), page_size=200_000).rows
        except (RfcReadError, NoData) as exc:
            logger.warning("DD03L chunk indisponível: %s", exc)
    by_tab: dict[str, list[dict[str, str]]] = {}
    for r in d3:
        if r.get("TABNAME") in nameset and r.get("FIELDNAME") and not r["FIELDNAME"].startswith("."):
            by_tab.setdefault(r["TABNAME"], []).append(r)
    money_tables: list[str] = []
    for name in names:
        flds = by_tab.get(name, [])
        if not flds:
            continue
        money = sorted({f["FIELDNAME"] for f in flds
                        if f.get("DATATYPE") in _MONEY_DDIC_TYPES
                        and any(h in f["FIELDNAME"] for h in _MONEY_NAME_HINTS)})
        out["tables"][name] = {"n_fields": len(flds), "money_fields": money}
        if money:
            money_tables.append(name)
    out["money_bearing_tables"] = money_tables
    out["conclusion"] = (
        "[PROVED] Na cadeia de posting HR só têm campo monetário: "
        + ", ".join(t for t in MONEY_TABLES_IN_CHAIN if t in money_tables or t == "PPOIX")
        + ". PPDIX (LINUM->DOCNUM/DOCLIN) e PPOPX (índice POSTNUM/TSLIN) NÃO têm "
          "valor — não podem servir de checkpoint de montante."
    )
    return out


def inspect_ppopx(connection: Any, params: AnalysisParams, run: str,
                  target_rows: list[dict[str, str]] | None = None) -> dict[str, Any]:
    """PPOPX: populada? quantas linhas no run? tem campo monetário? há overlap de
    chave com as linhas PPOIX da linha alvo?"""
    out: dict[str, Any] = {}
    flds = _table_field_names(connection, params, "PPOPX")
    if not flds:
        return {"error": "DDIC PPOPX indisponível"}
    out["fields"] = flds
    out["has_money_field"] = any(
        h in n for n in flds for h in _MONEY_NAME_HINTS
    )
    try:
        rows = read_table(connection, "PPOPX",
                          fields=[f for f in ("PERNR", "SEQNO", "RUNID", "POSTNUM", "TSLIN", "ACTSIGN")
                                  if f in flds],
                          options=opt_and(opt_eq("RUNID", run)), page_size=200_000).rows
    except (RfcReadError, NoData) as exc:
        return {**out, "error": f"leitura PPOPX: {exc}", "rows_for_run": 0}
    rows = [r for r in rows if not r.get("RUNID") or r.get("RUNID") == run]
    out["rows_for_run"] = len(rows)
    out["actsign_dist"] = _count(rows, "ACTSIGN")
    out["tslin_zero_rows"] = sum(1 for r in rows if _is_zero_tslin(r.get("TSLIN", "")))

    if target_rows:
        keysets = {
            "PERNR+SEQNO+POSTNUM+TSLIN": lambda r: (r.get("PERNR"), r.get("SEQNO"), r.get("POSTNUM"), r.get("TSLIN")),
            "PERNR+SEQNO+POSTNUM": lambda r: (r.get("PERNR"), r.get("SEQNO"), r.get("POSTNUM")),
            "PERNR+SEQNO+TSLIN": lambda r: (r.get("PERNR"), r.get("SEQNO"), r.get("TSLIN")),
            "PERNR+SEQNO": lambda r: (r.get("PERNR"), r.get("SEQNO")),
        }
        overlap: dict[str, Any] = {}
        for label, kf in keysets.items():
            pset = {kf(r) for r in rows}
            matched = [r for r in target_rows if kf(r) in pset]
            overlap[label] = {
                "matched_rows": len(matched),
                "matched_sum": str(_sum_betrg(matched)),
            }
        out["overlap_with_target_line"] = overlap
        out["conclusion"] = (
            "[PROVED] Nenhuma linha PPOIX da linha alvo tem correspondência em "
            "PPOPX (0 em todas as chaves testadas) — PPOPX não é um índice de "
            "'já lançado anteriormente' para esta linha."
            if all(v["matched_rows"] == 0 for v in overlap.values())
            else "[OBSERVED] Há overlap parcial PPOPX<->PPOIX — ver `overlap_with_target_line`."
        )
    return out


# ---------------------------------------------------------------------------
# 8 — checkpoints
# ---------------------------------------------------------------------------

def build_posting_checkpoints(trace: PostingDeltaTrace, intermediates: dict[str, Any],
                              ppopx: dict[str, Any], ppdst_amount: Decimal | None,
                              ppdst_rows: int) -> list[PostingCheckpoint]:
    cps: list[PostingCheckpoint] = []
    cps.append(PostingCheckpoint(
        stage="1-PPOIX", source="PPOIX",
        key=f"RUNID={trace.run_id} TSLIN in {trace.feeder_linums}",
        row_count=trace.ppoix_rows, amount=trace.ppoix_sum,
        derivation="SUM(_signed(BETRG, NEG_POSTNG)) das linhas cujo TSLIN alimenta (via PPDIX) a linha alvo",
        evidence="[PROVED]",
    ))
    cps.append(PostingCheckpoint(
        stage="2-PPOPX", source="PPOPX",
        key=f"RUNID={trace.run_id}",
        row_count=int(ppopx.get("rows_for_run", 0) or 0), amount=None,
        derivation="índice POSTNUM/TSLIN, ACTSIGN='P'; sem campo monetário",
        evidence="[PROVED] sem valor persistido",
    ))
    cps.append(PostingCheckpoint(
        stage="3-PPDST", source="PPDST",
        key=f"DOCNUM={trace.docnum}",
        row_count=ppdst_rows,
        amount=ppdst_amount,
        derivation="split de custeio da linha PPDIT (WRBTR); vazio => sem decomposição persistida",
        evidence="[PROVED] vazio para este run" if ppdst_rows == 0 else "[OBSERVED]",
    ))
    cps.append(PostingCheckpoint(
        stage="4-PPDIT", source="PPDIT",
        key=f"{trace.docnum}/{trace.doclin}",
        row_count=1, amount=trace.ppdit_wrbtr,
        derivation="_signed(WRBTR, NEG_POSTNG) da linha do documento de posting HR",
        evidence="[PROVED]",
    ))
    return cps


# ---------------------------------------------------------------------------
# 4 / 9 — reconciliação da linha alvo, LINUM a LINUM
# ---------------------------------------------------------------------------

def reconcile_target_line(connection: Any, params: AnalysisParams, run: str,
                          docnum: str, doclin: str,
                          run_ppoix: list[dict[str, str]] | None = None) -> dict[str, Any]:
    linums, ppdix_hits = _feeders_for_line(connection, params, run, docnum, doclin)
    px = run_ppoix if run_ppoix is not None else _read_run_ppoix(connection, params, run)
    line_rows = [r for r in px if r.get("TSLIN") in set(linums)]
    pit = _ppdit_line(connection, params, docnum, doclin)
    wrbtr = _signed(pit.get("WRBTR", "0"), pit.get("NEG_POSTNG", "")) if pit else _ZERO

    per_linum = []
    for ln in linums:
        rr = [r for r in line_rows if r.get("TSLIN") == ln]
        per_linum.append({
            "linum": ln,
            "rows": len(rr),
            "sum": str(_sum_betrg(rr)),
            "by_lgart": _sum_by(rr, "LGART"),
            "by_momag": _sum_by(rr, "MOMAG"),
            "by_postnum": _sum_by(rr, "POSTNUM"),
        })
    total = _sum_betrg(line_rows)
    return {
        "docnum": docnum, "doclin": doclin,
        "ppdit_account": (pit or {}).get("HKONT", ""),
        "ppdit_ktosl": (pit or {}).get("KTOSL", ""),
        "feeder_linums": linums,
        "ppdix_feeder_rows_all_runs": [
            {"runid": r["RUNID"], "evtyp": r.get("EVTYP", ""), "linum": r["LINUM"]}
            for r in ppdix_hits
        ],
        "ppoix_rows": len(line_rows),
        "ppoix_sum": str(total),
        "ppdit_wrbtr": str(wrbtr),
        "delta": str(total - wrbtr),
        "per_linum": per_linum,
        "no_amount_bearing_intermediate": True,
        "note": (
            "[PROVED] A linha PPDIT é alimentada apenas pelos LINUM listados "
            "(PPDIX, todos os RUNID/EVTYP verificados). PPDIX não transporta "
            "montante; o valor só reaparece, já transformado, em PPDIT.WRBTR."
        ),
    }


# ---------------------------------------------------------------------------
# 10 / 11 — TSLIN = 0
# ---------------------------------------------------------------------------

def analyze_zero_tslin(connection: Any, params: AnalysisParams, run: str,
                       run_ppoix: list[dict[str, str]] | None = None) -> dict[str, Any]:
    px = run_ppoix if run_ppoix is not None else _read_run_ppoix(connection, params, run)
    zero = [r for r in px if _is_zero_tslin(r.get("TSLIN", ""))]
    non_zero = [r for r in px if not _is_zero_tslin(r.get("TSLIN", ""))]
    total = _sum_betrg(zero)
    dims = ["LGART", "PERNR", "SEQNO", "MOMAG", "POSTNUM", "ACTSIGN", "SWRETROACT", "SWPER"]
    out: dict[str, Any] = {
        "run": run,
        "rows": len(zero),
        "sum": str(total),
        "n_pernr": len({r.get("PERNR") for r in zero}),
        "n_lgart": len({r.get("LGART") for r in zero}),
        "non_zero_sum": str(_sum_betrg(non_zero)),
        "run_total": str(_sum_betrg(px)),
    }
    for d in dims:
        out[f"sum_by_{d.lower()}"] = _sum_by(zero, d, top=40)
    # pares D/C do mesmo LGART que se anulam (working set interno)
    lg = _sum_by(zero, "LGART", top=10_000)
    net_zero_lgart = sorted(k for k, v in lg.items() if _q(v["sum"]) == 0)
    residual_lgart = {k: v for k, v in lg.items() if _q(v["sum"]) != 0}
    out["lgart_net_zero"] = net_zero_lgart
    out["lgart_with_residual"] = dict(sorted(residual_lgart.items(),
                                             key=lambda kv: -abs(_q(kv[1]["sum"]))))
    out["observation"] = (
        "[OBSERVED] Os PPOIX com TSLIN=0 não são emitidos como linha FI própria. "
        "A maioria são pares débito/crédito do mesmo LGART que se anulam "
        "(working set interno do split); o resíduo por LGART (ver "
        "`lgart_with_residual`) é o que o programa redistribui pelas linhas "
        "efectivamente lançadas."
    )
    return out


def find_transferred_nontransferred_pairs(connection: Any, params: AnalysisParams, run: str,
                                          pernrs: Iterable[str] | None = None,
                                          run_ppoix: list[dict[str, str]] | None = None,
                                          feeder_linums: Iterable[str] | None = None,
                                          ) -> list[dict[str, Any]]:
    px = run_ppoix if run_ppoix is not None else _read_run_ppoix(connection, params, run)
    fset = set(feeder_linums or [])
    pset = set(pernrs) if pernrs is not None else None
    transferred: dict[tuple[str, str], list[dict[str, str]]] = {}
    zero: dict[tuple[str, str], list[dict[str, str]]] = {}
    for r in px:
        if pset is not None and r.get("PERNR") not in pset:
            continue
        key = (r.get("PERNR", ""), r.get("LGART", ""))
        if fset and r.get("TSLIN") in fset:
            transferred.setdefault(key, []).append(r)
        elif _is_zero_tslin(r.get("TSLIN", "")):
            zero.setdefault(key, []).append(r)
    pairs = []
    for key in sorted(set(transferred) & set(zero)):
        pn, lg = key
        t_sum = _sum_betrg(transferred[key])
        z_sum = _sum_betrg(zero[key])
        pairs.append({
            "pernr": pn, "lgart": lg,
            "transfer_rows": len(transferred[key]), "transfer_sum": str(t_sum),
            "zero_rows": len(zero[key]), "zero_sum": str(z_sum),
            "transfer_minus_zero": str(t_sum - z_sum),
            "transfer_minus_2x_zero": str(t_sum - 2 * z_sum),
        })
    pairs.sort(key=lambda d: -abs(_q(d["transfer_sum"])))
    return pairs


# ---------------------------------------------------------------------------
# 12 — mapa de uso de SEQNO
# ---------------------------------------------------------------------------

def build_seqno_usage_map(connection: Any, params: AnalysisParams, run: str,
                          run_ppoix: list[dict[str, str]] | None = None,
                          other_runs: Iterable[str] | None = None) -> dict[str, Any]:
    px = run_ppoix if run_ppoix is not None else _read_run_ppoix(connection, params, run)
    this = {(r.get("PERNR", ""), r.get("SEQNO", "")) for r in px}
    other_seen: dict[tuple[str, str], set[str]] = {}
    for rid in sorted(set(other_runs or [])):
        if rid == run:
            continue
        try:
            o = _read_ppoix_options(connection, params, opt_and(opt_eq("RUNID", rid)),
                                    max_rows=1_000_000)
        except (RfcReadError, NoData) as exc:
            logger.warning("SEQNO map: run %s indisponível: %s", rid, exc)
            continue
        for r in o:
            if r.get("RUNID") and r.get("RUNID") != rid:
                continue
            k = (r.get("PERNR", ""), r.get("SEQNO", ""))
            other_seen.setdefault(k, set()).add(rid)

    classified: dict[str, int] = {}
    sample: list[dict[str, Any]] = []
    for (pn, sq) in sorted(this):
        seen_in = other_seen.get((pn, sq), set())
        cls = "FIRST_SEEN" if not seen_in else "REUSED"
        classified[cls] = classified.get(cls, 0) + 1
        if len(sample) < 60:
            sample.append({"pernr": pn, "seqno": sq, "class": cls,
                           "also_in_runs": sorted(seen_in)})
    return {
        "run": run,
        "pernr_seqno_pairs": len(this),
        "compared_runs": sorted(set(other_runs or []) - {run}),
        "classification_counts": classified,
        "sample": sample,
        "note": (
            "[OBSERVED] SEQNO é o nº de sequência do resultado de Payroll por "
            "PERNR. 'REUSED' = o mesmo par PERNR+SEQNO existe também noutro run "
            "listado (posting do mesmo resultado). Sem PYXX não se lê o conteúdo "
            "do resultado — a classificação PREVIOUSLY_TRANSFERRED fica em "
            "[HYPOTHESIS]."
        ),
    }


# ---------------------------------------------------------------------------
# 13 / 14 — runs anteriores e valores previamente lançados
# ---------------------------------------------------------------------------

def find_previous_runs(connection: Any, params: AnalysisParams, run: str) -> dict[str, Any]:
    """Runs de posting HR da MESMA empresa/conta, por PPDHD (não assume 1297)."""
    try:
        hd = read_table(connection, "PPDHD",
                        fields=["RUNID", "DOCNUM", "BUKRS", "BUDAT", "BLDAT", "BLART", "XBLNR"],
                        page_size=200_000).rows
    except RfcReadError as exc:
        return {"error": f"PPDHD indisponível: {exc}"}
    same_company = sorted({r["RUNID"] for r in hd if r.get("BUKRS") == params.empresa})
    by_run = {}
    for r in hd:
        by_run.setdefault(r["RUNID"], {"bukrs": r.get("BUKRS", ""), "docs": [],
                                       "budat": r.get("BUDAT", "")})
        by_run[r["RUNID"]]["docs"].append(r["DOCNUM"])
    prior = sorted(x for x in same_company if x < run)
    later = sorted(x for x in same_company if x > run)
    return {
        "run": run,
        "company": params.empresa,
        "runs_same_company": same_company,
        "prior_runs_same_company": prior,
        "later_runs_same_company": later,
        "by_run": {k: by_run[k] for k in same_company if k in by_run},
        "conclusion": (
            f"[PROVED] A empresa {params.empresa} tem posting HR nos runs "
            f"{same_company}. "
            + ("NÃO há run anterior da mesma empresa para este período — "
               "netting contra run anterior não é aplicável."
               if not prior else
               f"Runs anteriores da mesma empresa: {prior}.")
        ),
    }


def trace_previous_postings(connection: Any, params: AnalysisParams,
                            prev_runs_info: dict[str, Any],
                            current_run: str = "",
                            current_ppoix: list[dict[str, str]] | None = None) -> dict[str, Any]:
    """Para cada run (prévio ou gémeo) da mesma empresa: o que foi lançado na
    conta alvo (PPDIT), e a soma PPOIX."""
    out: dict[str, Any] = {"per_run": {}}
    acct = params.conta.lstrip("0")
    for rid in prev_runs_info.get("runs_same_company", []):
        docs = prev_runs_info.get("by_run", {}).get(rid, {}).get("docs", [])
        if not docs:
            continue
        try:
            pit = read_table(connection, "PPDIT",
                             fields=["DOCNUM", "DOCLIN", "HKONT", "KTOSL", "WRBTR", "NEG_POSTNG"],
                             options=opt_and(opt_in("DOCNUM", docs)), page_size=100_000).rows
        except (RfcReadError, NoData):
            pit = []
        docset = set(docs)
        acct_lines = [r for r in pit
                      if r.get("DOCNUM") in docset
                      and r["HKONT"].lstrip("0") == acct and r.get("KTOSL") == params.hr_posting_key]
        if rid == current_run and current_ppoix is not None:
            px = current_ppoix
        else:
            try:
                px = _read_ppoix_options(connection, params, opt_and(opt_eq("RUNID", rid)),
                                         max_rows=1_000_000)
                px = [r for r in px if not r.get("RUNID") or r.get("RUNID") == rid]
            except (RfcReadError, NoData):
                px = []
        out["per_run"][rid] = {
            "docs": docs,
            "ppoix_rows": len(px),
            "ppoix_sum": str(_sum_betrg(px)),
            "account_lines": [
                {"docnum": r["DOCNUM"], "doclin": r["DOCLIN"],
                 "wrbtr": str(_signed(r["WRBTR"], r.get("NEG_POSTNG", "")))}
                for r in acct_lines
            ],
        }
    return out


# ---------------------------------------------------------------------------
# 17 / 18 / 19 — breakdowns
# ---------------------------------------------------------------------------

def analyze_lgart_breakdown(line_rows: list[dict[str, str]], wrbtr: Decimal) -> dict[str, Any]:
    by = _sum_by(line_rows, "LGART", top=10_000)
    total = _sum_betrg(line_rows)
    return {
        "by_lgart": by,
        "ppoix_sum": str(total),
        "ppdit_wrbtr": str(wrbtr),
        "delta": str(total - wrbtr),
        "note": "[PROVED] Composição por rubrica da linha alvo (soma PPOIX por TSLIN alimentador).",
    }


def analyze_pernr_breakdown(line_rows: list[dict[str, str]]) -> dict[str, Any]:
    per: dict[str, dict[str, Any]] = {}
    for r in line_rows:
        pn = r.get("PERNR", "")
        e = per.setdefault(pn, {"rows": 0, "sum": _ZERO, "seqnos": set(), "lgarts": set()})
        e["rows"] += 1
        e["sum"] += _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))
        e["seqnos"].add(r.get("SEQNO", ""))
        e["lgarts"].add(r.get("LGART", ""))
    rows = [{"pernr": pn, "rows": e["rows"], "sum": str(e["sum"]),
             "n_seqno": len(e["seqnos"]), "seqnos": sorted(e["seqnos"]),
             "lgarts": sorted(e["lgarts"])}
            for pn, e in per.items()]
    rows.sort(key=lambda d: -abs(_q(d["sum"])))
    multi = [r for r in rows if r["n_seqno"] > 1]
    return {
        "n_pernr": len(rows),
        "top": rows[:40],
        "multi_seqno_pernr": multi,
        "note": "[PROVED] Distribuição por PERNR das linhas da linha alvo.",
    }


def analyze_momag(connection: Any, params: AnalysisParams,
                  line_rows: list[dict[str, str]]) -> dict[str, Any]:
    by = _sum_by(line_rows, "MOMAG", top=100)
    # tentar semântica em T52EK/T030 é opcional; se não houver, MOMAG = UNKNOWN
    return {
        "by_momag": by,
        "semantics": "UNKNOWN",
        "note": (
            "[OBSERVED] MOMAG (account-assignment modifier) toma os valores "
            f"{sorted(by)} nas linhas da linha alvo. A semântica exacta não é "
            "confirmável só por RFC_READ_TABLE das tabelas de posting — fica "
            "UNKNOWN. Ambos os MOMAG desembocam na MESMA linha PPDIT."
        ),
    }


# ---------------------------------------------------------------------------
# 15 / 20 / 21 — candidatos ao delta, primeira divergência, classificação
# ---------------------------------------------------------------------------

def find_delta_candidates(line_rows: list[dict[str, str]], delta: Decimal,
                          *, zero_tslin_rows: list[dict[str, str]] | None = None) -> dict[str, Any]:
    """Procura relações TECNICAMENTE COERENTES (não brute force) cujo valor seja
    exactamente `delta`. Cada candidato é [CANDIDATE] até haver cadeia completa.
    """
    target = abs(delta)
    out: dict[str, Any] = {"delta": str(delta), "abs_delta": str(target), "candidates": []}

    # (a) linha única == delta
    singles = [r for r in line_rows
               if abs(_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))) == target]
    for r in singles:
        out["candidates"].append({
            "kind": "single_ppoix_row_equals_delta",
            "pernr": r.get("PERNR"), "seqno": r.get("SEQNO"), "lgart": r.get("LGART"),
            "betrg": str(_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))),
            "status": "[CANDIDATE] valor igual ao delta, mas a linha JÁ está "
                      "incluída na soma — excluí-la seria arbitrário sem regra.",
        })

    # (b) subconjunto coerente: mesmo PERNR+LGART, ou par claims /561 vs /563,
    #     ou mesma cadeia SEQNO. Só combinações pequenas e com nexo.
    non_559 = [r for r in line_rows if r.get("LGART") != "/559"]
    vals = [(f'{r.get("PERNR")}|{r.get("SEQNO")}|{r.get("LGART")}',
             _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))) for r in non_559]
    for k in range(1, min(len(vals), 6) + 1):
        for combo in combinations(vals, k):
            if abs(sum((v for _, v in combo), _ZERO) - target) < Decimal("0.005") \
               or abs(sum((v for _, v in combo), _ZERO) + target) < Decimal("0.005"):
                out["candidates"].append({
                    "kind": "coherent_subset_non_559",
                    "members": [lbl for lbl, _ in combo],
                    "sum": str(sum((v for _, v in combo), _ZERO)),
                    "status": "[CANDIDATE] subconjunto que soma o delta; só é "
                              "prova com nexo contabilístico (mesma linha, par "
                              "F/C, mesma cadeia retro).",
                })
                if len(out["candidates"]) > 30:
                    break

    # (c) fracções de cêntimo
    sub_cent = [r for r in line_rows
                if _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")) !=
                _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")).quantize(_CENT)]
    out["sub_cent_rows"] = len(sub_cent)

    out["note"] = (
        "Regras de coincidência numérica NÃO são causa. Um candidato só passa a "
        "[PROVED] com cadeia técnica completa (PERNR/LGART/SEQNO/POSTNUM/MOMAG/"
        "FPPER coerentes ligando o valor à transformação)."
    )
    return out


def find_first_divergence(checkpoints: list[PostingCheckpoint]) -> dict[str, Any]:
    """Percorre os checkpoints com valor por ordem e reporta o primeiro estágio
    onde o montante muda."""
    valued = [c for c in checkpoints if c.amount is not None]
    if len(valued) < 2:
        return {"result": "INSUFFICIENT_CHECKPOINTS"}
    prev = valued[0]
    for cur in valued[1:]:
        if cur.amount != prev.amount:
            return {
                "result": "DIVERGES",
                "between": f"{prev.stage} -> {cur.stage}",
                "from_amount": str(prev.amount),
                "to_amount": str(cur.amount),
                "delta": str((cur.amount or _ZERO) - (prev.amount or _ZERO)),
                "skipped_stages_without_value": [
                    c.stage for c in checkpoints
                    if c.amount is None and prev.stage < c.stage < cur.stage
                ],
                "evidence": (
                    "[PROVED] O montante persistido só existe em PPOIX e PPDIT; "
                    "entre eles não há checkpoint com valor (PPOPX/PPDIX sem "
                    "campo monetário; PPDST vazio). A transformação ocorre dentro "
                    "do programa de posting HR (RPCIPE00/SAPLHRPP) ao construir a "
                    "linha do documento — não é observável como registo."
                ),
            }
        prev = cur
    return {"result": "NO_DIVERGENCE", "amount": str(prev.amount)}


def classify_delta_origin(trace: PostingDeltaTrace) -> dict[str, Any]:
    delta = trace.delta
    rec = trace.reconciliation
    ppopx = trace.ppopx
    prev = trace.previous_runs
    cands = trace.delta_candidates

    # PPDST decompõe a linha?
    ppdst_rows = 0
    for c in trace.checkpoints:
        if c.source == "PPDST":
            ppdst_rows = c.row_count
    if ppdst_rows and c.amount is not None and c.amount == trace.ppdit_wrbtr:
        return _cls("PROVED_AT_INTERMEDIATE_STAGE",
                    "PPDST contém o split da linha PPDIT com o mesmo total — a "
                    "transformação é observável nesse estágio.")

    # PPOPX fecha o delta?
    overlap = (ppopx.get("overlap_with_target_line") or {})
    if overlap:
        for label, v in overlap.items():
            if v.get("matched_rows") and abs(_q(v["matched_sum"]) - delta) < _CENT:
                return _cls("PROVED_AT_PPOPX",
                            f"As linhas PPOIX com correspondência em PPOPX ({label}) "
                            f"somam exactamente o delta ({v['matched_sum']}).")

    # Netting contra run anterior da mesma empresa?
    if prev.get("prior_runs_same_company"):
        pp = trace.previous_postings.get("per_run", {})
        for rid in prev["prior_runs_same_company"]:
            for al in pp.get(rid, {}).get("account_lines", []):
                if abs(_q(al["wrbtr"]) - delta) < _CENT:
                    return _cls("PROVED_PREVIOUS_RUN_NETTING",
                                f"O run anterior {rid} lançou {al['wrbtr']} na conta "
                                "alvo, igual ao delta — netting confirmado.")
        # há run anterior mas não fecha
        note_prev = "Há run(s) anterior(es) da mesma empresa mas nenhum lança um valor igual ao delta."
    else:
        note_prev = ("[PROVED] Não existe run de posting anterior da empresa "
                     f"{trace.company} — netting contra run anterior está EXCLUÍDO.")

    # Candidato coerente exacto?
    proven = [c for c in cands.get("candidates", []) if c.get("status", "").startswith("[PROVED]")]
    if proven:
        return _cls("PROVED_OTHER_RULE",
                    "Candidato com cadeia técnica completa: " + json.dumps(proven[0], ensure_ascii=False))

    # Fracções de cêntimo que expliquem arredondamento?
    if cands.get("sub_cent_rows"):
        return _cls("PARTIALLY_EXPLAINED",
                    "Existem linhas PPOIX com fracção de cêntimo — parte do delta "
                    "pode ser arredondamento na agregação (confirmar caso a caso).")

    # Nada fecha => a divergência está provadamente entre PPOIX e PPDIT
    fd = trace.first_divergence or {}
    if fd.get("result") == "DIVERGES" and fd.get("between", "").endswith("4-PPDIT"):
        return _cls(
            "PROVED_BETWEEN_PPOIX_AND_PPDIT",
            "O delta é introduzido na construção da linha do documento pelo "
            "programa de posting HR. Não há tabela intermédia com montante "
            "(PPDIX/PPOPX sem campo de valor; PPDST vazio); PPOPX sem overlap; "
            f"{note_prev} "
            "A regra exacta que produz precisamente este valor não é "
            "reconstruível a partir das tabelas transparentes acessíveis por "
            "RFC_READ_TABLE — residual_rule = UNEXPLAINED. "
            "Contexto [OBSERVED]: o delta desta linha é uma fracção da diferença "
            "de todo o documento entre 'SUM(PPOIX) por linha de transferência' e "
            "'PPDIT.WRBTR', diferença essa espelhada pelos PPOIX com "
            "TSLIN in {0,17,...} (working set do split que o programa "
            "redistribui e não emite como linha própria).",
        )
    return _cls("UNEXPLAINED",
                "Não foi possível localizar o estágio nem a regra com os dados "
                "read-only disponíveis.")


def _cls(klass: str, why: str) -> dict[str, Any]:
    assert klass in DELTA_ORIGIN_CLASSES, klass
    return {"classification": klass, "rationale": why}


# ---------------------------------------------------------------------------
# helpers de agregação
# ---------------------------------------------------------------------------

def _count(rows: Iterable[dict[str, str]], field_name: str) -> dict[str, int]:
    out: dict[str, int] = {}
    for r in rows:
        k = r.get(field_name, "")
        out[k] = out.get(k, 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


def _sum_by(rows: Iterable[dict[str, str]], field_name: str, *, top: int = 25) -> dict[str, Any]:
    agg: dict[str, list[Any]] = {}
    for r in rows:
        k = r.get(field_name, "")
        e = agg.setdefault(k, [0, _ZERO])
        e[0] += 1
        e[1] += _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))
    items = sorted(agg.items(), key=lambda kv: -abs(kv[1][1]))[:top]
    return {k: {"rows": v[0], "sum": str(v[1])} for k, v in items}


# ---------------------------------------------------------------------------
# Orquestrador
# ---------------------------------------------------------------------------

def trace_posting_delta(connection: Any, params: AnalysisParams, *, docnum: str, doclin: str,
                        run: str | None = None) -> PostingDeltaTrace:
    run = (run or params.primary_run_10)
    docnum = str(docnum).strip().zfill(10)
    doclin = str(doclin).strip().zfill(10)

    tr = PostingDeltaTrace(run_id=run, company=params.empresa, account=params.conta_10,
                           docnum=docnum, doclin=doclin)

    px = _read_run_ppoix(connection, params, run)
    linums, _hits = _feeders_for_line(connection, params, run, docnum, doclin)
    tr.feeder_linums = linums
    line_rows = [r for r in px if r.get("TSLIN") in set(linums)]
    tr.ppoix_rows = len(line_rows)
    tr.ppoix_sum = _sum_betrg(line_rows)

    pit = _ppdit_line(connection, params, docnum, doclin)
    tr.ppdit_wrbtr = _signed(pit.get("WRBTR", "0"), pit.get("NEG_POSTNG", "")) if pit else _ZERO
    tr.delta = tr.ppoix_sum - tr.ppdit_wrbtr

    if not line_rows:
        tr.warn(f"Sem PPOIX para run {run} / linha {docnum}/{doclin} (LINUMs {linums}).")

    # 5 / 6
    tr.intermediates = discover_posting_intermediates(connection, params)
    tr.ppopx = inspect_ppopx(connection, params, run, target_rows=line_rows)

    # PPDST — decompõe a linha?
    ppdst_amount: Decimal | None = None
    ppdst_rows = 0
    try:
        ds = read_table(connection, "PPDST",
                        fields=["DOCNUM", "DOCLIN", "WRBTR", "WAERS", "CURTP"],
                        options=opt_and(opt_eq("DOCNUM", docnum)), page_size=100_000).rows
        ds = [r for r in ds if r.get("DOCNUM") == docnum and r["DOCLIN"] == doclin]
        ppdst_rows = len(ds)
        if ds:
            ppdst_amount = sum((_q(r.get("WRBTR", "0")) for r in ds), _ZERO)
    except (RfcReadError, NoData) as exc:
        tr.warn(f"PPDST não lida: {exc}")

    # 8
    tr.checkpoints = build_posting_checkpoints(tr, tr.intermediates, tr.ppopx,
                                               ppdst_amount, ppdst_rows)
    # 4 / 9
    tr.reconciliation = reconcile_target_line(connection, params, run, docnum, doclin, run_ppoix=px)
    # 10 / 11
    tr.zero_tslin = analyze_zero_tslin(connection, params, run, run_ppoix=px)
    tr.tslin_pairs = find_transferred_nontransferred_pairs(
        connection, params, run, pernrs={r.get("PERNR") for r in line_rows},
        run_ppoix=px, feeder_linums=linums)
    # 12 — comparação SEQNO limitada aos runs da MESMA empresa (barato);
    # a comparação alargada fica no comando --trace-seqno-history.
    prev = find_previous_runs(connection, params, run)
    tr.previous_runs = prev
    compare_runs = set(prev.get("runs_same_company", []))
    tr.seqno_usage = build_seqno_usage_map(connection, params, run, run_ppoix=px,
                                           other_runs=compare_runs)
    # 13 / 14
    tr.previous_postings = trace_previous_postings(connection, params, prev,
                                                   current_run=run, current_ppoix=px)
    # 17 / 18 / 19
    tr.lgart_breakdown = analyze_lgart_breakdown(line_rows, tr.ppdit_wrbtr)
    tr.pernr_breakdown = analyze_pernr_breakdown(line_rows)
    tr.momag = analyze_momag(connection, params, line_rows)
    # 15
    tr.delta_candidates = find_delta_candidates(
        line_rows, tr.delta,
        zero_tslin_rows=[r for r in px if _is_zero_tslin(r.get("TSLIN", ""))])
    # 20
    tr.first_divergence = find_first_divergence(tr.checkpoints)
    # 21
    tr.classification = classify_delta_origin(tr)
    return tr


# ---------------------------------------------------------------------------
# Comandos auxiliares (--analyze-zero-tslin, --trace-seqno-history)
# ---------------------------------------------------------------------------

def analyze_zero_tslin_standalone(connection: Any, params: AnalysisParams,
                                  run: str | None = None) -> dict[str, Any]:
    run = run or params.primary_run_10
    return analyze_zero_tslin(connection, params, run)


def trace_seqno_history_standalone(connection: Any, params: AnalysisParams,
                                   run: str | None = None) -> dict[str, Any]:
    run = run or params.primary_run_10
    px = _read_run_ppoix(connection, params, run)
    compare = set(params.posting_runs)
    prev = find_previous_runs(connection, params, run)
    compare |= set(prev.get("runs_same_company", []))
    usage = build_seqno_usage_map(connection, params, run, run_ppoix=px, other_runs=compare)
    usage["previous_runs"] = prev
    return usage


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_posting_delta_json(tr: PostingDeltaTrace, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(tr.as_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info("JSON escrito: %s", path)
    return path


def write_posting_delta_csvs(tr: PostingDeltaTrace, output_dir: Path, run: str,
                             docnum: str, doclin: str) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{run}_{int(docnum)}_{int(doclin)}"
    written: list[Path] = []

    p = output_dir / f"posting_delta_items_{stem}.csv"
    with p.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["stage", "source", "key", "row_count", "amount", "currency",
                    "derivation", "evidence"])
        for c in tr.checkpoints:
            d = c.as_dict()
            w.writerow([d["stage"], d["source"], d["key"], d["row_count"],
                        d["amount"] or "", d["currency"], d["derivation"], d["evidence"]])
        for ln in tr.reconciliation.get("per_linum", []):
            w.writerow(["LINUM", "PPOIX", ln["linum"], ln["rows"], ln["sum"], "EUR",
                        f"by_lgart={ln['by_lgart']}", "[PROVED]"])
    written.append(p)

    p = output_dir / f"zero_tslin_{run}.csv"
    with p.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["dimension", "value", "rows", "sum"])
        for dim in ("lgart", "pernr", "seqno", "momag", "postnum", "actsign",
                    "swretroact", "swper"):
            for k, v in (tr.zero_tslin.get(f"sum_by_{dim}", {}) or {}).items():
                w.writerow([dim, k, v["rows"], v["sum"]])
    written.append(p)

    p = output_dir / f"seqno_history_{run}.csv"
    with p.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["pernr", "seqno", "class", "also_in_runs"])
        for s in tr.seqno_usage.get("sample", []):
            w.writerow([s["pernr"], s["seqno"], s["class"], ",".join(s["also_in_runs"])])
    written.append(p)

    p = output_dir / f"previous_run_trace_{run}.csv"
    with p.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["runid", "ppoix_rows", "ppoix_sum", "account_doclin", "account_wrbtr"])
        for rid, info in tr.previous_postings.get("per_run", {}).items():
            if info.get("account_lines"):
                for al in info["account_lines"]:
                    w.writerow([rid, info["ppoix_rows"], info["ppoix_sum"],
                                al["doclin"], al["wrbtr"]])
            else:
                w.writerow([rid, info["ppoix_rows"], info["ppoix_sum"], "", ""])
    written.append(p)
    return written


def _fmt(v: Any) -> str:
    if v in (None, ""):
        return "(n/d)"
    try:
        q = Decimal(str(v)).quantize(_CENT)
    except Exception:  # noqa: BLE001
        return str(v)
    s = f"{abs(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if q < 0 else ''}{s}"


def print_posting_delta_report(tr: PostingDeltaTrace) -> None:
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # pragma: no cover  # noqa: BLE001
        pass
    L = "=" * 68
    print(L)
    print("POSTING DELTA TRACE")
    print(L)
    print(f"Run ............ {tr.run_id}")
    print(f"Empresa/Conta .. {tr.company} / {tr.account.lstrip('0')}")
    print(f"Linha PPDIT .... {tr.docnum}/{tr.doclin}")
    print(f"LINUMs feeder .. {', '.join(tr.feeder_linums) or '(nenhum)'}")
    print("")
    print(f"  SUM PPOIX (linhas de transferência) . {_fmt(tr.ppoix_sum)}  ({tr.ppoix_rows} linhas)")
    print(f"  PPDIT.WRBTR ......................... {_fmt(tr.ppdit_wrbtr)}")
    print(f"  DELTA .............................. {_fmt(tr.delta)}")
    print("")
    print("-" * 68)
    print("CHECKPOINTS")
    print("-" * 68)
    for c in tr.checkpoints:
        amt = _fmt(c.amount) if c.amount is not None else "(sem valor persistido)"
        print(f"  {c.stage:<9} {c.source:<7} linhas={c.row_count:<5} {amt:>18}  {c.evidence}")
        print(f"            {c.derivation}")
    print("")
    print("-" * 68)
    print("PPOPX")
    print("-" * 68)
    print(f"  linhas no run ....... {tr.ppopx.get('rows_for_run', 'n/d')}")
    print(f"  campo monetário .... {tr.ppopx.get('has_money_field')}")
    for label, v in (tr.ppopx.get("overlap_with_target_line") or {}).items():
        print(f"  overlap [{label}] : {v['matched_rows']} linhas  soma {_fmt(v['matched_sum'])}")
    if tr.ppopx.get("conclusion"):
        print(f"  => {tr.ppopx['conclusion']}")
    print("")
    print("-" * 68)
    print("LINUM (reconciliação por linha de transferência)")
    print("-" * 68)
    for ln in tr.reconciliation.get("per_linum", []):
        print(f"  LINUM {ln['linum']}  {ln['rows']:>4} linhas  soma {_fmt(ln['sum'])}")
        for lg, info in ln["by_lgart"].items():
            print(f"       {lg:<8} {info['rows']:>4}  {_fmt(info['sum']):>16}")
        print(f"       MOMAG: " + ", ".join(f"{k or '-'}={_fmt(v['sum'])}"
                                            for k, v in ln["by_momag"].items()))
    print("")
    print("-" * 68)
    print("TSLIN ZERO (working set não emitido como linha)")
    print("-" * 68)
    z = tr.zero_tslin
    print(f"  linhas={z.get('rows')}  soma={_fmt(z.get('sum'))}  "
          f"PERNR={z.get('n_pernr')}  LGART={z.get('n_lgart')}")
    print(f"  LGART que se anulam (net 0): {len(z.get('lgart_net_zero', []))}")
    print("  LGART com resíduo (redistribuído):")
    for k, v in list((z.get("lgart_with_residual") or {}).items())[:12]:
        print(f"     {k:<8} {v['rows']:>4}  {_fmt(v['sum']):>16}")
    print("")
    print("-" * 68)
    print("SEQNO HISTORY")
    print("-" * 68)
    print(f"  pares PERNR+SEQNO no run : {tr.seqno_usage.get('pernr_seqno_pairs')}")
    print(f"  runs comparados ......... {tr.seqno_usage.get('compared_runs')}")
    print(f"  classificação ........... {tr.seqno_usage.get('classification_counts')}")
    print("")
    print("-" * 68)
    print("PREVIOUS RUNS")
    print("-" * 68)
    pr = tr.previous_runs
    print(f"  runs da empresa {tr.company}: {pr.get('runs_same_company')}")
    print(f"  anteriores .............. {pr.get('prior_runs_same_company')}")
    print(f"  => {pr.get('conclusion', '')}")
    for rid, info in tr.previous_postings.get("per_run", {}).items():
        al = "; ".join(f"{x['doclin']}={_fmt(x['wrbtr'])}" for x in info.get("account_lines", []))
        print(f"     {rid}: PPOIX {info['ppoix_rows']} l. soma {_fmt(info['ppoix_sum'])}"
              + (f" | conta alvo: {al}" if al else ""))
    print("")
    print("-" * 68)
    print("DELTA — candidatos e classificação")
    print("-" * 68)
    for c in tr.delta_candidates.get("candidates", [])[:10]:
        print(f"  - {c['kind']}: {c.get('status', '')}")
    print(f"  sub-cent rows .......... {tr.delta_candidates.get('sub_cent_rows')}")
    fd = tr.first_divergence
    print("")
    print(f"  PRIMEIRA DIVERGÊNCIA: {fd.get('result')}"
          + (f"  ({fd.get('between')}: {_fmt(fd.get('from_amount'))} -> {_fmt(fd.get('to_amount'))}, "
             f"delta {_fmt(fd.get('delta'))})" if fd.get('result') == 'DIVERGES' else ""))
    if fd.get("evidence"):
        print(f"    {fd['evidence']}")
    print("")
    print(f"  >>> CLASSIFICAÇÃO: {tr.classification.get('classification')}")
    print(f"      {tr.classification.get('rationale', '')}")
    print("")
    if tr.warnings:
        print("AVISOS:")
        for w in tr.warnings:
            print(f"  - {w}")
    print(L)
