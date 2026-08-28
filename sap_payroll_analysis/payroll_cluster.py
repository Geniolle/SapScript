"""Fase 3 — contexto de Payroll (automático), RGDIR, timeline de resultados,
pares original→recalculado, e tentativa de leitura da RT.

Tudo é obtido automaticamente por `RFC_READ_TABLE` sobre tabelas
transparentes. Nada é pedido ao utilizador. Só se classifica um dado como
`MANUAL_REQUIRED` quando está efectivamente dentro do cluster PCL2 e a
função de leitura não é RFC-enabled.

Descoberto no sistema real (ECC, 10.1.1.101/100, empresa 1010):

* MOLGA 19 (Portugal), ABKRS Z2, PERMO 01 (mensal), RELID PCL2 = ``RP``.
* **RGDIR**: tabela transparente ``HRPY_RGDIR`` — lida na íntegra por RFC.
* **RT / CRT / BT / WPBP**: existem tabelas transparentes ``P2RX_*`` /
  ``HRPADNLP_P2RX_*`` no DDIC deste sistema mas estão **vazias** (o framework
  "Payroll Results Tables" não está activo). A RT vive só no cluster
  ``PCL2(RP)``.
* ``HR_GET_PAYROLL_RESULTS`` não é RFC-enabled; ``PYXX_READ_PAYROLL_RESULT``
  devolve ``DA300 «No active nametab»`` em contexto RFC stateless.

Conclusão: os montantes por rubrica da RT continuam `MANUAL_REQUIRED`
(cluster). Tudo o resto — cadeia PPDHD→…→PERNR, RGDIR completo, timeline,
pares de recálculo, catálogo de tabelas — é automático.
"""

from __future__ import annotations

import logging
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from .config import (
    AnalysisParams,
    PAYROLL_PERIOD_YYYYMM,
    READ_ONLY_TABLE_WHITELIST,
    pad_run,
)
from .ddic import describe_table
from .sap_reader import NoData, RfcReadError, opt_and, opt_eq, opt_in, read_table, sap_str_to_decimal
from .security import safe_rfc_call

logger = logging.getLogger(__name__)

RT_READ_FUNCTIONS = ("PYXX_READ_PAYROLL_RESULT",)


def _months_between(a: str, b: str) -> int:
    try:
        ay, am = int(a[:4]), int(a[4:6])
        by, bm = int(b[:4]), int(b[4:6])
    except (ValueError, IndexError):
        return 0
    return (by * 12 + bm) - (ay * 12 + am)

#: Tabelas transparentes candidatas a conter resultados de Payroll (descoberta
#: automática confirma quais existem e quais têm dados neste sistema).
PAYROLL_RESULT_TABLE_CANDIDATES = (
    "HRPY_RGDIR", "HRPY_RGDIR_TEMP", "HRPY_WPBP", "HRPY_GROUPING",
    "P2RX_RT", "P2RX_RT_PERSON", "P2RX_CRT", "P2RX_BT", "P2RX_WPBP",
    "P2RX_VERSC", "P2RX_ARRRS", "P2RX_DDNTK", "P2RX_GRT",
    "HRPADNLP_P2RX_RT", "HRPADNLP_P2RX_BT",
)


# ===========================================================================
# Modelos
# ===========================================================================

@dataclass
class PayrollDirectoryEntry:
    """Uma entrada do RGDIR (HRPY_RGDIR)."""

    pernr: str
    seqnr: str
    abkrs: str = ""
    fpper: str = ""   # FOR period YYYYMM
    fpbeg: str = ""
    fpend: str = ""
    inper: str = ""   # IN period YYYYMM
    ipend: str = ""
    srtza: str = ""   # A=actual/current  P=previous  O=old
    payty: str = ""   # ''=regular  senão off-cycle
    payid: str = ""
    void: str = ""
    reversal: str = ""
    outofseq: str = ""
    ocrsn: str = ""
    bondt: str = ""
    rundt: str = ""
    permo: str = ""

    @property
    def is_retro(self) -> bool:
        return bool(self.fpper) and bool(self.inper) and self.fpper < self.inper

    @property
    def months_late(self) -> int:
        """Nº de meses entre FOR-period e IN-period (0 = calculado no próprio mês)."""
        return _months_between(self.fpper, self.inper)

    @property
    def is_offcycle(self) -> bool:
        return bool(self.payty.strip())

    @property
    def is_void(self) -> bool:
        return self.void.strip().upper() in {"X", "V"} or self.reversal.strip().upper() == "X"

    def classify(self) -> str:
        """Classificação da entrada (independente do período do posting).

        `RETRO_LAG` = recálculo de rotina 1 mês depois (esta folha corre com
        desfasamento sistemático de 1 período). `RETRO_CORR` = recálculo
        posterior a esse desfasamento (correcção real).
        """
        if self.is_void:
            return "VOID/REVERSAL"
        if self.is_offcycle:
            base = "OFF_CYCLE"
        elif self.months_late >= 2:
            base = "RETRO_CORR"
        elif self.months_late == 1:
            base = "RETRO_LAG"
        else:
            base = "ORIGINAL"
        state = {"A": "CURRENT", "P": "PREVIOUS", "O": "OLD"}.get(self.srtza.strip().upper(), self.srtza)
        return f"{base}/{state}"

    def as_dict(self) -> dict[str, Any]:
        return {
            "pernr": self.pernr, "seqnr": self.seqnr, "abkrs": self.abkrs,
            "fpper": self.fpper, "fpbeg": self.fpbeg, "fpend": self.fpend,
            "inper": self.inper, "ipend": self.ipend, "srtza": self.srtza,
            "payty": self.payty, "payid": self.payid, "void": self.void,
            "reversal": self.reversal, "outofseq": self.outofseq, "ocrsn": self.ocrsn,
            "rundt": self.rundt,
            "retro": self.is_retro, "offcycle": self.is_offcycle, "voided": self.is_void,
            "classification": self.classify(),
        }


@dataclass
class ResultPair:
    """Evolução dos SEQNR de um `(PERNR, FPPER)` ao longo dos IN-periods."""

    pernr: str
    fpper: str
    entries: list[PayrollDirectoryEntry] = field(default_factory=list)

    @property
    def in_periods(self) -> list[str]:
        return sorted({e.inper for e in self.entries})

    @property
    def original(self) -> PayrollDirectoryEntry | None:
        cand = [e for e in self.entries if e.inper == self.fpper and not e.is_void]
        return sorted(cand, key=lambda e: e.seqnr)[0] if cand else None

    @property
    def current(self) -> PayrollDirectoryEntry | None:
        cand = [e for e in self.entries if e.srtza == "A" and not e.is_void]
        return sorted(cand, key=lambda e: e.seqnr)[-1] if cand else None

    @property
    def status(self) -> str:
        recalcs = len(self.in_periods)
        if any(e.is_void for e in self.entries) and recalcs <= 1:
            return "RESULT_VOIDED"
        if self.original is None:
            return "RESULT_CURRENT_ONLY"
        if recalcs <= 1:
            return "RESULT_UNCHANGED"
        return "RESULT_RECALCULATED"

    def as_dict(self) -> dict[str, Any]:
        o, c = self.original, self.current
        return {
            "pernr": self.pernr, "fpper": self.fpper,
            "in_periods": self.in_periods, "recalc_count": len(self.in_periods),
            "original_seqnr": o.seqnr if o else None,
            "original_inper": o.inper if o else None,
            "current_seqnr": c.seqnr if c else None,
            "current_inper": c.inper if c else None,
            "status": self.status,
        }


@dataclass
class PayrollTimeline:
    pernr: str
    entries: list[PayrollDirectoryEntry] = field(default_factory=list)
    pairs: list[ResultPair] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "pernr": self.pernr,
            "entries": [e.as_dict() for e in self.entries],
            "pairs": [p.as_dict() for p in self.pairs],
        }


@dataclass
class HrpyTableInfo:
    table: str
    table_class: str = ""
    description: str = ""
    exists: bool = False
    accessible: bool = False   # legível por RFC_READ_TABLE
    populated: bool | None = None
    field_count: int = 0
    note: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "table": self.table, "class": self.table_class, "description": self.description,
            "exists": self.exists, "accessible": self.accessible, "populated": self.populated,
            "field_count": self.field_count, "note": self.note,
        }


@dataclass
class RtReadAttempt:
    function: str = ""
    attempted: bool = False
    ok: bool = False
    reason: str = ""
    detail: str = ""
    sample: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {"function": self.function, "attempted": self.attempted, "ok": self.ok,
                "reason": self.reason, "detail": self.detail, "sample": self.sample}


@dataclass
class PayrollContext:
    """Resultado de `collect_payroll_context` — a cadeia completa, automática."""

    run_id: str = ""
    company: str = ""
    account: str = ""
    in_period: str = ""
    molga: str = ""
    abkrs: str = ""
    permo: str = ""
    relid: str = ""
    doc_lines: list[tuple[str, str]] = field(default_factory=list)
    transfer_linums: list[str] = field(default_factory=list)
    pernrs: list[str] = field(default_factory=list)
    ppoix_by_pernr_wt: dict[str, dict[str, str]] = field(default_factory=dict)
    rgdir_by_pernr: dict[str, list[PayrollDirectoryEntry]] = field(default_factory=dict)
    pa_by_pernr: dict[str, dict[str, str]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)


@dataclass
class PayrollClusterReport:
    run_id: str = ""
    period: str = PAYROLL_PERIOD_YYYYMM
    company: str = ""
    molga: str = ""
    abkrs: str = ""
    permo: str = ""
    relid: str = ""

    pernr_count: int = 0
    rgdir_entries: list[PayrollDirectoryEntry] = field(default_factory=list)
    rgdir_for_inper: list[PayrollDirectoryEntry] = field(default_factory=list)

    current_pernr: list[str] = field(default_factory=list)
    retro_pernr: list[str] = field(default_factory=list)
    fpper_distribution: dict[str, int] = field(default_factory=dict)
    srtza_distribution: dict[str, int] = field(default_factory=dict)
    classification_distribution: dict[str, int] = field(default_factory=dict)
    offcycle_count: int = 0
    void_count: int = 0
    retro_months_hist: dict[str, int] = field(default_factory=dict)  # nº de PERNR por nº de meses retro

    ppoix_ref_by_pernr: dict[str, str] = field(default_factory=dict)
    ppoix_ref_total: Decimal = Decimal("0")
    ppoix_ref_retro_total: Decimal = Decimal("0")
    ppoix_ref_current_total: Decimal = Decimal("0")
    ppoix_ref_unclassified_total: Decimal = Decimal("0")

    timelines: list[PayrollTimeline] = field(default_factory=list)
    recalc_pairs: list[dict[str, Any]] = field(default_factory=list)
    ppoix_rgdir_view: list[dict[str, Any]] = field(default_factory=list)
    hrpy_catalog: list[HrpyTableInfo] = field(default_factory=list)

    rt_attempt: RtReadAttempt = field(default_factory=RtReadAttempt)
    run_1299_comparison: dict[str, Any] = field(default_factory=dict)
    residual_notes: dict[str, str] = field(default_factory=dict)
    per_pernr_diag: list[dict[str, Any]] = field(default_factory=list)

    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)

    def as_dict(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id, "period": self.period, "company": self.company,
            "molga": self.molga, "abkrs": self.abkrs, "permo": self.permo, "relid": self.relid,
            "pernr_count": self.pernr_count,
            "current_pernr_count": len(self.current_pernr),
            "retro_pernr_count": len(self.retro_pernr),
            "fpper_distribution": self.fpper_distribution,
            "srtza_distribution": self.srtza_distribution,
            "classification_distribution": self.classification_distribution,
            "retro_months_hist": self.retro_months_hist,
            "offcycle_count": self.offcycle_count,
            "void_count": self.void_count,
            "ppoix_ref_total": str(self.ppoix_ref_total),
            "ppoix_ref_retro_total": str(self.ppoix_ref_retro_total),
            "ppoix_ref_current_total": str(self.ppoix_ref_current_total),
            "ppoix_ref_unclassified_total": str(self.ppoix_ref_unclassified_total),
            # `recalc_pairs` só os que alimentam o run (CSV tem tudo)
            "recalc_pairs": [p for p in self.recalc_pairs if p.get("contributes_to_run")][:5000],
            "recalc_pairs_total": len(self.recalc_pairs),
            "ppoix_rgdir_view": self.ppoix_rgdir_view,
            "hrpy_catalog": [t.as_dict() for t in self.hrpy_catalog],
            "rt_attempt": self.rt_attempt.as_dict(),
            "run_1299_comparison": self.run_1299_comparison,
            "residual_notes": self.residual_notes,
            # timelines completas -> só no CSV; aqui uma amostra
            "timelines_sample": [t.as_dict() for t in self.timelines[:15]],
            "timelines_total": len(self.timelines),
            "rgdir_for_inper_count": len(self.rgdir_for_inper),
            "per_pernr_diag": self.per_pernr_diag,
            "resolved": self.resolved,
            "warnings": self.warnings,
        }


# ===========================================================================
# RGDIR — leitura completa (validada por DDIC)
# ===========================================================================

_RGDIR_WANT = ["PERNR", "SEQNR", "ABKRS", "FPPER", "FPBEG", "FPEND", "INPER", "IPEND",
               "SRTZA", "PAYTY", "PAYID", "VOID", "REVERSAL", "OUTOFSEQ", "OCRSN",
               "BONDT", "RUNDT", "PERMO"]


def read_rgdir(connection: Any, pernrs: list[str], *, since: str = "",
               page_size: int = 100_000) -> list[PayrollDirectoryEntry]:
    """Lê o RGDIR dos PERNR indicados (DDIC-validado).

    `since` (YYYYMM) limita a `FPPER >= since OR INPER >= since` — evita
    arrastar anos de histórico irrelevante para a cadeia de retro.
    """
    diag = describe_table(connection, "HRPY_RGDIR")
    if not (diag.exists and diag.authorized):
        raise RfcReadError(f"HRPY_RGDIR indisponível: {diag.note}", table="HRPY_RGDIR")
    avail = {n.upper() for n in diag.field_names()}
    fields = [f for f in _RGDIR_WANT if f in avail] or ["PERNR", "SEQNR", "FPPER", "INPER", "SRTZA"]
    window = [{"TEXT": "("}, {"TEXT": f"FPPER >= '{since}'"}, {"TEXT": "OR"},
              {"TEXT": f"INPER >= '{since}'"}, {"TEXT": ")"}] if since else []

    rows: list[dict[str, str]] = []
    for start in range(0, len(pernrs), 100):
        chunk = pernrs[start : start + 100]
        try:
            rows.extend(read_table(connection, "HRPY_RGDIR", fields=fields,
                                   options=opt_and(opt_in("PERNR", chunk), window),
                                   page_size=page_size).rows)
        except NoData:
            continue
    return [
        PayrollDirectoryEntry(
            pernr=r.get("PERNR", ""), seqnr=r.get("SEQNR", ""), abkrs=r.get("ABKRS", ""),
            fpper=r.get("FPPER", ""), fpbeg=r.get("FPBEG", ""), fpend=r.get("FPEND", ""),
            inper=r.get("INPER", ""), ipend=r.get("IPEND", ""), srtza=r.get("SRTZA", ""),
            payty=r.get("PAYTY", ""), payid=r.get("PAYID", ""), void=r.get("VOID", ""),
            reversal=r.get("REVERSAL", ""), outofseq=r.get("OUTOFSEQ", ""), ocrsn=r.get("OCRSN", ""),
            bondt=r.get("BONDT", ""), rundt=r.get("RUNDT", ""), permo=r.get("PERMO", ""),
        )
        for r in rows
    ]


def rgdir_window_start(in_period: str, months_back: int = 18) -> str:
    """YYYYMM `months_back` meses antes de `in_period` (para limitar o RGDIR)."""
    try:
        y, m = int(in_period[:4]), int(in_period[4:6])
    except ValueError:
        return ""
    idx = y * 12 + (m - 1) - months_back
    return f"{idx // 12:04d}{idx % 12 + 1:02d}"


def build_timeline(pernr: str, entries: list[PayrollDirectoryEntry]) -> PayrollTimeline:
    """Ordena as entradas de um PERNR e agrupa em pares por FOR-period."""
    ordered = sorted(entries, key=lambda e: (e.fpper, e.inper, e.seqnr))
    tl = PayrollTimeline(pernr=pernr, entries=ordered)
    by_fpper: dict[str, list[PayrollDirectoryEntry]] = defaultdict(list)
    for e in ordered:
        by_fpper[e.fpper].append(e)
    tl.pairs = [ResultPair(pernr=pernr, fpper=fp, entries=rows)
                for fp, rows in sorted(by_fpper.items())]
    return tl


# ===========================================================================
# Contexto de Payroll (automático)
# ===========================================================================

def discover_payroll_context(connection: Any, params: AnalysisParams, pernrs: list[str]) -> dict[str, str]:
    ctx = {"molga": "", "abkrs": "", "permo": "", "relid": "", "land1": ""}
    sample = pernrs[:120]
    try:
        pa = read_table(connection, "PA0001", fields=["PERNR", "BEGDA", "ENDDA", "ABKRS", "BUKRS"],
                        options=opt_and(opt_in("PERNR", sample)), page_size=50_000).rows
        end = f"{params.ano:04d}{params.mes:02d}30"
        abkrs = Counter(r["ABKRS"] for r in pa if r.get("BEGDA", "") <= end <= r.get("ENDDA", "99991231"))
        if abkrs:
            ctx["abkrs"] = abkrs.most_common(1)[0][0]
    except RfcReadError as exc:
        logger.warning("PA0001: %s", exc)

    if ctx["abkrs"]:
        try:
            t549a = read_table(connection, "T549A", fields=["ABKRS", "PERMO"],
                               options=opt_and(opt_eq("ABKRS", ctx["abkrs"]))).rows
            if t549a:
                ctx["permo"] = t549a[0].get("PERMO", "")
        except RfcReadError:
            pass

    diag = describe_table(connection, "T500L")
    avail = {n.upper() for n in diag.field_names()}
    want = [f for f in ("MOLGA", "RELID", "INTCA", "RPLND") if f in avail] or ["MOLGA", "RELID", "INTCA"]
    try:
        for r in read_table(connection, "T500L", fields=want, page_size=5000).rows:
            if r.get("INTCA") == "PT" or r.get("RPLND") == "P":
                ctx["molga"] = r.get("MOLGA", "")
                ctx["relid"] = r.get("RELID", "")
                ctx["land1"] = r.get("INTCA", r.get("RPLND", ""))
                break
    except RfcReadError as exc:
        logger.warning("T500L: %s", exc)
    return ctx


def _pernrs_and_ppoix_for_run(connection: Any, params: AnalysisParams, run: str,
                              account_10: str, account_raw: str,
                              ) -> tuple[list[tuple[str, str]], list[str], list[str], dict[str, dict[str, Decimal]]]:
    """PPDHD→PPDIT→PPDIX→PPOIX para um run: linhas de posting da conta, LINUMs,
    PERNR e BETRG por rubrica. Tudo automático, sem input manual."""
    hdr = read_table(connection, "PPDHD", fields=["DOCNUM", "RUNID", "BUKRS"],
                     options=opt_and(opt_eq("RUNID", run)), page_size=params.page_size).rows
    docs = sorted({r["DOCNUM"] for r in hdr})
    if not docs:
        return [], [], [], {}

    acct = {account_10, account_raw}
    pit_rows: list[dict[str, str]] = []
    for s in range(0, len(docs), 100):
        try:
            pit_rows.extend(read_table(
                connection, "PPDIT", fields=["DOCNUM", "DOCLIN", "BUKRS", "HKONT"],
                options=opt_and(opt_in("DOCNUM", docs[s:s + 100])), page_size=params.page_size).rows)
        except NoData:
            pass
    doc_lines = sorted({
        (r["DOCNUM"], r["DOCLIN"]) for r in pit_rows
        if r.get("BUKRS") == params.empresa
        and (r.get("HKONT") in acct or r.get("HKONT", "").lstrip("0") == account_raw.lstrip("0"))
    })
    if not doc_lines:
        return [], [], [], {}

    dix = read_table(connection, "PPDIX", fields=["RUNID", "LINUM", "DOCNUM", "DOCLIN"],
                     options=opt_and(opt_eq("RUNID", run)), page_size=params.page_size).rows
    dl = set(doc_lines)
    linums = sorted({r["LINUM"] for r in dix if (r["DOCNUM"], r["DOCLIN"]) in dl})
    if not linums:
        return doc_lines, [], [], {}

    poix = read_table(connection, "PPOIX",
                      fields=["RUNID", "PERNR", "TSLIN", "LGART", "KOMOK", "BETRG"],
                      options=opt_and(opt_eq("RUNID", run)), page_size=20_000, max_rows=500_000).rows
    lin = set(linums)
    by_pernr_wt: dict[str, dict[str, Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal("0")))
    pernrs: list[str] = []
    for r in poix:
        if r["TSLIN"] not in lin:
            continue
        pn = r["PERNR"].strip()
        by_pernr_wt[pn][r["LGART"].strip()] += sap_str_to_decimal(r.get("BETRG", "0"))
        if pn not in pernrs:
            pernrs.append(pn)
    return doc_lines, linums, sorted(pernrs), {k: dict(v) for k, v in by_pernr_wt.items()}


def collect_payroll_context(
    connection: Any,
    params: AnalysisParams,
    *,
    runid: str | None = None,
    company: str | None = None,
    account: str | None = None,
    in_period: str | None = None,
) -> PayrollContext:
    """Percorre automaticamente PPDHD→PPDIT→PPDIX→PPOIX→PERNR→HRPY_RGDIR→PA0001.

    Não pede nada ao utilizador; o PERNR é derivado do posting.
    """
    from dataclasses import replace

    run = pad_run(runid or params.primary_run)
    p = params
    if company or account:
        p = replace(params, empresa=company or params.empresa, conta=account or params.conta)
    ctx = PayrollContext(
        run_id=run, company=p.empresa, account=p.conta_10,
        in_period=in_period or f"{p.ano:04d}{p.mes:02d}",
    )

    doc_lines, linums, pernrs, ppoix = _pernrs_and_ppoix_for_run(
        connection, p, run, p.conta_10, p.conta)
    ctx.doc_lines = doc_lines
    ctx.transfer_linums = linums
    ctx.pernrs = pernrs
    ctx.ppoix_by_pernr_wt = {k: {wt: str(v) for wt, v in d.items()} for k, d in ppoix.items()}
    if not pernrs:
        ctx.warn(f"Sem PERNR para run {run} / empresa {p.empresa} / conta {p.conta}.")
        return ctx

    meta = discover_payroll_context(connection, p, pernrs)
    ctx.molga, ctx.abkrs, ctx.permo, ctx.relid = meta["molga"], meta["abkrs"], meta["permo"], meta["relid"]

    try:
        entries = read_rgdir(connection, pernrs, since=rgdir_window_start(ctx.in_period))
        for e in entries:
            ctx.rgdir_by_pernr.setdefault(e.pernr, []).append(e)
    except RfcReadError as exc:
        ctx.warn(f"RGDIR indisponível: {exc}")

    try:
        end = f"{p.ano:04d}{p.mes:02d}30"
        pa_rows: list[dict[str, str]] = []
        for s in range(0, len(pernrs), 100):
            pa_rows.extend(read_table(
                connection, "PA0001",
                fields=["PERNR", "BEGDA", "ENDDA", "ABKRS", "BUKRS", "WERKS", "PERSG", "PERSK"],
                options=opt_and(opt_in("PERNR", pernrs[s:s + 100])), page_size=50_000).rows)
        for r in pa_rows:
            if r.get("BEGDA", "") <= end <= r.get("ENDDA", "99991231"):
                ctx.pa_by_pernr[r["PERNR"]] = r
    except RfcReadError as exc:
        ctx.warn(f"PA0001 indisponível: {exc}")

    ctx.resolved = bool(pernrs and ctx.rgdir_by_pernr)
    logger.info("collect_payroll_context: run=%s PERNR=%s RGDIR=%s PA0001=%s",
                run, len(pernrs), len(ctx.rgdir_by_pernr), len(ctx.pa_by_pernr))
    return ctx


# ===========================================================================
# Catálogo automático de tabelas de resultado de Payroll
# ===========================================================================

def discover_payroll_result_tables(connection: Any) -> list[HrpyTableInfo]:
    """Descobre no DDIC as tabelas transparentes ligadas a resultados de
    Payroll e testa acessibilidade/dados por RFC (sem assumir nomes)."""
    catalog: dict[str, HrpyTableInfo] = {}

    # 1) por nome, via DD02L
    for pat in ("HRPY_%", "HRPADNLP_%", "P2RX_%", "PYD_D_RES%"):
        try:
            rows = read_table(connection, "DD02L", fields=["TABNAME", "TABCLASS"],
                              options=[{"TEXT": f"TABNAME LIKE '{pat}'"}], page_size=20_000).rows
        except RfcReadError:
            continue
        for r in rows:
            if r["TABCLASS"] in {"TRANSP", "VIEW"}:
                catalog.setdefault(r["TABNAME"], HrpyTableInfo(
                    table=r["TABNAME"], table_class=r["TABCLASS"], exists=True))

    # 2) candidatas conhecidas + textos
    for name in PAYROLL_RESULT_TABLE_CANDIDATES:
        catalog.setdefault(name, HrpyTableInfo(table=name))
    try:
        texts = read_table(connection, "DD02T", fields=["TABNAME", "DDLANGUAGE", "DDTEXT"],
                           options=opt_and(opt_in("TABNAME", sorted(catalog))), page_size=20_000).rows
        best: dict[str, str] = {}
        for r in texts:
            if r["DDLANGUAGE"] in {"E", "P"} and (r["TABNAME"] not in best or r["DDLANGUAGE"] == "E"):
                best[r["TABNAME"]] = r["DDTEXT"]
        for name, txt in best.items():
            catalog[name].description = txt
    except RfcReadError:
        pass

    # 3) nº de campos via DD03L (funciona para qualquer tabela) + acessibilidade
    #    e dados só para as tabelas que estão na whitelist read-only.
    wl = {t.upper() for t in READ_ONLY_TABLE_WHITELIST}
    try:
        f_rows = read_table(connection, "DD03L", fields=["TABNAME", "FIELDNAME"],
                            options=opt_and(opt_in("TABNAME", sorted(catalog))), page_size=200_000).rows
        fcount: dict[str, int] = Counter(
            r["TABNAME"] for r in f_rows
            if r.get("FIELDNAME") and not r["FIELDNAME"].startswith("."))
    except RfcReadError:
        fcount = {}

    out: list[HrpyTableInfo] = []
    for name in sorted(catalog):
        info = catalog[name]
        info.field_count = int(fcount.get(name, 0))
        info.exists = info.field_count > 0 or info.exists
        if name.upper() not in wl:
            info.note = "não testada quanto a dados (fora da whitelist read-only)"
            out.append(info)
            continue
        try:
            res = read_table(connection, name, max_rows=1)
            info.accessible = True
            info.populated = len(res.rows) > 0
            info.exists = True
        except RfcReadError as exc:
            info.accessible = False
            info.exists = exc.kind != "TABLE_NOT_AVAILABLE"
            info.note = exc.kind
        out.append(info)
    return out


# ===========================================================================
# Tentativa de leitura da RT (cluster)
# ===========================================================================

def attempt_read_rt(connection: Any, params: AnalysisParams, pernr: str, seqnr: str,
                    relid: str) -> RtReadAttempt:
    att = RtReadAttempt(function=RT_READ_FUNCTIONS[0], attempted=True)
    for cid in [relid, "99", "RP"]:
        if not cid:
            continue
        try:
            res = safe_rfc_call(
                connection, "PYXX_READ_PAYROLL_RESULT",
                CLUSTERID=cid, EMPLOYEENUMBER=pernr, SEQUENCENUMBER=seqnr,
                READ_ONLY_INTERNATIONAL="X",
            )
            pr = res.get("PAYROLL_RESULT")
            att.ok = True
            att.reason = f"OK (CLUSTERID={cid})"
            if isinstance(pr, dict):
                inter = pr.get("INTER") or pr.get("INTERNATIONAL") or {}
                rt = inter.get("RT") if isinstance(inter, dict) else None
                att.detail = f"chaves={list(pr.keys())}"
                if isinstance(rt, list):
                    att.sample = rt[:20]
            return att
        except Exception as exc:  # noqa: BLE001
            raw = str(exc)
            is_da300 = ("DA" in raw and "300" in raw) or "nametab" in raw.lower()
            if is_da300:
                att.reason = f"{type(exc).__name__} DA300 «No active nametab» (CLUSTERID={cid})"
                att.detail = (
                    "O IMPORT do cluster PCL2 dentro deste FM não funciona em "
                    "contexto RFC stateless (nametab do resultado país não activa). "
                    "HR_GET_PAYROLL_RESULTS não é RFC-enabled; as tabelas "
                    "transparentes P2RX_* existem mas estão vazias. É preciso um "
                    "wrapper Z read-only no SAP — não há alternativa de escrita."
                )
            else:
                att.reason = f"{type(exc).__name__}: {raw[:200]}"
    return att


# ===========================================================================
# Orquestração
# ===========================================================================

def _target_pernrs_from_link(link_report: Any) -> list[str]:
    seen: list[str] = []
    for r in getattr(link_report, "link_sample", []):
        pn = str(r.get("PERNR", "")).strip()
        if pn and pn not in seen:
            seen.append(pn)
    return sorted(seen)


_REF_VIEW_WTS = ("/558", "/559", "/561", "/563", "0029")


def analyse_cluster(
    connection: Any,
    params: AnalysisParams,
    payroll_report: Any,
    link_report: Any,
    *,
    build_timelines_for: int = 400,
    try_rt: bool = True,
) -> PayrollClusterReport:
    rep = PayrollClusterReport(run_id=pad_run(params.primary_run), company=params.empresa)

    if not (link_report and getattr(link_report, "resolved", False) and link_report.link_sample):
        rep.warn("Fase 2 sem resultado — Fase 3 (cluster) não pode isolar os PERNR.")
        return rep

    pernrs = _target_pernrs_from_link(link_report)
    rep.pernr_count = len(pernrs)
    inper = rep.period

    ctx = discover_payroll_context(connection, params, pernrs)
    rep.molga, rep.abkrs, rep.permo, rep.relid = ctx["molga"], ctx["abkrs"], ctx["permo"], ctx["relid"]

    # --- catálogo automático de tabelas de resultado ---
    try:
        rep.hrpy_catalog = discover_payroll_result_tables(connection)
    except RfcReadError as exc:
        rep.warn(f"Catálogo HRPY/P2RX indisponível: {exc}")

    # --- RGDIR (janela relevante para a cadeia de retro) ---
    try:
        rep.rgdir_entries = read_rgdir(connection, pernrs, since=rgdir_window_start(inper))
    except RfcReadError as exc:
        rep.warn(f"RGDIR indisponível: {exc}")
        return rep

    by_pernr: dict[str, list[PayrollDirectoryEntry]] = defaultdict(list)
    for e in rep.rgdir_entries:
        by_pernr[e.pernr].append(e)

    # entradas que o run transferiu = TODAS as de INPER == período do posting
    # (a SRTZA de hoje pode já ser P/O por causa de runs posteriores).
    in_run = [e for e in rep.rgdir_entries if e.inper == inper]
    rep.rgdir_for_inper = sorted(in_run, key=lambda e: (e.pernr, e.fpper, e.seqnr))
    rep.fpper_distribution = dict(Counter(e.fpper for e in in_run))
    rep.srtza_distribution = dict(Counter(e.srtza for e in in_run))
    rep.classification_distribution = dict(Counter(e.classify() for e in in_run))
    rep.offcycle_count = sum(1 for e in in_run if e.is_offcycle)
    rep.void_count = sum(1 for e in in_run if e.is_void)

    cur_set, retro_set, retro_corr_set = set(), set(), set()
    retro_months: dict[str, int] = {}
    for e in in_run:
        if e.fpper == inper:
            cur_set.add(e.pernr)
        elif e.fpper < inper:
            retro_set.add(e.pernr)
            retro_months[e.pernr] = retro_months.get(e.pernr, 0) + 1
            if e.months_late >= 2:
                retro_corr_set.add(e.pernr)
    rep.current_pernr = sorted(cur_set)
    rep.retro_pernr = sorted(retro_set)
    rep.retro_months_hist = dict(Counter(retro_months.get(pn, 0) for pn in pernrs))
    unclassified = [pn for pn in pernrs if pn not in cur_set and pn not in retro_set]
    rep.residual_notes["retro_lag_vs_corr"] = (
        f"{len(retro_set) - len(retro_corr_set)}/{len(retro_set)} PERNR com retro só "
        f"de rotina (desfasamento sistemático de 1 mês); {len(retro_corr_set)} com "
        f"correcção real (>=2 meses). Esta folha finaliza cada período no run seguinte."
    )

    # --- PPOIX de referência por PERNR (da Fase 2) ---
    ref_types = set(params.wage_types_referencia)
    ppoix_ref: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    ppoix_wt: dict[str, dict[str, Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal("0")))
    for r in link_report.link_sample:
        pn = str(r.get("PERNR", "")).strip()
        amt = sap_str_to_decimal(r.get("BETRG", "0"))
        ppoix_wt[pn][r.get("LGART", "")] += amt
        if r.get("LGART") in ref_types:
            ppoix_ref[pn] += amt
    rep.ppoix_ref_by_pernr = {k: str(v) for k, v in ppoix_ref.items()}
    rep.ppoix_ref_total = sum(ppoix_ref.values(), Decimal("0"))
    rep.ppoix_ref_retro_total = sum((v for k, v in ppoix_ref.items() if k in retro_set and k not in cur_set), Decimal("0"))
    rep.ppoix_ref_current_total = sum((v for k, v in ppoix_ref.items() if k in cur_set and k not in retro_set), Decimal("0"))
    mixed = sum((v for k, v in ppoix_ref.items() if k in cur_set and k in retro_set), Decimal("0"))
    rep.ppoix_ref_unclassified_total = sum((v for k, v in ppoix_ref.items() if k in set(unclassified)), Decimal("0"))
    if mixed:
        rep.residual_notes["ppoix_ref_mixed_current_and_retro"] = str(mixed)
    if unclassified:
        rep.warn(
            f"{len(unclassified)} PERNR do PPOIX sem entrada RGDIR com INPER={inper} "
            f"(PPOIX /558+/559 desses = {rep.ppoix_ref_unclassified_total})."
        )

    # --- timelines + pares original→recalculado ---
    for pn in pernrs[:build_timelines_for]:
        tl = build_timeline(pn, by_pernr.get(pn, []))
        rep.timelines.append(tl)
        for pair in tl.pairs:
            if pair.status in {"RESULT_RECALCULATED", "RESULT_VOIDED", "RESULT_CURRENT_ONLY"} or \
               (pair.fpper and pair.fpper <= inper and any(e.inper == inper for e in pair.entries)):
                d = pair.as_dict()
                d["contributes_to_run"] = any(e.inper == inper for e in pair.entries)
                rep.recalc_pairs.append(d)

    # --- vista única PPOIX x RGDIR por PERNR ---
    for pn in pernrs:
        run_entries = [e for e in by_pernr.get(pn, []) if e.inper == inper]
        fppers = sorted({e.fpper for e in run_entries})
        cur_e = next((e for e in run_entries if e.fpper == inper), None)
        wt = ppoix_wt.get(pn, {})
        row = {
            "pernr": pn,
            **{f"ppoix_{w.strip('/')}": str(wt.get(w, Decimal("0"))) for w in _REF_VIEW_WTS},
            "ppoix_ref": str(ppoix_ref.get(pn, Decimal("0"))),
            "rgdir_run_fppers": fppers,
            "rgdir_run_seqnrs": [e.seqnr for e in sorted(run_entries, key=lambda e: e.seqnr)],
            "retro_months": len([f for f in fppers if f < inper]),
            "has_current": cur_e is not None,
            "current_seqnr": cur_e.seqnr if cur_e else "",
            "classes": sorted({e.classify() for e in run_entries}),
        }
        rep.ppoix_rgdir_view.append(row)
        rep.per_pernr_diag.append({
            "pernr": pn, "ppoix_ref": row["ppoix_ref"],
            "rgdir_run_fppers": fppers, "retro_months": row["retro_months"],
            "retro": pn in retro_set, "current": pn in cur_set, "rt_value": None,
        })

    # --- tentativa RT (saltável: o modo manual-rt-request não toca no cluster) ---
    if try_rt and rep.rgdir_for_inper:
        e0 = rep.rgdir_for_inper[0]
        rep.rt_attempt = attempt_read_rt(connection, params, e0.pernr, e0.seqnr, rep.relid)
        if not rep.rt_attempt.ok:
            rep.warn(f"RT não legível por RFC read-only ({rep.rt_attempt.function}: "
                     f"{rep.rt_attempt.reason}).")
    elif not try_rt:
        rep.rt_attempt.reason = "não tentada (modo manual-rt-request)"

    # --- comparação run 1299 ---
    rep.run_1299_comparison = _compare_run(connection, params, payroll_report, link_report, "0000001299")

    # --- notas de resíduo ---
    ppoix_total = abs(getattr(link_report, "ppoix_total", Decimal("0")))
    posting = abs(getattr(link_report, "posting_line_amount", Decimal("0")))
    ref_ppoix = abs(getattr(link_report, "reference_total", Decimal("0")))
    n_recalc = sum(1 for p in rep.recalc_pairs if p["status"] == "RESULT_RECALCULATED"
                   and p["contributes_to_run"])
    rep.residual_notes.update({
        "ppoix_vs_ppdit": str((ppoix_total - posting).quantize(Decimal("0.01"))),
        "ppoix_ref_vs_rh": str((ref_ppoix - params.valor_rh_referencia).quantize(Decimal("0.01"))),
        "retro_share": (
            f"{len(retro_set)}/{len(pernrs)} PERNR com componente retro no run "
            f"(FOR-period < {inper}); {len(cur_set)}/{len(pernrs)} com componente do "
            f"próprio período. {n_recalc} pares FOR-period recalculados alimentam este run."
        ),
        "explicacao": (
            f"Esta folha corre com desfasamento sistemático de 1 mês: o resultado "
            f"definitivo de cada FOR-period é produzido no run seguinte. O run "
            f"{rep.run_id} (IN-period {inper}) transferiu, por PERNR, o resultado do "
            f"próprio período MAIS o recálculo de rotina de {inper[:4]}-05 (e, para "
            f"poucos casos, correcções mais antigas). A referência RH (724.046,64) é "
            f"outro recorte. Os resíduos 427,74 (PPOIX vs RH) e 265,65 (PPOIX vs "
            f"PPDIT) só se atribuem ao cêntimo por PERNR com a RT (definitivo vs "
            f"provisório do período) — cluster PCL2, não legível por RFC aqui."
        ),
    })
    rep.resolved = True
    logger.info(
        "Fase 3: %s PERNR (%s c/ retro, %s c/ período; classes=%s); RT_ok=%s",
        len(pernrs), len(retro_set), len(cur_set), rep.classification_distribution, rep.rt_attempt.ok,
    )
    return rep


def _compare_run(connection: Any, params: AnalysisParams, payroll_report: Any,
                 link_report: Any, other_run: str) -> dict[str, Any]:
    from dataclasses import replace

    from .payroll_wagetypes import link_wage_types_to_posting_line

    try:
        other = link_wage_types_to_posting_line(
            connection, replace(params, primary_run=other_run), payroll_report, run_id=other_run,
        )
    except Exception as exc:  # noqa: BLE001
        return {"run": other_run, "ok": False, "error": str(exc)}
    if not other.resolved:
        return {"run": other_run, "ok": False, "error": "; ".join(other.warnings)}

    base_pernr = {r["PERNR"] for r in link_report.link_sample}
    other_pernr = {r["PERNR"] for r in other.link_sample}
    same_lgart = link_report.by_wage_type == other.by_wage_type
    same_amount = (link_report.posting_line_amount == other.posting_line_amount
                   and link_report.ppoix_total == other.ppoix_total)
    classification = ("REPEAT_POSTING / RERUN"
                      if (base_pernr == other_pernr and same_lgart) else "DIFERENTE")
    return {
        "run": other_run, "ok": True,
        "posting_line_amount": str(other.posting_line_amount),
        "ppoix_total": str(other.ppoix_total),
        "pernr_count": len(other_pernr),
        "same_pernr_set": base_pernr == other_pernr,
        "same_lgart_totals": same_lgart,
        "same_amounts": same_amount,
        "by_wage_type": other.by_wage_type,
        "classification": classification,
    }
