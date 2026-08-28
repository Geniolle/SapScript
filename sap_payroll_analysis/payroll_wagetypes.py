"""Origem por rubrica salarial (wage type) — PPOIX.

PPOIX liga o ciclo de lançamentos (RUNID) à rubrica salarial (LGART),
ao montante (BETRG) e à conta simbólica (KOMOK). É o ponto de partida para
explicar quais rubricas compõem o valor lançado na conta 23120000 e quais
justificam a diferença face a /558 + /559.

Esta fase produz a estrutura e os totais por rubrica; o passo seguinte
(documentado em `report.NEXT_STEPS`) é filtrar por empresa e cruzar KOMOK
com a determinação de contas (T52EK / T030) para isolar a conta 23120000.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from .config import AnalysisParams, pad_run
from .ddic import describe_table, guess_fields
from .models import TableDiag
from .sap_reader import (
    NoData,
    RfcReadError,
    opt_and,
    opt_eq,
    opt_in,
    read_table,
    sap_str_to_decimal,
)

logger = logging.getLogger(__name__)

_PREF = {"run": "RUNID", "wage_type": "LGART", "amount": "BETRG", "currency": "WAERS",
         "symbolic": "KOMOK", "pernr": "PERNR", "seq": "SEQNO", "neg": "NEG_POSTNG",
         "post": "POSTNUM", "actsign": "ACTSIGN"}


@dataclass
class WageTypeReport:
    diag: TableDiag | None = None
    resolved_fields: dict[str, str | None] = field(default_factory=dict)
    by_wage_type: dict[str, dict[str, Any]] = field(default_factory=dict)
    by_wage_type_symbolic: dict[str, dict[str, Any]] = field(default_factory=dict)
    reference_total: Decimal = Decimal("0")
    reference_rows: int = 0
    symbolic_accounts: list[str] = field(default_factory=list)
    truncated: bool = False
    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)


def probe_reference_wage_types(
    connection: Any, params: AnalysisParams, *, all_wage_types: bool = False, max_rows: int = 300_000
) -> WageTypeReport:
    """Lê PPOIX para os runs em análise e agrega BETRG por rubrica.

    `all_wage_types=False` (default) restringe a `WAGE_TYPES_REFERENCIA`
    (/558, /559) — leitura pequena e rápida.
    """
    report = WageTypeReport()
    report.diag = describe_table(connection, "PPOIX")
    if not (report.diag.exists and report.diag.authorized):
        report.warn(f"PPOIX indisponível/sem autorização: {report.diag.note}")
        return report

    names = {n.upper() for n in report.diag.field_names()}
    guesses = guess_fields(report.diag)
    f: dict[str, str | None] = {}
    for key, pref in _PREF.items():
        if pref in names:
            f[key] = pref
        else:
            concept = {"run": "posting_run", "wage_type": "wage_type", "amount": "valor",
                       "currency": "moeda", "symbolic": "conta_simbolica", "pernr": "pernr"}.get(key)
            g = guesses.get(concept) if concept else None
            f[key] = g.chosen if g and g.chosen else None
    report.resolved_fields = dict(f)

    if not (f["run"] and f["wage_type"] and f["amount"]):
        report.warn("PPOIX: não foi possível resolver RUNID/LGART/BETRG.")
        return report

    runs = [pad_run(r) for r in params.posting_runs]
    groups = [opt_in(f["run"], runs)]
    if not all_wage_types:
        groups.append(opt_in(f["wage_type"], list(params.wage_types_referencia)))

    fields = [x for x in (f["run"], f["wage_type"], f["symbolic"], f["currency"],
                          f["neg"], f["amount"]) if x]
    try:
        res = read_table(connection, "PPOIX", fields=fields,
                         options=opt_and(*groups), page_size=20_000, max_rows=max_rows)
    except NoData:
        report.resolved = True
        return report
    except RfcReadError as exc:
        report.warn(f"Falha a ler PPOIX: {exc}")
        return report

    report.truncated = res.truncated
    wt: dict[str, list[Any]] = defaultdict(lambda: [0, Decimal("0")])
    wtk: dict[tuple[str, str], list[Any]] = defaultdict(lambda: [0, Decimal("0")])
    symbolic: set[str] = set()

    for r in res.rows:
        amount = sap_str_to_decimal(r.get(f["amount"], "0"))
        if f["neg"] and r.get(f["neg"], "").strip().upper() == "X":
            amount = -amount
        lgart = r.get(f["wage_type"], "").strip()
        komok = r.get(f["symbolic"], "").strip() if f["symbolic"] else ""
        wt[lgart][0] += 1
        wt[lgart][1] += amount
        wtk[(lgart, komok)][0] += 1
        wtk[(lgart, komok)][1] += amount
        if komok:
            symbolic.add(komok)

    report.by_wage_type = {
        k: {"rows": v[0], "amount": str(v[1])} for k, v in sorted(wt.items())
    }
    report.by_wage_type_symbolic = {
        f"{k[0]}|{k[1]}": {"rows": v[0], "amount": str(v[1])} for k, v in sorted(wtk.items())
    }
    report.symbolic_accounts = sorted(symbolic)
    ref_types = set(params.wage_types_referencia)
    report.reference_total = sum(
        (Decimal(v["amount"]) for k, v in report.by_wage_type.items() if k in ref_types),
        Decimal("0"),
    )
    report.reference_rows = sum(
        v["rows"] for k, v in report.by_wage_type.items() if k in ref_types
    )
    report.resolved = True
    logger.info("PPOIX: /558+/559 total=%s (%s linhas), contas simbólicas=%s",
                report.reference_total, report.reference_rows, report.symbolic_accounts)
    return report


# ===========================================================================
# FASE 2 — ligar as rubricas PPOIX à linha de posting contabilística
#
#   PPOIX.TSLIN  ==  PPDIX.LINUM
#   PPDIX.(DOCNUM, DOCLIN)  ==  PPDIT.(DOCNUM, DOCLIN)   <- linha da conta
#
# (POSTNUM/PPOPX é um split anterior à transferência; PPOPX.TSLIN vem a 0.)
# ===========================================================================

_PPDIX_PREF = {"run": "RUNID", "linum": "LINUM", "doc": "DOCNUM", "line": "DOCLIN"}


@dataclass
class WageTypeLinkReport:
    """Composição por rubrica de UMA linha de posting (um run, uma conta)."""

    run_id: str = ""
    company: str = ""
    account: str = ""
    posting_doc_lines: list[tuple[str, str]] = field(default_factory=list)
    posting_line_amount: Decimal = Decimal("0")   # PPDIT, com sinal
    transfer_linums: list[str] = field(default_factory=list)

    ppoix_rows: int = 0
    ppoix_total: Decimal = Decimal("0")           # soma de todas as rubricas ligadas
    by_wage_type: dict[str, dict[str, Any]] = field(default_factory=dict)
    by_wage_type_komok: dict[str, dict[str, Any]] = field(default_factory=dict)
    komok_set: list[str] = field(default_factory=list)

    reference_total: Decimal = Decimal("0")       # /558 + /559
    reference_by_type: dict[str, str] = field(default_factory=dict)
    other_total: Decimal = Decimal("0")           # tudo o resto
    other_by_type: dict[str, str] = field(default_factory=dict)

    residual_vs_posting: Decimal = Decimal("0")   # ppoix_total - posting_line_amount
    link_sample: list[dict[str, str]] = field(default_factory=list)
    account_determination: dict[str, Any] = field(default_factory=dict)

    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)

    def as_dict(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "company": self.company,
            "account": self.account,
            "posting_doc_lines": [f"{d}/{l}" for d, l in self.posting_doc_lines],
            "posting_line_amount": str(self.posting_line_amount),
            "transfer_linums": self.transfer_linums,
            "ppoix_rows": self.ppoix_rows,
            "ppoix_total": str(self.ppoix_total),
            "by_wage_type": self.by_wage_type,
            "by_wage_type_komok": self.by_wage_type_komok,
            "komok_set": self.komok_set,
            "reference_total": str(self.reference_total),
            "reference_by_type": self.reference_by_type,
            "other_total": str(self.other_total),
            "other_by_type": self.other_by_type,
            "residual_vs_posting": str(self.residual_vs_posting),
            "link_sample": self.link_sample[:300],
            "link_sample_truncated": len(self.link_sample) > 300,
            "account_determination": self.account_determination,
            "resolved": self.resolved,
            "warnings": self.warnings,
        }


def link_wage_types_to_posting_line(
    connection: Any,
    params: AnalysisParams,
    payroll_report: Any,
    run_id: str | None = None,
    *,
    sample_size: int = 20_000,
) -> WageTypeLinkReport:
    """Constrói a composição por rubrica da linha da conta `params.conta`
    para o run indicado (por omissão `params.primary_run`), empresa `params.empresa`.
    """
    run = pad_run(run_id or params.primary_run)
    rep = WageTypeLinkReport(run_id=run, company=params.empresa, account=params.conta_10)

    # 1) linhas de posting alvo (PPDIT) já isoladas na fase 1
    target_items = [
        it for it in getattr(payroll_report, "items", [])
        if it.run_id == run
        and (not it.company or it.company == params.empresa)
        and (it.account == params.conta_10 or it.account.lstrip("0") == params.conta.lstrip("0"))
    ]
    if not target_items:
        rep.warn(
            f"Sem linha de posting para run {run} / empresa {params.empresa} / "
            f"conta {params.conta}. Fase 2 não aplicável a este run."
        )
        return rep

    rep.posting_doc_lines = sorted({(it.doc_number, it.line) for it in target_items})
    rep.posting_line_amount = sum((it.signed_amount for it in target_items), Decimal("0"))
    ktosl = ""
    for it in target_items:
        ktosl = (it.raw or {}).get("KTOSL", "").strip() or ktosl
    ktosl = ktosl or params.hr_posting_key

    # 2) PPDIX: LINUM -> (DOCNUM, DOCLIN)  para o run
    dix = describe_table(connection, "PPDIX")
    if not (dix.exists and dix.authorized):
        rep.warn(f"PPDIX indisponível/sem autorização: {dix.note}")
        return rep
    dix_names = {n.upper() for n in dix.field_names()}
    dg = guess_fields(dix)
    df = {
        k: (pref if pref in dix_names else _guess_or_none(dg, _PPDIX_CONCEPT.get(k)))
        for k, pref in _PPDIX_PREF.items()
    }
    if not (df["run"] and df["linum"] and df["doc"] and df["line"]):
        rep.warn(f"PPDIX: não resolvi RUNID/LINUM/DOCNUM/DOCLIN (campos {dix.field_names()}).")
        return rep

    try:
        dix_rows = read_table(
            connection, "PPDIX",
            fields=[df["run"], df["linum"], df["doc"], df["line"]],
            options=opt_and(opt_eq(df["run"], run)), page_size=params.page_size,
        ).rows
    except (NoData, RfcReadError) as exc:
        rep.warn(f"Falha a ler PPDIX: {exc}")
        return rep

    linum_to_docline: dict[str, tuple[str, str]] = {}
    target_keys = set(rep.posting_doc_lines)
    target_linums: set[str] = set()
    for r in dix_rows:
        ln = r.get(df["linum"], "").strip()
        key = (r.get(df["doc"], "").strip(), r.get(df["line"], "").strip())
        linum_to_docline[ln] = key
        if key in target_keys:
            target_linums.add(ln)
    rep.transfer_linums = sorted(target_linums)
    if not target_linums:
        rep.warn("PPDIX não tem nenhuma linha de transferência (LINUM) a apontar "
                 "para a(s) linha(s) de posting alvo.")
        return rep

    # 3) PPOIX do run -> filtrar TSLIN in target_linums
    poix = describe_table(connection, "PPOIX")
    pn = {n.upper() for n in poix.field_names()}
    pg = guess_fields(poix)
    pf = {
        "run": _first(pn, ["RUNID"]) or _guess_or_none(pg, "posting_run"),
        "pernr": _first(pn, ["PERNR"]) or _guess_or_none(pg, "pernr"),
        "postnum": _first(pn, ["POSTNUM"]),
        "rtline": _first(pn, ["RTLINE"]),
        "tslin": _first(pn, ["TSLIN"]),
        "lgart": _first(pn, ["LGART"]) or _guess_or_none(pg, "wage_type"),
        "komok": _first(pn, ["KOMOK"]) or _guess_or_none(pg, "conta_simbolica"),
        "amount": _first(pn, ["BETRG"]) or _guess_or_none(pg, "valor"),
        "actsign": _first(pn, ["ACTSIGN"]),
        "neg": _first(pn, ["NEG_POSTNG"]),
    }
    if not (pf["run"] and pf["tslin"] and pf["lgart"] and pf["amount"]):
        rep.warn(f"PPOIX: não resolvi RUNID/TSLIN/LGART/BETRG (campos {poix.field_names()}).")
        return rep

    fields = [v for v in (pf["run"], pf["pernr"], pf["postnum"], pf["rtline"], pf["tslin"],
                          pf["lgart"], pf["komok"], pf["amount"], pf["actsign"], pf["neg"]) if v]
    try:
        prows = read_table(
            connection, "PPOIX", fields=fields,
            options=opt_and(opt_eq(pf["run"], run)), page_size=20_000, max_rows=500_000,
        ).rows
    except (NoData, RfcReadError) as exc:
        rep.warn(f"Falha a ler PPOIX: {exc}")
        return rep

    matched = [r for r in prows if r.get(pf["tslin"], "").strip() in target_linums]
    rep.ppoix_rows = len(matched)

    by_lg: dict[str, list[Any]] = defaultdict(lambda: [0, Decimal("0")])
    by_lgk: dict[tuple[str, str], list[Any]] = defaultdict(lambda: [0, Decimal("0")])
    komok: set[str] = set()
    for r in matched:
        amt = sap_str_to_decimal(r.get(pf["amount"], "0"))
        if pf["neg"] and r.get(pf["neg"], "").strip().upper() == "X":
            amt = -amt
        lg = r.get(pf["lgart"], "").strip()
        km = r.get(pf["komok"], "").strip() if pf["komok"] else ""
        by_lg[lg][0] += 1
        by_lg[lg][1] += amt
        by_lgk[(lg, km)][0] += 1
        by_lgk[(lg, km)][1] += amt
        if km:
            komok.add(km)

    rep.by_wage_type = {k: {"rows": v[0], "amount": str(v[1])} for k, v in sorted(by_lg.items())}
    rep.by_wage_type_komok = {
        f"{k[0]}|{k[1]}": {"rows": v[0], "amount": str(v[1])} for k, v in sorted(by_lgk.items())
    }
    rep.komok_set = sorted(komok)
    rep.ppoix_total = sum((Decimal(v["amount"]) for v in rep.by_wage_type.values()), Decimal("0"))

    ref = set(params.wage_types_referencia)
    rep.reference_by_type = {k: v["amount"] for k, v in rep.by_wage_type.items() if k in ref}
    rep.reference_total = sum((Decimal(a) for a in rep.reference_by_type.values()), Decimal("0"))
    rep.other_by_type = {k: v["amount"] for k, v in rep.by_wage_type.items() if k not in ref}
    rep.other_total = sum((Decimal(a) for a in rep.other_by_type.values()), Decimal("0"))
    rep.residual_vs_posting = rep.ppoix_total - rep.posting_line_amount

    # amostra "PPOIX LINK ANALYSIS"
    for r in matched[:sample_size]:
        tsl = r.get(pf["tslin"], "").strip()
        dn, dl = linum_to_docline.get(tsl, ("", ""))
        rep.link_sample.append({
            "PERNR": r.get(pf["pernr"], "").strip() if pf["pernr"] else "",
            "POSTNUM": r.get(pf["postnum"], "").strip() if pf["postnum"] else "",
            "RTLINE": r.get(pf["rtline"], "").strip() if pf["rtline"] else "",
            "LGART": r.get(pf["lgart"], "").strip(),
            "KOMOK": r.get(pf["komok"], "").strip() if pf["komok"] else "",
            "BETRG": r.get(pf["amount"], "").strip(),
            "TSLIN": tsl,
            "DOCNUM": dn,
            "DOCLIN": dl,
        })

    # 4) determinação de contas (T52EL / T52EK / T030)
    rep.account_determination = resolve_account_determination(
        connection, params, symkos=rep.komok_set,
        wage_types=list(rep.by_wage_type), ktosl=ktosl,
    )

    rep.resolved = True
    logger.info(
        "Fase 2 run %s: %s linhas PPOIX, total %s, /558+/559 %s, outras %s, resíduo vs posting %s",
        run, rep.ppoix_rows, rep.ppoix_total, rep.reference_total, rep.other_total,
        rep.residual_vs_posting,
    )
    return rep


_PPDIX_CONCEPT = {"run": "posting_run", "doc": "documento", "line": "item"}


def resolve_account_determination(
    connection: Any,
    params: AnalysisParams,
    *,
    symkos: list[str],
    wage_types: list[str],
    ktosl: str,
) -> dict[str, Any]:
    """Cruza a determinação de contas do Payroll (só leitura).

    * T52EL  rubrica (LGART) -> conta simbólica (SYMKO) + SIGN
    * T52EK  atributos da conta simbólica (KOART)
    * T030   KTOPL + KTOSL + BWMOD(=rubrica ou conta simbólica) -> KONTS/KONTH
    """
    out: dict[str, Any] = {"ktosl": ktosl, "symkos": symkos, "t52el": [], "t52ek": [],
                           "t030": [], "wage_types_to_target": [], "target_account": params.conta_10,
                           "conclusion": ""}
    target = {params.conta_10, params.conta.lstrip("0")}
    bwmods = sorted(set(wage_types) | set(symkos))

    # T52EL
    try:
        rows = read_table(connection, "T52EL",
                          fields=["MOLGA", "LGART", "ENDDA", "SIGN", "SYMKO", "SPPRC"],
                          options=opt_and(opt_in("SYMKO", symkos)) if symkos else None,
                          page_size=50_000).rows
        out["t52el"] = [
            {"MOLGA": r.get("MOLGA", ""), "LGART": r.get("LGART", ""), "SIGN": r.get("SIGN", ""),
             "SYMKO": r.get("SYMKO", ""), "ENDDA": r.get("ENDDA", "")}
            for r in rows if r.get("LGART") in set(wage_types) or not wage_types
        ]
    except (NoData, RfcReadError) as exc:
        out["t52el_error"] = str(exc)

    # T52EK
    try:
        rows = read_table(connection, "T52EK",
                          fields=["SYMKO", "KOART", "U_MOMAG", "NEG_POSTNG"],
                          options=opt_and(opt_in("SYMKO", symkos)) if symkos else None,
                          page_size=20_000).rows
        out["t52ek"] = [dict(r) for r in rows]
    except (NoData, RfcReadError) as exc:
        out["t52ek_error"] = str(exc)

    # T030 (determinação FI da conta): KTOSL == ktosl e BWMOD in {rubricas, contas simbólicas}
    try:
        rows = read_table(connection, "T030",
                          fields=["KTOPL", "KTOSL", "BWMOD", "KOMOK", "BKLAS", "KONTS", "KONTH"],
                          options=opt_and(opt_eq("KTOSL", ktosl)), page_size=80_000).rows
        picked = [r for r in rows if r.get("BWMOD", "").strip() in bwmods]
        out["t030"] = [
            {"KTOPL": r.get("KTOPL", ""), "BWMOD": r.get("BWMOD", ""), "KOMOK": r.get("KOMOK", ""),
             "KONTS": r.get("KONTS", ""), "KONTH": r.get("KONTH", "")}
            for r in picked
        ]
        for r in picked:
            acc_s = (r.get("KONTS", "") or "").lstrip("0")
            acc_h = (r.get("KONTH", "") or "").lstrip("0")
            if params.conta.lstrip("0") in {acc_s, acc_h}:
                bw = r.get("BWMOD", "").strip()
                if bw and bw not in out["wage_types_to_target"]:
                    out["wage_types_to_target"].append(bw)
    except (NoData, RfcReadError) as exc:
        out["t030_error"] = str(exc)

    wtt = sorted(out["wage_types_to_target"])
    if wtt:
        out["conclusion"] = (
            f"Confirmado: conta simbólica {', '.join(symkos) or '(n/d)'} (KOART "
            + ",".join(sorted({r.get('KOART', '') for r in out.get('t52ek', [])})) + "); "
            f"as rubricas {', '.join(wtt)} têm entrada em T030 "
            f"(KTOPL/{ktosl}/BWMOD) para a conta {params.conta}."
        )
    elif symkos:
        out["conclusion"] = (
            f"Conta simbólica {', '.join(symkos)} identificada em T52EL/T52EK, mas não "
            f"encontrei entrada T030 explícita KTOSL={ktosl}/BWMOD -> {params.conta}. "
            f"A ligação está evidenciada empiricamente (100% das linhas PPOIX da linha "
            f"de posting têm KOMOK nessas contas simbólicas)."
        )
    else:
        out["conclusion"] = "Sem contas simbólicas nas linhas PPOIX ligadas."
    return out


def _first(available: set[str], candidates: list[str]) -> str:
    for c in candidates:
        if c.upper() in available:
            return c
    return ""


def _guess_or_none(guesses: dict[str, Any], concept: str | None) -> str | None:
    if not concept:
        return None
    g = guesses.get(concept)
    return g.chosen if g and getattr(g, "chosen", None) else None
