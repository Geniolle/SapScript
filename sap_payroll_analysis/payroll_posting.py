"""Análise do posting do Payroll (HR posting, ECC) para uma conta do Razão.

Modelo real das tabelas (descoberto por DDIC neste sistema):

    PEVST  RUNID  ........................ registo/estado do ciclo de lançamentos
    PPDHD  RUNID -> DOCNUM ............... cabeçalho do documento de posting RH
    PPDIX  RUNID <-> DOCNUM/DOCLIN ....... índice (bridge alternativa)
    PPDIT  DOCNUM/DOCLIN .............. itens: HKONT, BUKRS, WRBTR, WAERS, KTOSL
    PPOIX  RUNID, PERNR, LGART, KOMOK, BETRG ... origem por rubrica salarial

Notas confirmadas neste sistema:
* PPDIT **não tem** campo de posting run — a ligação faz-se por DOCNUM via PPDHD.
* PPDHD.GJAHR/MONAT vêm a 0000/00; o período está em PPDHD.BUDAT.
* PPDIT.WRBTR já traz sinal (ex.: "727258.35-"). `NEG_POSTNG='X'` inverte.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from .config import AnalysisParams, pad_run
from .ddic import describe_table, guess_fields
from .models import FieldGuess, PostingItem, TableDiag
from .sap_reader import (
    NoData,
    RfcReadError,
    opt_and,
    opt_in,
    read_table,
    sap_str_to_decimal,
)

logger = logging.getLogger(__name__)

PAYROLL_TABLES: tuple[str, ...] = ("PEVST", "PPDHD", "PPDIT", "PPOIX", "PPDIX")

# Nomes preferidos por conceito (validados por DDIC); a heurística cobre o resto.
_PPDHD_PREF = {"run": "RUNID", "doc": "DOCNUM", "company": "BUKRS",
               "budat": "BUDAT", "gjahr": "GJAHR", "monat": "MONAT",
               "doctyp": "DOCTYP", "revdoc": "REVDOC", "xblnr": "XBLNR", "blart": "BLART"}
_PPDIT_PREF = {"doc": "DOCNUM", "line": "DOCLIN", "account": "HKONT", "company": "BUKRS",
               "currency": "WAERS", "amount": "WRBTR", "neg": "NEG_POSTNG",
               "ktosl": "KTOSL", "pernr": "PERNR", "ittyp": "ITTYP", "text": "SGTXT"}


@dataclass
class PostingHeader:
    run_id: str
    doc_number: str
    company: str = ""
    budat: str = ""
    doc_type: str = ""
    rev_doc: str = ""
    reference: str = ""


@dataclass
class PayrollPostingReport:
    table_diags: dict[str, TableDiag] = field(default_factory=dict)
    field_guesses: dict[str, dict[str, FieldGuess]] = field(default_factory=dict)
    resolved_fields: dict[str, str | None] = field(default_factory=dict)

    headers: list[PostingHeader] = field(default_factory=list)
    doc_to_run: dict[str, str] = field(default_factory=dict)
    run_to_docs: dict[str, list[str]] = field(default_factory=dict)

    # todos os itens dos documentos dos runs (para diagnóstico por conta/empresa)
    all_items: list[PostingItem] = field(default_factory=list)
    by_company_account: list[dict[str, Any]] = field(default_factory=list)

    # itens da conta em análise
    items: list[PostingItem] = field(default_factory=list)
    totals_by_run: dict[str, Decimal] = field(default_factory=dict)
    totals_by_run_company: dict[tuple[str, str], Decimal] = field(default_factory=dict)
    total: Decimal = Decimal("0")

    runs_found: dict[str, bool] = field(default_factory=dict)
    runs_with_account: dict[str, bool] = field(default_factory=dict)
    companies_with_account: list[str] = field(default_factory=list)
    match_company: str | None = None  # empresa cujo total bate com a referência FI
    match_runs: list[str] = field(default_factory=list)  # runs cujo valor == referência FI
    duplicate_run_groups: list[list[str]] = field(default_factory=list)

    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)


def diagnose_tables(connection: Any) -> dict[str, TableDiag]:
    out: dict[str, TableDiag] = {}
    for table in PAYROLL_TABLES:
        logger.info("A verificar tabela %s", table)
        diag = describe_table(connection, table)
        logger.info("%s: existe=%s autorizado=%s campos=%s (%s)",
                    table, diag.exists, diag.authorized, diag.field_count, diag.note or "OK")
        out[table] = diag
    return out


def _resolve(diag: TableDiag | None, prefs: dict[str, str], guesses: dict[str, FieldGuess],
             concept_map: dict[str, str]) -> dict[str, str | None]:
    """Nome preferido se existir na tabela; senão o melhor palpite semântico."""
    names = {n.upper() for n in (diag.field_names() if diag else [])}
    out: dict[str, str | None] = {}
    for key, pref in prefs.items():
        if pref.upper() in names:
            out[key] = pref
            continue
        concept = concept_map.get(key)
        g = guesses.get(concept) if concept else None
        out[key] = g.chosen if g and g.chosen else None
    return out


def analyze(connection: Any, params: AnalysisParams) -> PayrollPostingReport:
    report = PayrollPostingReport()
    report.table_diags = diagnose_tables(connection)
    for name, diag in report.table_diags.items():
        if diag.exists and diag.fields:
            report.field_guesses[name] = guess_fields(diag)

    ppdhd = report.table_diags.get("PPDHD")
    ppdit = report.table_diags.get("PPDIT")
    runs = [pad_run(r) for r in params.posting_runs]

    # ------------------------------------------------------------------ headers
    hdr_f = _resolve(
        ppdhd, _PPDHD_PREF, report.field_guesses.get("PPDHD", {}),
        {"run": "posting_run", "doc": "documento", "company": "empresa",
         "budat": "data_lancamento", "gjahr": "exercicio", "monat": "periodo"},
    )
    report.resolved_fields.update({f"PPDHD.{k}": v for k, v in hdr_f.items()})

    if not (ppdhd and ppdhd.exists and ppdhd.authorized and hdr_f["run"] and hdr_f["doc"]):
        report.warn("PPDHD indisponível ou sem RUNID/DOCNUM: não é possível mapear runs -> documentos.")
        return report

    _load_headers(connection, params, hdr_f, runs, report)
    if not report.headers:
        report.warn(f"Nenhum documento PPDHD para os runs {runs}.")
        # ainda assim continua: runs_found já preenchido

    # ------------------------------------------------------------------- items
    it_f = _resolve(
        ppdit, _PPDIT_PREF, report.field_guesses.get("PPDIT", {}),
        {"doc": "documento", "line": "item", "account": "conta", "company": "empresa",
         "currency": "moeda", "amount": "valor"},
    )
    report.resolved_fields.update({f"PPDIT.{k}": v for k, v in it_f.items()})

    if not (ppdit and ppdit.exists and hdr_f["doc"] and it_f["doc"] and it_f["account"] and it_f["amount"]):
        report.warn(
            "PPDIT: não foi possível resolver DOCNUM/HKONT/WRBTR "
            f"(doc={it_f['doc']} conta={it_f['account']} valor={it_f['amount']}). "
            "Etapa de itens interrompida."
        )
        return report

    docs = sorted(report.doc_to_run)
    if not docs:
        report.warn("Sem DOCNUMs para ler em PPDIT.")
        return report

    _load_items(connection, params, it_f, docs, report)
    report.resolved = True
    _summarise(params, report)
    return report


def _load_headers(
    connection: Any, params: AnalysisParams, f: dict[str, str | None],
    runs: list[str], report: PayrollPostingReport,
) -> None:
    fields = _distinct([f["doc"], f["run"], f["company"], f["budat"], f["doctyp"],
                        f["revdoc"], f["xblnr"]])
    try:
        rows = read_table(
            connection, "PPDHD", fields=fields,
            options=opt_and(opt_in(f["run"], runs)), page_size=params.page_size,
        ).rows
    except RfcReadError as exc:
        report.warn(f"Falha a ler PPDHD: {exc}")
        rows = []

    for r in rows:
        h = PostingHeader(
            run_id=r.get(f["run"], "").strip(),
            doc_number=r.get(f["doc"], "").strip(),
            company=r.get(f["company"], "").strip() if f["company"] else "",
            budat=r.get(f["budat"], "").strip() if f["budat"] else "",
            doc_type=r.get(f["doctyp"], "").strip() if f["doctyp"] else "",
            rev_doc=r.get(f["revdoc"], "").strip() if f["revdoc"] else "",
            reference=r.get(f["xblnr"], "").strip() if f["xblnr"] else "",
        )
        report.headers.append(h)
        if h.doc_number:
            report.doc_to_run[h.doc_number] = h.run_id
            report.run_to_docs.setdefault(h.run_id, []).append(h.doc_number)

    for run in runs:
        report.runs_found[run] = run in report.run_to_docs


def _load_items(
    connection: Any, params: AnalysisParams, f: dict[str, str | None],
    docs: list[str], report: PayrollPostingReport,
) -> None:
    fields = _distinct([f["doc"], f["line"], f["company"], f["account"], f["ktosl"],
                        f["currency"], f["neg"], f["pernr"], f["ittyp"], f["amount"], f["text"]])
    rows: list[dict[str, str]] = []
    for start in range(0, len(docs), 100):
        chunk = docs[start : start + 100]
        try:
            rows.extend(
                read_table(
                    connection, "PPDIT", fields=fields,
                    options=opt_and(opt_in(f["doc"], chunk)), page_size=params.page_size,
                ).rows
            )
        except NoData:
            continue
        except RfcReadError as exc:
            report.warn(f"Falha a ler PPDIT (chunk {start}): {exc}")
            return

    acct_targets = {params.conta_10, params.conta.strip()}
    agg: dict[tuple[str, str], list[Any]] = defaultdict(lambda: [0, Decimal("0"), Decimal("0")])

    for r in rows:
        raw_amt = r.get(f["amount"], "")
        amount = sap_str_to_decimal(raw_amt)
        neg = (r.get(f["neg"], "").strip().upper() == "X") if f["neg"] else False
        signed = -amount if neg else amount
        company = r.get(f["company"], "").strip() if f["company"] else ""
        account = r.get(f["account"], "").strip()
        doc = r.get(f["doc"], "").strip()
        run_id = report.doc_to_run.get(doc, "")

        item = PostingItem(
            run_id=run_id,
            doc_number=doc,
            line=r.get(f["line"], "").strip() if f["line"] else "",
            account=account,
            company=company,
            currency=r.get(f["currency"], "").strip() if f["currency"] else "",
            debit_credit="NEG" if neg else "",
            amount_raw=raw_amt,
            amount=abs(amount),
            signed_amount=signed,
            raw=r,
        )
        report.all_items.append(item)
        a = agg[(company, account)]
        a[0] += 1
        a[1] += signed
        a[2] += abs(signed)

        if account in acct_targets or account.lstrip("0") == params.conta.strip().lstrip("0"):
            report.items.append(item)
            report.totals_by_run[run_id] = report.totals_by_run.get(run_id, Decimal("0")) + signed
            key = (run_id, company)
            report.totals_by_run_company[key] = report.totals_by_run_company.get(key, Decimal("0")) + signed

    report.by_company_account = sorted(
        ({"company": k[0], "account": k[1], "lines": v[0],
          "signed_sum": str(v[1]), "abs_sum": str(v[2])} for k, v in agg.items()),
        key=lambda d: -abs(Decimal(d["signed_sum"])),
    )


def _summarise(params: AnalysisParams, report: PayrollPostingReport) -> None:
    report.total = sum((i.signed_amount for i in report.items), Decimal("0"))

    companies = sorted({i.company for i in report.items if i.company})
    report.companies_with_account = companies

    for run in report.runs_found:
        has = any(i.run_id == run for i in report.items)
        report.runs_with_account[run] = has

    # empresa / run cujo total (abs) coincide com a referência FI informada
    ref = params.valor_fi_referencia
    by_company: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    for i in report.items:
        by_company[i.company] += i.signed_amount
    for comp, tot in by_company.items():
        if abs(abs(tot) - ref) <= params.tolerancia:
            report.match_company = comp
            break

    for (run, comp), tot in report.totals_by_run_company.items():
        if abs(abs(tot) - ref) <= params.tolerancia:
            report.match_runs.append(run)
            if report.match_company is None:
                report.match_company = comp
    report.match_runs.sort()

    # runs que lançaram exactamente o mesmo valor na mesma empresa (possível duplicação)
    seen: dict[tuple[str, str], list[str]] = defaultdict(list)
    for (run, comp), tot in report.totals_by_run_company.items():
        seen[(comp, str(tot))].append(run)
    report.duplicate_run_groups = [sorted(v) for v in seen.values() if len(v) > 1]

    if params.empresa not in companies and companies:
        report.warn(
            f"Empresa {params.empresa} não tem lançamentos na conta {params.conta} "
            f"nestes runs. Empresas com movimento: {', '.join(companies)}."
            + (f" A empresa {report.match_company} bate com a referência FI "
               f"({ref})." if report.match_company else "")
        )
    if report.match_runs:
        report.warn(
            f"Runs cujo valor em {params.conta} é exactamente a referência FI "
            f"{ref}: {', '.join(report.match_runs)}."
        )
    for grp in report.duplicate_run_groups:
        report.warn(
            f"Runs {', '.join(grp)} lançaram o MESMO valor na mesma empresa/conta "
            "— possível execução repetida; confirmar qual foi transferida para FI."
        )
    logger.info("Posting RH: %s itens conta-alvo, total %s (empresas %s)",
                len(report.items), report.total, companies)


def _distinct(values: list[str | None]) -> list[str]:
    seen: list[str] = []
    for v in values:
        if v and v not in seen:
            seen.append(v)
    return seen
