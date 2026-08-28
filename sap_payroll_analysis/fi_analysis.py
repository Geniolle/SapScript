"""Análise FI: o que foi efectivamente contabilizado na conta do Razão.

Detecção de release e de acesso:

* S/4HANA -> `ACDOCA` (Universal Journal), se existir e for legível.
* ECC / clássico -> partidas individuais em `BSIS` (em aberto) + `BSAS`
  (compensadas). `BSEG` é tentado como último recurso.
* Saldo de controlo por período via BAPI **read-only**
  `BAPI_GL_GETGLACCPERIODBALANCES` (não altera nada, não faz COMMIT).

Convenção de sinal (igual à do posting RH): Débito = +, Crédito = -.
`RFC_READ_TABLE` levanta `TABLE_WITHOUT_DATA` quando nada corresponde ao
filtro — isso é tratado como "sem linhas", não como erro.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from .config import AnalysisParams
from .ddic import describe_table, guess_fields
from .models import FIItem, TableDiag
from .sap_reader import (
    NoData,
    RfcReadError,
    normalize_sign,
    opt_and,
    opt_eq,
    opt_in,
    read_table,
    sap_str_to_decimal,
)
from .security import safe_rfc_call

logger = logging.getLogger(__name__)

_LINE_TABLES = ("BSIS", "BSAS")
_BSIS_FIELDS = ["BUKRS", "HKONT", "GJAHR", "BELNR", "BUZEI", "SHKZG", "DMBTR", "WRBTR",
                "WAERS", "PSWSL", "BUDAT", "BLDAT", "MONAT", "BLART", "XBLNR", "ZUONR",
                "SGTXT", "AWTYP", "AWKEY"]


@dataclass
class FIReport:
    source: str = ""  # "ACDOCA" | "BSIS/BSAS" | "BSEG" | ""
    company_used: str = ""  # empresa que efectivamente produziu dados FI
    table_diags: dict[str, TableDiag] = field(default_factory=dict)
    resolved_fields: dict[str, str] = field(default_factory=dict)
    items: list[FIItem] = field(default_factory=list)
    total: Decimal = Decimal("0")
    total_debit: Decimal = Decimal("0")
    total_credit: Decimal = Decimal("0")
    bapi_period_balance: dict[str, Any] = field(default_factory=dict)
    bapi_messages: list[str] = field(default_factory=list)
    companies_tried: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    resolved: bool = False

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)
            logger.warning(msg)


def analyze(connection: Any, params: AnalysisParams, extra_companies: list[str] | None = None) -> FIReport:
    report = FIReport()
    for table in ("ACDOCA", "BKPF", "BSEG", "BSIS", "BSAS"):
        report.table_diags[table] = describe_table(connection, table)
        d = report.table_diags[table]
        logger.info("%s: existe=%s autorizado=%s campos=%s", table, d.exists, d.authorized, d.field_count)

    companies: list[str] = []
    for comp in [params.empresa, *(extra_companies or [])]:
        if comp and comp not in companies:
            companies.append(comp)

    acdoca = report.table_diags.get("ACDOCA")
    for comp in companies:
        report.companies_tried.append(comp)
        cp = _with_company(params, comp)

        _bapi_period_balance(connection, cp, report)

        if acdoca and acdoca.exists and acdoca.authorized and acdoca.fields:
            try:
                _analyze_acdoca(connection, cp, acdoca, report)
            except RfcReadError as exc:
                report.warn(f"ACDOCA falhou ({exc}); a tentar BSIS/BSAS.")
        if not report.resolved:
            _analyze_bsis_bsas(connection, cp, report)
        if not report.resolved:
            _analyze_bseg(connection, cp, report)

        if report.items:
            report.company_used = comp
            break
        if report.bapi_period_balance.get("period_movement") is not None:
            report.company_used = comp
            mv = Decimal(str(report.bapi_period_balance["period_movement"]))
            report.total = mv
            report.resolved = True
            report.source = report.source or "BAPI_GL_GETGLACCPERIODBALANCES"
            break

    if not report.resolved and not report.items:
        report.warn(
            f"FI: nenhuma das empresas testadas ({', '.join(report.companies_tried)}) "
            f"tem movimento na conta {params.conta} no exercício {params.gjahr} neste "
            f"sistema/mandante. O valor 'integrado em FI' poderá estar noutro sistema "
            f"(ex.: PRD/QAD) ou a transferência RH->FI ainda não foi executada."
        )

    return report


def _with_company(params: AnalysisParams, company: str) -> AnalysisParams:
    if company == params.empresa:
        return params
    from dataclasses import replace

    return replace(params, empresa=company)


# ---------------------------------------------------------------------------
# BAPI de saldo por período (read-only)
# ---------------------------------------------------------------------------

def _bapi_period_balance(connection: Any, params: AnalysisParams, report: FIReport) -> None:
    try:
        res = safe_rfc_call(
            connection,
            "BAPI_GL_GETGLACCPERIODBALANCES",
            COMPANYCODE=params.empresa,
            FISCALYEAR=params.gjahr,
            GLACCT=params.conta_10,
            CURRENCYTYPE="10",
        )
    except Exception as exc:  # noqa: BLE001
        report.warn(f"BAPI_GL_GETGLACCPERIODBALANCES indisponível: {exc}")
        return

    rows = res.get("ACCOUNT_BALANCES", []) or []
    ret = res.get("RETURN", {}) or {}
    ret_list = ret if isinstance(ret, list) else [ret]
    for r in ret_list:
        msg = str(r.get("MESSAGE", "")).strip() if isinstance(r, dict) else ""
        if msg:
            tag = f"[{r.get('TYPE', '?')}] empresa {params.empresa}: {msg}"
            if tag not in report.bapi_messages:
                report.bapi_messages.append(tag)
            if r.get("TYPE") in {"E", "A"}:
                report.warn(f"BAPI RETURN {tag}")

    target = None
    for r in rows:
        per = str(r.get("PERIOD", "") or r.get("FIS_PERIOD", "")).lstrip("0")
        if per == str(params.mes):
            target = r
            break

    def _d(*keys: str) -> Decimal | None:
        for r in ([target] if target else []) + list(rows[:1] if not target else []):
            for k in keys:
                if k in r and str(r[k]).strip() not in {"", "0", "0.0"}:
                    try:
                        return sap_str_to_decimal(str(r[k]))
                    except ValueError:
                        return None
        return None

    debit = _d("DEBIT_BAL", "DEBITBALANCE", "TOT_DEBIT", "DEBIT")
    credit = _d("CREDIT_BAL", "CREDITBALANCE", "TOT_CREDIT", "CREDIT")
    balance = _d("BALANCE", "PERIOD_BAL", "BALANCE_PER")
    movement = None
    if debit is not None or credit is not None:
        movement = (debit or Decimal("0")) - (credit or Decimal("0"))
    elif balance is not None:
        movement = balance

    report.bapi_period_balance = {
        "period": params.mes,
        "raw_rows": rows,
        "debit": None if debit is None else str(debit),
        "credit": None if credit is None else str(credit),
        "balance": None if balance is None else str(balance),
        "period_movement": None if movement is None else str(movement),
    }
    if movement is not None:
        logger.info("BAPI saldo período %s: mov=%s (D=%s C=%s)", params.mes, movement, debit, credit)


# ---------------------------------------------------------------------------
# ACDOCA (S/4)
# ---------------------------------------------------------------------------

def _analyze_acdoca(connection: Any, params: AnalysisParams, diag: TableDiag, report: FIReport) -> None:
    names = {n.upper() for n in diag.field_names()}
    g = guess_fields(diag)
    company = _first(names, ["RBUKRS", "BUKRS"]) or g["empresa"].chosen or ""
    account = _first(names, ["RACCT", "HKONT"]) or g["conta"].chosen or ""
    year = _first(names, ["RYEAR", "GJAHR"]) or g["exercicio"].chosen or ""
    period = _first(names, ["POPER"]) or g["periodo"].chosen or ""
    drcrk = _first(names, ["DRCRK"]) or g["debito_credito"].chosen or ""
    amount = _first(names, ["WSL", "HSL", "TSL"]) or g["valor"].chosen or ""
    currency = _first(names, ["RWCUR", "RTCUR", "RHCUR"]) or ""
    doc = _first(names, ["BELNR"]) or "BELNR"
    docln = _first(names, ["DOCLN"]) or "DOCLN"
    budat = _first(names, ["BUDAT"]) or ""
    blart = _first(names, ["BLART"]) or ""
    ledger = _first(names, ["RLDNR"]) or ""
    report.resolved_fields = {"empresa": company, "conta": account, "exercicio": year,
                              "periodo": period, "sinal": drcrk, "valor": amount}
    if not (company and account and amount and year):
        report.warn("ACDOCA: campos essenciais não resolvidos.")
        return
    groups = [opt_eq(company, params.empresa),
              opt_in(account, sorted({params.conta_10, params.conta.strip()})),
              opt_eq(year, params.gjahr)]
    if period:
        groups.append(opt_eq(period, params.poper))
    if ledger:
        groups.append(opt_eq(ledger, "0L"))
    fields = _distinct([doc, docln, company, account, year, period, budat, blart, drcrk, currency, amount])
    rows = read_table(connection, "ACDOCA", fields=fields, options=opt_and(*groups),
                      page_size=params.page_size).rows
    report.source = "ACDOCA"
    for r in rows:
        val = sap_str_to_decimal(r.get(amount, "0"))
        flag = r.get(drcrk, "") if drcrk else ""
        signed = normalize_sign(val, flag) if flag else val
        report.items.append(FIItem(
            source="ACDOCA", document=r.get(doc, "").strip(), fiscal_year=r.get(year, "").strip(),
            line=r.get(docln, "").strip() if docln else "", posting_date=r.get(budat, "").strip() if budat else "",
            period=r.get(period, "").strip() if period else "", account=r.get(account, "").strip(),
            company=r.get(company, "").strip(), currency=r.get(currency, "").strip() if currency else "",
            debit_credit=flag, amount_raw=r.get(amount, ""), amount=abs(val), signed_amount=signed,
            doc_type=r.get(blart, "").strip() if blart else "", raw=r,
        ))
    if report.items:
        _finalize(report)
        report.resolved = True


# ---------------------------------------------------------------------------
# BSIS + BSAS (ECC)
# ---------------------------------------------------------------------------

def _analyze_bsis_bsas(connection: Any, params: AnalysisParams, report: FIReport) -> None:
    before = len(report.items)
    for table in _LINE_TABLES:
        diag = report.table_diags.get(table)
        if not (diag and diag.exists and diag.authorized):
            continue
        avail = {n.upper() for n in diag.field_names()}
        fields = [f for f in _BSIS_FIELDS if f in avail] or _BSIS_FIELDS
        groups = [
            opt_eq("BUKRS", params.empresa),
            opt_in("HKONT", sorted({params.conta_10, params.conta.strip()})),
            opt_eq("GJAHR", params.gjahr),
        ]
        try:
            rows = read_table(connection, table, fields=fields, options=opt_and(*groups),
                              page_size=params.page_size).rows
        except NoData:
            rows = []
        except RfcReadError as exc:
            report.warn(f"{table}: {exc}")
            continue

        for r in rows:
            monat = r.get("MONAT", "").strip()
            if monat and monat.lstrip("0") not in {str(params.mes), ""}:
                continue
            val = sap_str_to_decimal(r.get("WRBTR", "0"))
            flag = r.get("SHKZG", "")
            report.items.append(FIItem(
                source=table, document=r.get("BELNR", "").strip(), fiscal_year=r.get("GJAHR", "").strip(),
                line=r.get("BUZEI", "").strip(), posting_date=r.get("BUDAT", "").strip(),
                period=monat, account=r.get("HKONT", "").strip(), company=r.get("BUKRS", "").strip(),
                currency=(r.get("WAERS") or r.get("PSWSL") or "").strip(), debit_credit=flag,
                amount_raw=r.get("WRBTR", ""), amount=abs(val), signed_amount=normalize_sign(val, flag),
                doc_type=r.get("BLART", "").strip(), reference=r.get("XBLNR", "").strip(),
                text=r.get("SGTXT", "").strip(), raw=r,
            ))

    if len(report.items) > before:
        report.source = "BSIS/BSAS"
        _finalize(report)
        report.resolved = True


# ---------------------------------------------------------------------------
# BSEG (último recurso — pode não ser legível via RFC_READ_TABLE)
# ---------------------------------------------------------------------------

def _analyze_bseg(connection: Any, params: AnalysisParams, report: FIReport) -> None:
    diag = report.table_diags.get("BSEG")
    if not (diag and diag.exists and diag.authorized):
        return
    fields = ["BUKRS", "BELNR", "GJAHR", "BUZEI", "HKONT", "SHKZG", "WRBTR", "PSWSL"]
    groups = [opt_eq("BUKRS", params.empresa),
              opt_in("HKONT", sorted({params.conta_10, params.conta.strip()})),
              opt_eq("GJAHR", params.gjahr)]
    try:
        bseg_rows = read_table(connection, "BSEG", fields=fields, options=opt_and(*groups),
                               page_size=params.page_size).rows
    except NoData:
        bseg_rows = []
    except RfcReadError as exc:
        report.warn(f"BSEG não legível via RFC_READ_TABLE: {exc}")
        return

    if not bseg_rows:
        return
    belnrs = sorted({r["BELNR"] for r in bseg_rows if r.get("BELNR")})
    head = _bkpf_period_index(connection, params, belnrs)
    for r in bseg_rows:
        h = head.get((r.get("BELNR", ""), r.get("GJAHR", "")))
        if head and h is None:
            continue
        if h and h.get("MONAT", "").lstrip("0") not in {str(params.mes), ""}:
            continue
        val = sap_str_to_decimal(r.get("WRBTR", "0"))
        flag = r.get("SHKZG", "")
        report.items.append(FIItem(
            source="BSEG", document=r.get("BELNR", "").strip(), fiscal_year=r.get("GJAHR", "").strip(),
            line=r.get("BUZEI", "").strip(), period=(h or {}).get("MONAT", ""),
            posting_date=(h or {}).get("BUDAT", ""), account=r.get("HKONT", "").strip(),
            company=r.get("BUKRS", "").strip(), currency=r.get("PSWSL", "").strip(),
            debit_credit=flag, amount_raw=r.get("WRBTR", ""), amount=abs(val),
            signed_amount=normalize_sign(val, flag), doc_type=(h or {}).get("BLART", ""), raw=r,
        ))
    if report.items:
        report.source = "BSEG"
        _finalize(report)
        report.resolved = True


def _bkpf_period_index(connection: Any, params: AnalysisParams, belnrs: list[str]) -> dict[tuple[str, str], dict[str, str]]:
    index: dict[tuple[str, str], dict[str, str]] = {}
    diag = None  # BKPF assumed transparent
    for start in range(0, len(belnrs), 60):
        chunk = belnrs[start : start + 60]
        try:
            rows = read_table(connection, "BKPF",
                              fields=["BUKRS", "BELNR", "GJAHR", "MONAT", "BUDAT", "BLART", "XBLNR", "STBLG"],
                              options=opt_and(opt_eq("BUKRS", params.empresa), opt_eq("GJAHR", params.gjahr),
                                              opt_in("BELNR", chunk)),
                              page_size=params.page_size).rows
        except (NoData, RfcReadError):
            continue
        for r in rows:
            index[(r.get("BELNR", ""), r.get("GJAHR", ""))] = r
    return index


# ---------------------------------------------------------------------------

def _finalize(report: FIReport) -> None:
    report.total = sum((i.signed_amount for i in report.items), Decimal("0"))
    report.total_debit = sum((i.signed_amount for i in report.items if i.signed_amount > 0), Decimal("0"))
    report.total_credit = sum((i.signed_amount for i in report.items if i.signed_amount < 0), Decimal("0"))


def _first(available: set[str], candidates: list[str]) -> str:
    for c in candidates:
        if c.upper() in available:
            return c
    return ""


def _distinct(values: list[str]) -> list[str]:
    seen: list[str] = []
    for v in values:
        if v and v not in seen:
            seen.append(v)
    return seen
