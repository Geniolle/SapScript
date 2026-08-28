"""Fase 5.0 — reconciliação RH (Payroll) × programa de pagamentos (REGU*).

Pergunta única:

    Para cada colaborador, o valor que o RH determinou como líquido a
    transferir (`/559` corrente) é EXACTAMENTE o valor que o programa de
    pagamentos levou às tabelas REGU*?  E `SUM(RH) == SUM(REGU)`?

Corre obrigatoriamente no R/3 (`SAP_R3_*`). Sem esses parâmetros, aborta —
nunca usa o fallback `SAP_DEV_*` (a Fase 4.x mostrou que aponta para outro
sistema, sem estes dados).

100 % READ-ONLY: só `RFC_READ_TABLE` e leitura de DDIC via `safe_rfc_call`.
NÃO consulta S/4H. NÃO usa PCL2 / PYXX. NÃO executa F110 nem pagamentos.
"""

from __future__ import annotations

import csv
import json
import logging
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import AnalysisParams
from .sap_reader import NoData, RfcReadError, opt_and, opt_eq, opt_in, read_table
from .wagetype_trace import _signed

logger = logging.getLogger(__name__)

_ZERO = Decimal("0")

#: rubrica que representa a transferência bancária líquida corrente.
BANK_TRANSFER_LGART = "/559"
#: rubricas relacionadas (para a tabela de composição — NÃO somadas cegamente).
RELATED_LGARTS = ("/558", "/559", "/561", "/563", "0029")

REGU_TABLES = ("REGUH", "REGUP", "REGUV", "REGUHM", "REGUT")

#: diferença conhecida de OUTRA análise. NUNCA forçar REGU a explicá-la.
KNOWN_427 = Decimal("427.74")

#: nomes preferidos por papel (o 1.º presente no DDIC ganha).
_ROLE_FIELDS: dict[str, tuple[str, ...]] = {
    "company": ("ZBUKR", "BUKRS", "ABSBU"),
    "run_date": ("LAUFD",),
    "run_id": ("LAUFI",),
    "proposal_flag": ("XVORL",),
    "real_flag": ("XECHT",),
    "amount": ("RWBTR", "RBETR", "WRBTR", "DMBTR", "NEBTR", "SKBTR"),
    "currency": ("WAERS", "RWCUR", "PSWSL"),
    "vendor": ("LIFNR",),
    "customer": ("KUNNR",),
    "payee": ("EMPFG",),
    "pernr": ("PERNR",),
    "payment_doc": ("VBLNR",),
    "acct_doc": ("BELNR",),
    "fiscal_year": ("GJAHR",),
    "payment_method": ("RZAWE", "ZLSCH"),
    "name": ("ZNME1", "NAME1", "ZNAME1"),
    "bank_country": ("ZBNKS", "UBNKS"),
    "bank_key": ("ZBNKL", "UBNKL", "UBNKY"),
    "bank_account": ("ZBNKN", "UBKNT", "UBNKN"),
    "house_bank": ("HBKID",),
    "reference": ("XBLNR", "ZUONR", "SGTXT", "KIDNO"),
    "posting_date": ("ZALDT", "VALUT", "BUDAT"),
    "line_item": ("BUZEI",),
}


def _q(v: Any) -> Decimal:
    try:
        return Decimal(str(v))
    except Exception:  # noqa: BLE001
        return _ZERO


def _digits(v: str) -> str:
    return "".join(ch for ch in str(v or "") if ch.isdigit())


# ---------------------------------------------------------------------------
# Modelos
# ---------------------------------------------------------------------------

@dataclass
class PayrollPaymentExpectation:
    pernr: str = ""
    payroll_run: str = ""
    seqno: str = ""
    lgart: str = BANK_TRANSFER_LGART
    expected_payment_amount: Decimal = _ZERO
    currency: str = "EUR"
    derivation: str = ""
    confidence: str = ""            # PROVED / OBSERVED / CANDIDATE / HYPOTHESIS
    evidence: str = ""
    related_amounts: dict[str, str] = field(default_factory=dict)
    line_classes: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "PERNR": self.pernr, "payroll_run": self.payroll_run, "seqno": self.seqno,
            "lgart": self.lgart, "expected_payment_amount": str(self.expected_payment_amount),
            "currency": self.currency, "derivation": self.derivation,
            "confidence": self.confidence, "evidence": self.evidence,
            "related_amounts": self.related_amounts, "line_classes": self.line_classes,
        }


@dataclass
class ReguPayment:
    laufd: str = ""
    laufi: str = ""
    company: str = ""
    beneficiary_key: str = ""
    pernr: str = ""
    vendor: str = ""
    customer: str = ""
    payee: str = ""
    payment_doc: str = ""
    acct_doc: str = ""
    fiscal_year: str = ""
    amount: Decimal = _ZERO
    currency: str = ""
    payment_method: str = ""
    bank_reference: str = ""
    name: str = ""
    reference: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "laufd": self.laufd, "laufi": self.laufi, "company": self.company,
            "beneficiary_key": self.beneficiary_key, "PERNR": self.pernr,
            "vendor": self.vendor, "customer": self.customer, "payee": self.payee,
            "payment_doc": self.payment_doc, "acct_doc": self.acct_doc,
            "fiscal_year": self.fiscal_year, "amount": str(self.amount),
            "currency": self.currency, "payment_method": self.payment_method,
            "bank_reference": self.bank_reference, "name": self.name,
            "reference": self.reference,
        }


@dataclass
class PaymentRunCandidate:
    laufd: str = ""
    laufi: str = ""
    company: str = ""
    is_real: str = ""              # X / "" / "?"
    payment_count: int = 0
    total: Decimal = _ZERO
    currency: str = ""
    first_date: str = ""
    last_date: str = ""
    score: int = 0
    confidence: str = ""           # HIGH_CONFIDENCE / MEDIUM_CONFIDENCE / LOW_CONFIDENCE
    evidence: list[str] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "laufd": self.laufd, "laufi": self.laufi, "company": self.company,
            "is_real": self.is_real, "payment_count": self.payment_count,
            "total": str(self.total), "currency": self.currency,
            "first_date": self.first_date, "last_date": self.last_date,
            "score": self.score, "confidence": self.confidence, "evidence": self.evidence,
        }


@dataclass
class ReconLine:
    pernr: str = ""
    rh_expected: Decimal | None = None
    regu_paid: Decimal | None = None
    difference: Decimal | None = None
    status: str = ""               # EXACT_MATCH/DIFFERENCE/RH_ONLY/REGU_ONLY/AMBIGUOUS/UNMATCHED
    match_method: str = ""         # PERNR_DIRECT/EMPFG_PERNR/LIFNR_MAP/DOC_REF/VALUE_CANDIDATE/NONE
    seqno: str = ""
    payment_doc: str = ""
    payment_run_date: str = ""
    payment_run_id: str = ""
    rh_rows: int = 0
    regu_rows: int = 0
    note: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "PERNR": self.pernr,
            "RH_EXPECTED": None if self.rh_expected is None else str(self.rh_expected),
            "REGU_PAID": None if self.regu_paid is None else str(self.regu_paid),
            "DIFFERENCE": None if self.difference is None else str(self.difference),
            "STATUS": self.status, "MATCH_METHOD": self.match_method, "SEQNO": self.seqno,
            "PAYMENT_DOC": self.payment_doc, "PAYMENT_RUN_DATE": self.payment_run_date,
            "PAYMENT_RUN_ID": self.payment_run_id, "rh_rows": self.rh_rows,
            "regu_rows": self.regu_rows, "note": self.note,
        }


@dataclass
class PaymentReconciliation:
    payroll_run: str = ""
    company: str = ""
    period: str = ""
    schema: dict[str, Any] = field(default_factory=dict)
    payment_run_candidates: list[PaymentRunCandidate] = field(default_factory=list)
    selected_payment_run: dict[str, Any] = field(default_factory=dict)
    employee_identity_mapping: dict[str, Any] = field(default_factory=dict)
    payroll_expected: list[PayrollPaymentExpectation] = field(default_factory=list)
    regu_payments: list[ReguPayment] = field(default_factory=list)
    reconciliation: list[ReconLine] = field(default_factory=list)
    totals: dict[str, Any] = field(default_factory=dict)
    differences: list[dict[str, Any]] = field(default_factory=list)
    classification: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def warn(self, m: str) -> None:
        if m not in self.warnings:
            self.warnings.append(m)
            logger.warning(m)

    def as_dict(self) -> dict[str, Any]:
        return {
            "payroll": {"run": self.payroll_run, "company": self.company, "period": self.period},
            "regu_schema": self.schema,
            "payment_run_candidates": [c.as_dict() for c in self.payment_run_candidates],
            "selected_payment_run": self.selected_payment_run,
            "employee_identity_mapping": self.employee_identity_mapping,
            "payroll_expected": [e.as_dict() for e in self.payroll_expected],
            "regu_payments": [p.as_dict() for p in self.regu_payments],
            "reconciliation": [r.as_dict() for r in self.reconciliation],
            "totals": self.totals,
            "differences": self.differences,
            "classification": self.classification,
            "warnings": self.warnings,
        }


# ---------------------------------------------------------------------------
# 1 — DDIC das tabelas REGU*
# ---------------------------------------------------------------------------

def _table_fields(connection: Any, table: str) -> list[str]:
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
    rows.sort(key=lambda r: int(r["POSITION"]) if str(r.get("POSITION", "")).isdigit() else 0)
    return [r["FIELDNAME"].strip() for r in rows]


def _resolve_roles(fieldset: Sequence[str]) -> dict[str, str]:
    present = set(fieldset)
    roles: dict[str, str] = {}
    for role, names in _ROLE_FIELDS.items():
        for n in names:
            if n in present:
                roles[role] = n
                break
    return roles


def inspect_regu_schema(connection: Any, params: AnalysisParams) -> dict[str, Any]:
    out: dict[str, Any] = {"tables": {}}
    for t in REGU_TABLES:
        flds = _table_fields(connection, t)
        if not flds:
            out["tables"][t] = {"exists": False}
            continue
        roles = _resolve_roles(flds)
        out["tables"][t] = {
            "exists": True,
            "n_fields": len(flds),
            "fields": flds,
            "roles": roles,
            "has_pernr": "pernr" in roles,
            "has_payee": "payee" in roles,
            "amount_field": roles.get("amount", ""),
        }
    reguh = out["tables"].get("REGUH", {})
    out["conclusion"] = (
        "[PROVED] REGUH presente; campo de montante = "
        f"{reguh.get('amount_field') or 'n/d'}; identificação do colaborador "
        f"via {'PERNR directo' if reguh.get('has_pernr') else ('EMPFG' if reguh.get('has_payee') else 'LIFNR/indirecta')}."
        if reguh.get("exists") else
        "[OBSERVED] REGUH não encontrada por DDIC neste sistema."
    )
    return out


# ---------------------------------------------------------------------------
# 5 — descobrir runs de pagamento candidatos
# ---------------------------------------------------------------------------

def _period_window(period: str) -> tuple[str, str]:
    """202606 -> ('20260601', '20260715')  (mês + 1.ª quinzena do mês seguinte)."""
    y, m = int(period[:4]), int(period[4:6])
    start = f"{y:04d}{m:02d}01"
    nm_y, nm_m = (y + 1, 1) if m == 12 else (y, m + 1)
    end = f"{nm_y:04d}{nm_m:02d}15"
    return start, end


def discover_payment_runs(connection: Any, params: AnalysisParams, *, company: str,
                          period: str, schema: dict[str, Any],
                          window: tuple[str, str] | None = None) -> list[PaymentRunCandidate]:
    start, end = window or _period_window(period)
    reguh_roles = schema.get("tables", {}).get("REGUH", {}).get("roles", {})
    reguv_info = schema.get("tables", {}).get("REGUV", {})
    r_comp = reguh_roles.get("company", "ZBUKR")
    r_amt = reguh_roles.get("amount", "RWBTR")
    r_cur = reguh_roles.get("currency", "WAERS")
    r_pdate = reguh_roles.get("posting_date", "")

    # 1) universo de (LAUFD, LAUFI) — via REGUV se existir, senão distinct em REGUH
    runs: set[tuple[str, str, str]] = set()
    real_flag: dict[tuple[str, str], str] = {}
    if reguv_info.get("exists"):
        vr = reguv_info.get("roles", {})
        vcomp, vreal = vr.get("company", "ZBUKR"), vr.get("real_flag", "")
        want = [f for f in ["LAUFD", "LAUFI", vcomp, vreal] if f]
        try:
            for row in read_table(connection, "REGUV", fields=want,
                                  options=opt_and(opt_eq(vcomp, company)),
                                  page_size=100_000).rows:
                if start <= row.get("LAUFD", "") <= end:
                    key = (row["LAUFD"], row.get("LAUFI", ""), company)
                    runs.add(key)
                    if vreal:
                        real_flag[(key[0], key[1])] = row.get(vreal, "")
        except (RfcReadError, NoData) as exc:
            logger.warning("REGUV indisponível: %s", exc)

    # 2) sempre confirmar / completar por REGUH
    hwant = [f for f in ["LAUFD", "LAUFI", r_comp, r_amt, r_cur, r_pdate] if f]
    try:
        hrows_all = read_table(connection, "REGUH", fields=hwant,
                               options=opt_and(opt_eq(r_comp, company)),
                               page_size=200_000).rows
    except (RfcReadError, NoData) as exc:
        logger.warning("REGUH indisponível para descoberta: %s", exc)
        hrows_all = []
    hrows = [r for r in hrows_all if start <= r.get("LAUFD", "") <= end
             and r.get(r_comp) == company]
    for r in hrows:
        runs.add((r["LAUFD"], r.get("LAUFI", ""), company))

    cands: list[PaymentRunCandidate] = []
    for (laufd, laufi, comp) in sorted(runs):
        grp = [r for r in hrows if r.get("LAUFD") == laufd and r.get("LAUFI") == laufi]
        amounts = [_q(r.get(r_amt, "0")) for r in grp]
        curs = sorted({r.get(r_cur, "") for r in grp if r.get(r_cur)})
        dates = sorted({r.get(r_pdate, "") for r in grp if r_pdate and r.get(r_pdate)}) or [laufd]
        cands.append(PaymentRunCandidate(
            laufd=laufd, laufi=laufi, company=comp,
            is_real=real_flag.get((laufd, laufi), "?"),
            payment_count=len(grp), total=sum(amounts, _ZERO),
            currency=",".join(curs), first_date=dates[0], last_date=dates[-1],
        ))
    return cands


# ---------------------------------------------------------------------------
# 6 — escolher o run com múltiplas evidências (nunca só pelo total)
# ---------------------------------------------------------------------------

def select_payment_run(candidates: list[PaymentRunCandidate], *, period: str,
                       payroll_employee_count: int, payroll_currency: str = "EUR",
                       payroll_reference_total: Decimal | None = None,
                       ) -> tuple[dict[str, Any], list[PaymentRunCandidate]]:
    month_prefix = period[:6]
    for c in candidates:
        ev: list[str] = []
        score = 0
        ev.append(f"empresa={c.company}")
        score += 3
        if c.laufd.startswith(month_prefix):
            score += 2
            ev.append("LAUFD no mês do payroll")
        elif c.laufd[:6] and c.laufd[:6] > month_prefix:
            score += 1
            ev.append("LAUFD no mês seguinte (pagamento após processamento)")
        if payroll_employee_count and c.payment_count:
            ratio = c.payment_count / payroll_employee_count
            if 0.8 <= ratio <= 1.25:
                score += 3
                ev.append(f"nº beneficiários ~ nº colaboradores ({c.payment_count}/{payroll_employee_count})")
            elif 0.5 <= ratio <= 1.6:
                score += 1
                ev.append(f"nº beneficiários compatível ({c.payment_count}/{payroll_employee_count})")
        if payroll_currency and payroll_currency in (c.currency or ""):
            score += 1
            ev.append(f"moeda {payroll_currency}")
        if c.is_real == "X":
            score += 2
            ev.append("run real (XECHT=X)")
        elif c.is_real == "":
            ev.append("run é PROPOSTA (XECHT vazio)")
        # o total é evidência FRACA — peso mínimo
        if payroll_reference_total is not None and payroll_reference_total != 0:
            rel = abs(c.total) / abs(payroll_reference_total)
            if Decimal("0.95") <= rel <= Decimal("1.05"):
                score += 1
                ev.append("total ~ referência RH (peso baixo)")
        c.score = score
        c.evidence = ev
        c.confidence = ("HIGH_CONFIDENCE" if score >= 8
                        else "MEDIUM_CONFIDENCE" if score >= 5
                        else "LOW_CONFIDENCE")
    ranked = sorted(candidates, key=lambda c: (-c.score, c.laufd, c.laufi))
    if not ranked:
        return {}, ranked
    top = ranked[0]
    ambiguous = len(ranked) > 1 and ranked[1].score == top.score
    selected = {
        **top.as_dict(),
        "ambiguous": ambiguous,
        "selection_note": (
            "[CANDIDATE] empate de score entre >1 run — escolha não conclusiva."
            if ambiguous else
            f"[{'PROVED' if top.confidence == 'HIGH_CONFIDENCE' else 'OBSERVED'}] "
            "run seleccionado por empresa+período+nº beneficiários (total tem peso baixo)."
        ),
    }
    return selected, ranked


# ---------------------------------------------------------------------------
# 7 — como o colaborador aparece em REGU*
# ---------------------------------------------------------------------------

def resolve_employee_identity(connection: Any, params: AnalysisParams, *,
                              regu_rows: list[dict[str, str]], roles: dict[str, str],
                              payroll_pernrs: set[str]) -> dict[str, Any]:
    pf, ef, lf = roles.get("pernr"), roles.get("payee"), roles.get("vendor")
    norm_payroll = {p.lstrip("0") for p in payroll_pernrs}

    def _key(r: dict[str, str]) -> str:
        for rr in ("payment_doc", "payee", "vendor", "customer"):
            v = r.get(roles.get(rr, ""), "")
            if v:
                return f"{rr}:{v}"
        return "?"

    # 1) PERNR directo
    if pf:
        vals = [r.get(pf, "") for r in regu_rows if r.get(pf, "").strip("0")]
        hit = [v for v in vals if v.lstrip("0") in norm_payroll]
        if vals and len(hit) >= max(1, int(0.5 * len(vals))):
            mapping = {_key(r): r.get(pf, "").lstrip("0").zfill(8)
                       for r in regu_rows if r.get(pf, "").strip("0")}
            return {"method": "PERNR_FIELD", "confidence": "PROVED",
                    "field": pf, "mapped": len(mapping),
                    "mapping": mapping, "unresolved": [],
                    "evidence": f"REGUH.{pf} preenchido e {len(hit)}/{len(vals)} valores "
                                f"coincidem com PERNR do payroll."}

    # 2) EMPFG == PERNR
    if ef:
        vals = [r.get(ef, "") for r in regu_rows if r.get(ef, "").strip()]
        norm = [(_digits(v).lstrip("0"), v) for v in vals]
        hit = [nv for nv, _ in norm if nv and nv in norm_payroll]
        if vals and len(hit) >= max(1, int(0.5 * len(vals))):
            mapping = {_key(r): _digits(r.get(ef, "")).lstrip("0").zfill(8)
                       for r in regu_rows if _digits(r.get(ef, ""))}
            return {"method": "EMPFG_IS_PERNR", "confidence": "PROVED",
                    "field": ef, "mapped": len(mapping),
                    "mapping": mapping, "unresolved": [],
                    "evidence": f"REGUH.{ef} numérico e {len(hit)}/{len(vals)} valores "
                                f"coincidem com PERNR do payroll (após remover zeros à esquerda)."}

    # 3) via LIFNR — sem campo transparente PERNR<->LIFNR fiável no R/3 ECC
    if lf:
        lifnrs = sorted({r.get(lf, "") for r in regu_rows if r.get(lf, "").strip("0")})
        return {"method": "LIFNR_UNRESOLVED", "confidence": "HYPOTHESIS",
                "field": lf, "mapped": 0, "mapping": {},
                "unresolved": lifnrs[:200],
                "evidence": "Pagamento via fornecedor (LIFNR). Não há tabela "
                            "transparente PERNR<->LIFNR garantida; relação fica "
                            "por provar caso a caso (LFB1/LFA1/PA0009)."}

    return {"method": "UNKNOWN", "confidence": "UNEXPLAINED", "mapped": 0,
            "mapping": {}, "unresolved": [],
            "evidence": "REGUH sem PERNR, sem EMPFG e sem LIFNR utilizáveis."}


# ---------------------------------------------------------------------------
# 8 / 9 / 10 / 11 — valor RH esperado por colaborador
# ---------------------------------------------------------------------------

def build_payroll_payment_expectations(connection: Any, params: AnalysisParams, *,
                                       run: str,
                                       run_ppoix: list[dict[str, str]] | None = None,
                                       ) -> list[PayrollPaymentExpectation]:
    from .posting_delta_trace import _read_run_ppoix

    px = run_ppoix if run_ppoix is not None else _read_run_ppoix(connection, params, run)
    by_pernr: dict[str, list[dict[str, str]]] = {}
    for r in px:
        by_pernr.setdefault(r.get("PERNR", ""), []).append(r)

    out: list[PayrollPaymentExpectation] = []
    for pernr, rows in sorted(by_pernr.items()):
        n559 = [r for r in rows if r.get("LGART") == BANK_TRANSFER_LGART]
        transferred = [r for r in n559 if str(r.get("TSLIN", "")).strip("0") != ""]
        zero = [r for r in n559 if str(r.get("TSLIN", "")).strip("0") == ""]
        if not transferred:
            continue  # sem transferência bancária corrente -> não entra na reconciliação

        seqnos = sorted({r.get("SEQNO", "") for r in transferred})
        max_seq = seqnos[-1] if seqnos else ""
        line_classes: list[dict[str, Any]] = []
        for r in n559:
            is_tr = str(r.get("TSLIN", "")).strip("0") != ""
            if is_tr and r.get("SEQNO", "") == max_seq:
                cls = "CURRENT_PAYMENT"
            elif is_tr:
                cls = "RETRO_REFERENCE"
            else:
                cls = "PREVIOUS_VERSION"
            line_classes.append({
                "seqno": r.get("SEQNO", ""), "tslin": r.get("TSLIN", ""),
                "postnum": r.get("POSTNUM", ""), "betrg": str(_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))),
                "class": cls,
            })

        current = [r for r in transferred if r.get("SEQNO", "") == max_seq]
        signed = sum((_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")) for r in current), _ZERO)
        expected = -signed  # sinal payroll negativo -> valor a pagar positivo

        if len(seqnos) == 1 and len(current) == 1 and not zero:
            conf, deriv = "PROVED", "/559 corrente (TSLIN!=0), 1 registo, 1 SEQNO"
        elif len(seqnos) == 1:
            conf, deriv = "OBSERVED", f"SUM(/559) corrente (TSLIN!=0), {len(current)} registos, 1 SEQNO"
        else:
            conf, deriv = "CANDIDATE", (f"SUM(/559) do SEQNO mais recente {max_seq} "
                                        f"(há {len(seqnos)} SEQNO transferidos — retro)")

        related: dict[str, str] = {}
        for lg in RELATED_LGARTS:
            s = sum((_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))
                     for r in rows if r.get("LGART") == lg
                     and str(r.get("TSLIN", "")).strip("0") != ""), _ZERO)
            related[lg] = str(s)

        out.append(PayrollPaymentExpectation(
            pernr=pernr, payroll_run=run, seqno=max_seq, lgart=BANK_TRANSFER_LGART,
            expected_payment_amount=expected, currency=params.moeda,
            derivation=deriv, confidence=conf,
            evidence=(f"{len(n559)} linhas /559 ({len(transferred)} transferidas, "
                      f"{len(zero)} com TSLIN=0 => PREVIOUS_VERSION)"),
            related_amounts=related, line_classes=line_classes,
        ))
    return out


# ---------------------------------------------------------------------------
# 12 — pagamentos REGU do run seleccionado
# ---------------------------------------------------------------------------

def read_regu_payments(connection: Any, params: AnalysisParams, *, laufd: str, laufi: str,
                       company: str, schema: dict[str, Any]) -> list[ReguPayment]:
    reguh = schema.get("tables", {}).get("REGUH", {})
    roles = reguh.get("roles", {})
    if not reguh.get("exists"):
        return []
    want = sorted({f for f in roles.values() if f} | {"LAUFD", "LAUFI"})
    try:
        rows = read_table(connection, "REGUH", fields=want,
                          options=opt_and(opt_eq("LAUFD", laufd), opt_eq("LAUFI", laufi)),
                          page_size=200_000).rows
    except (RfcReadError, NoData) as exc:
        logger.warning("REGUH leitura do run falhou: %s", exc)
        return []
    comp_f = roles.get("company", "")
    rows = [r for r in rows if r.get("LAUFD") == laufd and r.get("LAUFI") == laufi
            and (not comp_f or r.get(comp_f) == company)]

    def g(r: dict[str, str], role: str) -> str:
        return r.get(roles.get(role, ""), "")

    out: list[ReguPayment] = []
    for r in rows:
        bank = " ".join(x for x in (g(r, "bank_country"), g(r, "bank_key"),
                                    g(r, "bank_account")) if x).strip()
        out.append(ReguPayment(
            laufd=laufd, laufi=laufi, company=g(r, "company") or company,
            pernr=g(r, "pernr").lstrip("0").zfill(8) if g(r, "pernr").strip("0") else "",
            vendor=g(r, "vendor"), customer=g(r, "customer"), payee=g(r, "payee"),
            payment_doc=g(r, "payment_doc"), acct_doc=g(r, "acct_doc"),
            fiscal_year=g(r, "fiscal_year"), amount=_q(g(r, "amount")),
            currency=g(r, "currency"), payment_method=g(r, "payment_method"),
            bank_reference=bank, name=g(r, "name"), reference=g(r, "reference"),
        ))
        out[-1].beneficiary_key = _beneficiary_key(out[-1], roles)
    return out


def _beneficiary_key(p: ReguPayment, roles: dict[str, str]) -> str:
    if p.pernr:
        return f"pernr:{p.pernr}"
    if p.payee:
        return f"payee:{p.payee}"
    if p.vendor:
        return f"vendor:{p.vendor}"
    if p.customer:
        return f"customer:{p.customer}"
    if p.payment_doc:
        return f"doc:{p.payment_doc}"
    return "?"


def aggregate_regu_by_employee(regu_payments: list[ReguPayment],
                               identity: dict[str, Any]) -> dict[str, Any]:
    mapping = identity.get("mapping", {})
    by_pernr: dict[str, list[ReguPayment]] = {}
    unmatched: list[ReguPayment] = []
    for p in regu_payments:
        pernr = p.pernr or mapping.get(p.beneficiary_key, "")
        if not pernr and identity.get("method") == "EMPFG_IS_PERNR":
            pernr = _digits(p.payee).lstrip("0").zfill(8) if _digits(p.payee) else ""
        if pernr:
            by_pernr.setdefault(pernr, []).append(p)
        else:
            unmatched.append(p)
    agg = {
        pernr: {
            "payments": [p.as_dict() for p in lst],
            "total": str(sum((p.amount for p in lst), _ZERO)),
            "count": len(lst),
            "currencies": sorted({p.currency for p in lst if p.currency}),
            "docs": sorted({p.payment_doc for p in lst if p.payment_doc}),
        }
        for pernr, lst in sorted(by_pernr.items())
    }
    return {"by_pernr": agg, "unmatched_regu": [p.as_dict() for p in unmatched]}


# ---------------------------------------------------------------------------
# 13 / 14 — matching e resultado por colaborador
# ---------------------------------------------------------------------------

_METHOD_BY_IDENTITY = {
    "PERNR_FIELD": "PERNR_DIRECT",
    "EMPFG_IS_PERNR": "EMPFG_PERNR",
    "LIFNR_UNRESOLVED": "VALUE_CANDIDATE",
    "UNKNOWN": "NONE",
}


def match_payroll_to_regu(expectations: list[PayrollPaymentExpectation],
                          regu_agg: dict[str, Any], identity: dict[str, Any],
                          selected_run: dict[str, Any]) -> list[ReconLine]:
    method = _METHOD_BY_IDENTITY.get(identity.get("method", "UNKNOWN"), "NONE")
    laufd, laufi = selected_run.get("laufd", ""), selected_run.get("laufi", "")
    by_pernr = regu_agg.get("by_pernr", {})
    seen_regu: set[str] = set()
    lines: list[ReconLine] = []

    for exp in expectations:
        rec = by_pernr.get(exp.pernr)
        if rec is None:
            lines.append(ReconLine(
                pernr=exp.pernr, rh_expected=exp.expected_payment_amount, regu_paid=None,
                difference=None, status="RH_ONLY", match_method="NONE",
                seqno=exp.seqno, payment_run_date=laufd, payment_run_id=laufi,
                rh_rows=1, regu_rows=0,
                note="RH tem /559 corrente mas não há pagamento REGU para este PERNR.",
            ))
            continue
        seen_regu.add(exp.pernr)
        paid = _q(rec["total"])
        diff = exp.expected_payment_amount - paid
        status = "EXACT_MATCH" if diff == 0 else "DIFFERENCE"
        if method in ("VALUE_CANDIDATE", "NONE"):
            status = "AMBIGUOUS" if diff == 0 else "DIFFERENCE"
        docs = ";".join(rec["docs"][:5])
        lines.append(ReconLine(
            pernr=exp.pernr, rh_expected=exp.expected_payment_amount, regu_paid=paid,
            difference=diff, status=status,
            match_method=method if method != "NONE" else "VALUE_CANDIDATE",
            seqno=exp.seqno, payment_doc=docs, payment_run_date=laufd, payment_run_id=laufi,
            rh_rows=1, regu_rows=rec["count"],
            note=("vários pagamentos REGU para o PERNR" if rec["count"] > 1 else ""),
        ))

    for pernr, rec in by_pernr.items():
        if pernr in seen_regu:
            continue
        lines.append(ReconLine(
            pernr=pernr, rh_expected=None, regu_paid=_q(rec["total"]), difference=None,
            status="REGU_ONLY", match_method=method if method != "NONE" else "VALUE_CANDIDATE",
            payment_doc=";".join(rec["docs"][:5]), payment_run_date=laufd, payment_run_id=laufi,
            rh_rows=0, regu_rows=rec["count"],
            note="Pagamento REGU sem /559 corrente correspondente no payroll.",
        ))
    lines.sort(key=lambda l: (l.status != "DIFFERENCE",
                              -(abs(l.difference) if l.difference is not None else _ZERO)))
    return lines


# ---------------------------------------------------------------------------
# classificação + 427,74
# ---------------------------------------------------------------------------

def classify_reconciliation(recon: PaymentReconciliation) -> dict[str, Any]:
    lines = recon.reconciliation
    ident = recon.employee_identity_mapping or {}
    sel = recon.selected_payment_run or {}
    rh_only = [l for l in lines if l.status == "RH_ONLY"]
    regu_only = [l for l in lines if l.status == "REGU_ONLY"]
    diffs = [l for l in lines if l.status == "DIFFERENCE"]
    ambig = [l for l in lines if l.status in ("AMBIGUOUS", "UNMATCHED")]
    matched = [l for l in lines if l.status in ("EXACT_MATCH", "DIFFERENCE")]

    incomplete = (
        not lines
        or not sel
        or sel.get("confidence") == "LOW_CONFIDENCE"
        or sel.get("ambiguous")
        or ident.get("method") in ("LIFNR_UNRESOLVED", "UNKNOWN")
        or bool(ambig)
    )
    if incomplete:
        klass, tag = "PARTIAL", "[OBSERVED]"
        why = ("Matching incompleto: "
               + ", ".join(x for x in [
                   "run de pagamento não conclusivo" if (not sel or sel.get("confidence") == "LOW_CONFIDENCE" or sel.get("ambiguous")) else "",
                   f"identidade do colaborador = {ident.get('method')}" if ident.get("method") in ("LIFNR_UNRESOLVED", "UNKNOWN") else "",
                   f"{len(ambig)} linhas ambíguas" if ambig else "",
               ] if x) + ".")
    elif not diffs and not rh_only and not regu_only and _q(recon.totals.get("difference", "0")) == 0:
        klass, tag = "EXACT_MATCH", "[PROVED]"
        why = "Todos os colaboradores com EXACT_MATCH e totais fecham a 0,00."
    else:
        klass, tag = "DIFFERENCE", "[PROVED]"
        why = (f"{len(diffs)} colaboradores com diferença, {len(rh_only)} só-RH, "
               f"{len(regu_only)} só-REGU; diferença de total = "
               f"{recon.totals.get('difference', '?')}.")
    return {"classification": klass, "evidence_tag": tag, "rationale": why,
            "counts": {"matched": len(matched), "exact": len(matched) - len(diffs),
                       "difference": len(diffs), "rh_only": len(rh_only),
                       "regu_only": len(regu_only), "ambiguous": len(ambig)}}


def _check_427(recon: PaymentReconciliation) -> None:
    diffs: list[dict[str, Any]] = []
    total_diff = _q(recon.totals.get("difference", "0"))
    if abs(abs(total_diff) - KNOWN_427) < Decimal("0.005"):
        diffs.append({"kind": "sum_rh_minus_sum_regu", "value": str(total_diff),
                      "status": "[CANDIDATE] igual a 427,74 — provar por PERNR antes de concluir."})
    # subconjunto coerente: soma das diferenças por PERNR do mesmo sinal
    per = [(l.pernr, l.difference) for l in recon.reconciliation
           if l.difference not in (None, _ZERO)]
    pos = sum((d for _, d in per if d > 0), _ZERO)
    neg = sum((d for _, d in per if d < 0), _ZERO)
    for label, v in (("soma_diferencas_positivas", pos), ("soma_diferencas_negativas", neg)):
        if abs(abs(v) - KNOWN_427) < Decimal("0.005"):
            diffs.append({"kind": label, "value": str(v),
                          "status": "[CANDIDATE] coincide com 427,74; verificar PERNR a PERNR."})
    if diffs:
        recon.differences.extend(diffs)
        recon.warn("Diferença coincide com 427,74 — marcada CANDIDATE, não provada.")


# ---------------------------------------------------------------------------
# Orquestrador
# ---------------------------------------------------------------------------

def reconcile_payroll_payments(connection: Any, params: AnalysisParams, *, run: str,
                               company: str, period: str,
                               payment_run_date: str | None = None,
                               payment_run_id: str | None = None,
                               ) -> PaymentReconciliation:
    from .posting_delta_trace import _read_run_ppoix

    run = str(run).strip().zfill(10)
    recon = PaymentReconciliation(payroll_run=run, company=company, period=period)

    recon.schema = inspect_regu_schema(connection, params)
    reguh = recon.schema.get("tables", {}).get("REGUH", {})
    if not reguh.get("exists"):
        recon.warn("REGUH não existe/acessível neste sistema — reconciliação impossível.")
        recon.classification = {"classification": "PARTIAL", "evidence_tag": "[OBSERVED]",
                                "rationale": "REGUH indisponível."}
        return recon
    roles = reguh.get("roles", {})

    px = _read_run_ppoix(connection, params, run)
    recon.payroll_expected = build_payroll_payment_expectations(
        connection, params, run=run, run_ppoix=px)
    payroll_pernrs = {e.pernr for e in recon.payroll_expected}
    rh_total = sum((e.expected_payment_amount for e in recon.payroll_expected), _ZERO)
    ref_558_559 = _q(params.valor_rh_referencia) if params.valor_rh_referencia else None

    cands = discover_payment_runs(connection, params, company=company, period=period,
                                  schema=recon.schema)
    selected, ranked = select_payment_run(
        cands, period=period, payroll_employee_count=len(payroll_pernrs),
        payroll_currency=params.moeda, payroll_reference_total=ref_558_559)
    recon.payment_run_candidates = ranked
    if payment_run_date and payment_run_id:
        selected = next(
            (c.as_dict() | {"ambiguous": False,
                            "selection_note": "[OBSERVED] run indicado explicitamente pelo utilizador."}
             for c in ranked if c.laufd == payment_run_date and c.laufi == payment_run_id),
            {"laufd": payment_run_date, "laufi": payment_run_id, "company": company,
             "selection_note": "[OBSERVED] run indicado pelo utilizador (não encontrado nos candidatos)."})
    recon.selected_payment_run = selected

    laufd, laufi = selected.get("laufd", ""), selected.get("laufi", "")
    if not (laufd and laufi):
        recon.warn("Sem run de pagamento identificável na janela — reconciliação parcial.")
        recon.totals = {"rh_expected_total": str(rh_total), "regu_paid_total": "0",
                        "difference": str(rh_total)}
        recon.classification = classify_reconciliation(recon)
        return recon

    regu_rows_raw = _read_reguh_raw(connection, roles, laufd, laufi, company)
    recon.employee_identity_mapping = resolve_employee_identity(
        connection, params, regu_rows=regu_rows_raw, roles=roles,
        payroll_pernrs=payroll_pernrs)

    recon.regu_payments = read_regu_payments(connection, params, laufd=laufd, laufi=laufi,
                                             company=company, schema=recon.schema)
    regu_agg = aggregate_regu_by_employee(recon.regu_payments, recon.employee_identity_mapping)
    recon.reconciliation = match_payroll_to_regu(
        recon.payroll_expected, regu_agg, recon.employee_identity_mapping, selected)

    regu_total = sum((p.amount for p in recon.regu_payments), _ZERO)
    matched_lines = [l for l in recon.reconciliation
                     if l.status in ("EXACT_MATCH", "DIFFERENCE")]
    matched_rh = sum((l.rh_expected or _ZERO for l in matched_lines), _ZERO)
    matched_regu = sum((l.regu_paid or _ZERO for l in matched_lines), _ZERO)
    recon.totals = {
        "rh_employees": len(recon.payroll_expected),
        "regu_beneficiaries": len(regu_agg.get("by_pernr", {})) + len(regu_agg.get("unmatched_regu", [])),
        "matched": len(matched_lines),
        "rh_expected_total": str(rh_total),
        "regu_paid_total": str(regu_total),
        "difference": str(rh_total - regu_total),
        "matched_rh_total": str(matched_rh),
        "matched_regu_total": str(matched_regu),
        "matched_difference": str(matched_rh - matched_regu),
        "unmatched_rh": len([l for l in recon.reconciliation if l.status == "RH_ONLY"]),
        "unmatched_regu": len([l for l in recon.reconciliation if l.status == "REGU_ONLY"]),
    }
    _check_427(recon)
    recon.differences.extend(
        l.as_dict() for l in recon.reconciliation if l.status == "DIFFERENCE")
    recon.classification = classify_reconciliation(recon)
    return recon


def _read_reguh_raw(connection: Any, roles: dict[str, str], laufd: str, laufi: str,
                    company: str) -> list[dict[str, str]]:
    want = sorted({f for f in roles.values() if f} | {"LAUFD", "LAUFI"})
    try:
        rows = read_table(connection, "REGUH", fields=want,
                          options=opt_and(opt_eq("LAUFD", laufd), opt_eq("LAUFI", laufi)),
                          page_size=200_000).rows
    except (RfcReadError, NoData):
        return []
    comp_f = roles.get("company", "")
    return [r for r in rows if r.get("LAUFD") == laufd and r.get("LAUFI") == laufi
            and (not comp_f or r.get(comp_f) == company)]


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_reconciliation_json(recon: PaymentReconciliation, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(recon.as_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info("JSON escrito: %s", path)
    return path


def write_reconciliation_csv(recon: PaymentReconciliation, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["PERNR", "RH_EXPECTED", "REGU_PAID", "DIFFERENCE", "STATUS",
                    "MATCH_METHOD", "SEQNO", "PAYMENT_DOC", "PAYMENT_RUN_DATE", "PAYMENT_RUN_ID"])
        for l in recon.reconciliation:
            d = l.as_dict()
            w.writerow([d["PERNR"], d["RH_EXPECTED"] or "", d["REGU_PAID"] or "",
                        d["DIFFERENCE"] if d["DIFFERENCE"] is not None else "", d["STATUS"],
                        d["MATCH_METHOD"], d["SEQNO"], d["PAYMENT_DOC"],
                        d["PAYMENT_RUN_DATE"], d["PAYMENT_RUN_ID"]])
    logger.info("CSV escrito: %s", path)
    return path


def _fmt(v: Any) -> str:
    if v in (None, ""):
        return "(n/d)"
    try:
        q = Decimal(str(v)).quantize(Decimal("0.01"))
    except Exception:  # noqa: BLE001
        return str(v)
    s = f"{abs(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if q < 0 else ''}{s}"


def print_reconciliation_report(recon: PaymentReconciliation) -> None:
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # pragma: no cover  # noqa: BLE001
        pass
    L = "=" * 60
    sel = recon.selected_payment_run or {}
    t = recon.totals or {}
    print(L)
    print("PAYROLL × REGU RECONCILIATION")
    print(L)
    print(f"Payroll run......... {recon.payroll_run}")
    print(f"Company............. {recon.company}")
    print(f"Period............. {recon.period}")
    print("")
    print(f"Payment run........ {sel.get('laufi', '(n/d)')}  ({sel.get('confidence', '')})")
    print(f"Payment date....... {sel.get('laufd', '(n/d)')}")
    print(f"Identity........... {recon.employee_identity_mapping.get('method', '(n/d)')} "
          f"({recon.employee_identity_mapping.get('confidence', '')})")
    print("")
    print(L)
    print("TOTALS")
    print(L)
    print(f"RH expected......... {_fmt(t.get('rh_expected_total'))}")
    print(f"REGU paid.......... {_fmt(t.get('regu_paid_total'))}")
    print(f"Difference......... {_fmt(t.get('difference'))}")
    print("")
    print(f"Employees RH....... {t.get('rh_employees', 0)}")
    print(f"Beneficiaries REGU. {t.get('regu_beneficiaries', 0)}")
    cc = (recon.classification or {}).get("counts", {})
    print(f"Exact matches...... {cc.get('exact', 0)}")
    print(f"Differences........ {cc.get('difference', 0)}")
    print(f"RH only............ {t.get('unmatched_rh', 0)}")
    print(f"REGU only.......... {t.get('unmatched_regu', 0)}")
    print("")
    print(L)
    print("DIFFERENCES")
    print(L)
    diff_lines = [l for l in recon.reconciliation if l.status == "DIFFERENCE"]
    if not diff_lines:
        print("  (nenhuma)")
    for l in diff_lines[:60]:
        print(f"  {l.pernr:<10} {_fmt(l.rh_expected):>14} {_fmt(l.regu_paid):>14} "
              f"{_fmt(l.difference):>12}  {l.match_method}")
    if recon.differences:
        extra = [d for d in recon.differences if "kind" in d]
        for d in extra:
            print(f"  [427?] {d['kind']} = {d['value']}  {d['status']}")
    print("")
    print(L)
    print("CONCLUSION")
    print(L)
    c = recon.classification or {}
    print(f"RH × REGU: {c.get('classification', '?')}   {c.get('evidence_tag', '')}")
    print(f"  {c.get('rationale', '')}")
    if recon.warnings:
        print("")
        print("AVISOS:")
        for w in recon.warnings:
            print(f"  - {w}")
    print(L)
