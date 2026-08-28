"""Testes da reconciliação de totais e do cálculo de diferença."""

from decimal import Decimal

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.fi_analysis import FIReport
from sap_payroll_analysis.models import PostingItem
from sap_payroll_analysis.payroll_posting import PayrollPostingReport
from sap_payroll_analysis.report import effective_payroll_total, overall_status, reconcile


def _payroll(total: str, company: str = DEFAULTS.empresa) -> PayrollPostingReport:
    r = PayrollPostingReport()
    r.resolved = True
    r.total = Decimal(total)
    r.items = [PostingItem(run_id="0000001296", company=company, signed_amount=Decimal(total))]
    r.totals_by_run_company = {("0000001296", company): Decimal(total)}
    r.companies_with_account = [company]
    return r


def _fi(total: str) -> FIReport:
    r = FIReport()
    r.resolved = True
    r.source = "BSIS/BSAS"
    r.total = Decimal(total)
    r.items = [object()]
    return r


def _line(lines, prefix):
    return next(l for l in lines if l.label.startswith(prefix))


def test_effective_total_prefers_requested_company():
    p = _payroll("-727258.35", company=DEFAULTS.empresa)
    total, comp = effective_payroll_total(p, DEFAULTS)
    assert total == Decimal("727258.35")
    assert comp == DEFAULTS.empresa


def test_effective_total_falls_back_to_match_company():
    other = "9990"  # empresa diferente da pedida
    p = _payroll("-727258.35", company=other)
    p.match_company = other
    total, comp = effective_payroll_total(p, DEFAULTS)
    assert total == Decimal("727258.35")
    assert comp == other


def test_reconcile_matches_reference():
    lines = reconcile(DEFAULTS, _payroll("-727258.35"), _fi("-727258.35"))
    rh_fi = _line(lines, "Posting RH x FI")
    assert rh_fi.diff == Decimal("0.00")
    assert rh_fi.status == "OK"

    to_explain = _line(lines, "Posting RH x /558+/559")
    assert to_explain.diff == Decimal("3211.71")
    assert to_explain.status == "A EXPLICAR"


def test_reconcile_divergent_posting_vs_fi():
    lines = reconcile(DEFAULTS, _payroll("-727000.00"), _fi("-727258.35"))
    assert _line(lines, "Posting RH x FI").status == "DIVERGENTE"
    assert overall_status(lines) == "DIVERGENTE"


def test_reconcile_indeterminate_when_no_fi():
    lines = reconcile(DEFAULTS, _payroll("-727258.35"), FIReport())
    assert _line(lines, "Posting RH x FI").status == "INDETERMINADO"
    assert overall_status(lines) == "NÃO FOI POSSÍVEL VALIDAR"
