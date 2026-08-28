"""Reconciliação, relatório de terminal e ficheiros de output (JSON/CSV)."""

from __future__ import annotations

import csv
import json
import logging
from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AnalysisParams
from .fi_analysis import FIReport
from .models import AnalysisResult, ReconLine
from .payroll_cluster import PayrollClusterReport
from .payroll_posting import PayrollPostingReport
from .payroll_wagetypes import WageTypeLinkReport, WageTypeReport

logger = logging.getLogger(__name__)

LINE = "-" * 64
DBLINE = "=" * 64


def fmt_money(value: Decimal | None) -> str:
    if value is None:
        return "(n/d)"
    quant = value.quantize(Decimal("0.01"))
    sign = "-" if quant < 0 else ""
    digits = f"{abs(quant):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}{digits}"


# ---------------------------------------------------------------------------
# Reconciliação
# ---------------------------------------------------------------------------

def _status(diff: Decimal | None, tol: Decimal) -> str:
    if diff is None:
        return "INDETERMINADO"
    return "OK" if abs(diff) <= tol else "DIVERGENTE"


def effective_payroll_total(payroll: PayrollPostingReport, params: AnalysisParams) -> tuple[Decimal | None, str]:
    """Total do posting RH usado na reconciliação e a que se refere.

    Ordem de preferência: (1) o run primário (evita somar 1298 + 1299 quando
    há duplicação); (2) a empresa pedida; (3) um run cujo valor é exactamente
    a referência FI; (4) a empresa que bate com a referência; (5) o total.
    """
    if not (payroll.resolved and payroll.items):
        return None, ""
    pr = params.primary_run.zfill(10)
    pr_items = [
        i for i in payroll.items
        if i.run_id == pr and (not i.company or i.company == params.empresa)
    ]
    if pr_items:
        tot = sum((i.signed_amount for i in pr_items), Decimal("0"))
        return abs(tot), f"{params.empresa} / run {params.primary_run}"
    if params.empresa in payroll.companies_with_account:
        tot = sum((i.signed_amount for i in payroll.items if i.company == params.empresa), Decimal("0"))
        return abs(tot), params.empresa
    if payroll.match_runs:
        run = payroll.match_runs[0]
        for (r, comp), val in payroll.totals_by_run_company.items():
            if r == run:
                return abs(val), f"{comp} / run {run}"
    if payroll.match_company:
        tot = sum((i.signed_amount for i in payroll.items if i.company == payroll.match_company), Decimal("0"))
        return abs(tot), payroll.match_company
    return abs(payroll.total), "TODAS"


def reconcile(
    params: AnalysisParams,
    payroll: PayrollPostingReport,
    fi: FIReport,
    wt: WageTypeReport | None = None,
    link: WageTypeLinkReport | None = None,
) -> list[ReconLine]:
    tol = params.tolerancia
    rh_total, rh_company = effective_payroll_total(payroll, params)
    fi_total = abs(fi.total) if (fi.resolved and (fi.items or fi.total)) else None
    ref_rh = params.valor_rh_referencia
    ref_fi = params.valor_fi_referencia

    lines: list[ReconLine] = []

    d1 = None if (rh_total is None or fi_total is None) else (rh_total - fi_total)
    lines.append(ReconLine(f"Posting RH x FI  (empresa {rh_company or '?'})", rh_total, fi_total, d1, _status(d1, tol)))

    d2 = None if rh_total is None else (rh_total - ref_rh)
    lines.append(ReconLine("Posting RH x /558+/559 (ref.)", rh_total, ref_rh, d2,
                           "INDETERMINADO" if d2 is None else ("OK" if abs(d2) <= tol else "A EXPLICAR")))

    d3 = None if fi_total is None else (ref_fi - fi_total)
    lines.append(ReconLine("FI informado (ref.) x FI lido", ref_fi, fi_total, d3, _status(d3, tol)))

    d4 = None if fi_total is None else (fi_total - ref_rh)
    lines.append(ReconLine("FI lido x /558+/559 (ref.)", fi_total, ref_rh, d4,
                           "INDETERMINADO" if d4 is None else ("OK" if abs(d4) <= tol else "A EXPLICAR")))

    if wt and wt.resolved and wt.reference_rows:
        wt_total = abs(wt.reference_total)
        d5 = wt_total - ref_rh
        lines.append(ReconLine("PPOIX /558+/559 (todas empresas) x ref.", wt_total, ref_rh, d5, _status(d5, tol)))

    # ---- Fase 2: composição da linha via PPOIX ----
    if link and link.resolved and link.ppoix_rows:
        p_line = abs(link.posting_line_amount)
        p_ppoix = abs(link.ppoix_total)
        p_ref = abs(link.reference_total)
        p_other = abs(link.other_total)

        d6 = p_ppoix - p_line
        lines.append(ReconLine(
            f"PPOIX ligado x linha PPDIT (run {link.run_id})", p_ppoix, p_line, d6, _status(d6, tol)))

        d7 = p_ref - ref_rh
        lines.append(ReconLine(
            "PPOIX /558+/559 (run) x ref. RH", p_ref, ref_rh, d7,
            "OK" if abs(d7) <= tol else "DIVERGENTE"))

        d8 = p_other - params.diferenca_referencia
        lines.append(ReconLine(
            "PPOIX outras rubricas (run) x diferença ref.", p_other, params.diferenca_referencia, d8,
            "OK" if abs(d8) <= tol else "APROX." if abs(d8) < Decimal("500") else "DIVERGENTE"))

    return lines


def overall_status(lines: list[ReconLine]) -> str:
    rel = [l for l in lines if l.label.startswith("Posting RH x FI")]
    if not rel or rel[0].status == "INDETERMINADO":
        return "NÃO FOI POSSÍVEL VALIDAR"
    return "OK" if rel[0].status == "OK" else "DIVERGENTE"


# ---------------------------------------------------------------------------
# Relatório de terminal
# ---------------------------------------------------------------------------

def render_terminal(
    params: AnalysisParams,
    payroll: PayrollPostingReport,
    fi: FIReport,
    recon: list[ReconLine],
    wt: WageTypeReport | None,
    link: WageTypeLinkReport | None = None,
    cluster: PayrollClusterReport | None = None,
    warnings: list[str] | None = None,
) -> str:
    warnings = warnings or []
    out: list[str] = []
    w = out.append

    w(DBLINE)
    w("ANÁLISE PAYROLL -> FI")
    w(DBLINE)
    w("")
    w(f"Empresa.............. {params.empresa}")
    w(f"Período............. {params.periodo_label}")
    w(f"Conta............... {params.conta}  ({params.conta_10})")
    w(f"Moeda............... {params.moeda}")
    w("")
    w("REFERÊNCIA INFORMADA")
    w(f"  Payroll /558 + /559.. {fmt_money(params.valor_rh_referencia)} {params.moeda}")
    w(f"  FI informado......... {fmt_money(params.valor_fi_referencia)} {params.moeda}")
    w(f"  Diferença...........  {fmt_money(params.diferenca_referencia)} {params.moeda}")
    w("")
    w(LINE)
    w("POSTING RUNS PCP0")
    w(LINE)
    for run in params.posting_runs:
        r10 = run.zfill(10)
        docs = payroll.run_to_docs.get(r10, [])
        found = payroll.runs_found.get(r10)
        acc = payroll.runs_with_account.get(r10)
        tag = "documentos: " + (", ".join(docs) if docs else "—")
        if found and acc:
            tag += "  [tem conta-alvo]"
        elif found:
            tag += "  [sem a conta-alvo]"
        else:
            tag += "  [não encontrado em PPDHD]"
        w(f"  {run}  {tag}")
    w("")
    w(LINE)
    w("POSTING RH -> conta em análise  (PPDHD -> PPDIT)")
    w(LINE)
    if payroll.resolved and payroll.items:
        w(f"  {'Run':<12}{'Empresa':<9}{'Conta':<14}{'Valor ' + params.moeda:>18}")
        for (run, comp) in sorted(payroll.totals_by_run_company):
            val = payroll.totals_by_run_company[(run, comp)]
            w(f"  {run:<12}{comp:<9}{params.conta_10:<14}{fmt_money(val):>18}")
        w(f"  {'':<35}{'-' * 18:>18}")
        w(f"  {'TOTAL (com sinal)':<35}{fmt_money(payroll.total):>18}")
        eff, comp = effective_payroll_total(payroll, params)
        w(f"  {'TOTAL p/ reconciliação (' + (comp or '?') + ')':<35}{fmt_money(eff):>18}")
        if payroll.match_runs:
            w(f"  Runs == referência FI ({fmt_money(params.valor_fi_referencia)}): "
              f"{', '.join(payroll.match_runs)}")
        for grp in payroll.duplicate_run_groups:
            w(f"  ATENÇÃO: runs {', '.join(grp)} com valor idêntico (possível duplicação).")
    else:
        w("  (sem dados — ver avisos)")
    w("")
    w("  Todas as contas movimentadas nestes runs (top 15, com sinal):")
    for row in payroll.by_company_account[:15]:
        mark = "  <== ALVO" if row["account"].lstrip("0") == params.conta.lstrip("0") else ""
        w(f"    empresa {row['company']:<6} conta {row['account']:<12} "
          f"linhas {str(row['lines']):>4}  {fmt_money(Decimal(row['signed_sum'])):>18}{mark}")
    w("")
    w(LINE)
    w(f"FI  (origem: {fi.source or 'n/d'}"
      + (f"; empresa {fi.company_used}" if fi.company_used else "") + ")")
    w(LINE)
    if fi.companies_tried:
        w(f"  Empresas testadas em FI: {', '.join(fi.companies_tried)}")
    bpb = fi.bapi_period_balance
    if bpb:
        w(f"  BAPI saldo período {bpb.get('period')}: mov={fmt_money(_dec(bpb.get('period_movement')))}  "
          f"(D={fmt_money(_dec(bpb.get('debit')))}  C={fmt_money(_dec(bpb.get('credit')))})")
    for m in fi.bapi_messages:
        w(f"  BAPI: {m}")
    if fi.resolved and fi.items:
        w(f"  {'Documento':<12}{'Ano':<6}{'Ln':<5}{'Per':<5}{'D/C':<4}{'Valor':>18}")
        for it in fi.items[:200]:
            w(f"  {it.document:<12}{it.fiscal_year:<6}{it.line:<5}{it.period:<5}{it.debit_credit:<4}"
              f"{fmt_money(it.signed_amount):>18}")
        if len(fi.items) > 200:
            w(f"  ... (+{len(fi.items) - 200} linhas — ver CSV)")
        w(f"  {'':<40}{'-' * 18:>18}")
        w(f"  {'TOTAL FI (com sinal)':<40}{fmt_money(fi.total):>18}")
        w(f"  {'  débitos':<40}{fmt_money(fi.total_debit):>18}")
        w(f"  {'  créditos':<40}{fmt_money(fi.total_credit):>18}")
    elif fi.resolved and fi.total:
        w(f"  TOTAL FI (via BAPI): {fmt_money(fi.total)}")
    else:
        w("  (sem partidas FI legíveis — ver avisos)")
    w("")
    if wt is not None:
        w(LINE)
        w("ORIGEM POR RUBRICA (PPOIX) — preparação para /558 e /559")
        w(LINE)
        if wt.resolved:
            for lg, info in wt.by_wage_type.items():
                w(f"  {lg:<8} linhas {str(info['rows']):>5}  {fmt_money(Decimal(info['amount'])):>18}")
            w(f"  {'TOTAL /558+/559 (todas empresas)':<32}{fmt_money(wt.reference_total):>18}")
            if wt.symbolic_accounts:
                w(f"  contas simbólicas (KOMOK): {', '.join(wt.symbolic_accounts)}")
        else:
            w("  (PPOIX não analisado — ver avisos)")
        w("")

    if link is not None:
        _render_wage_link(w, params, link)

    if cluster is not None:
        _render_cluster(w, params, cluster)

    w(LINE)
    w("RECONCILIAÇÃO")
    w(LINE)
    for l in recon:
        w(f"  {l.label:<42} esp={fmt_money(l.left):>15} obt={fmt_money(l.right):>15} "
          f"dif={fmt_money(l.diff):>13} [{l.status}]")
    w("")
    w(f"  Status: {overall_status(recon)}")
    w("")
    if warnings:
        w(LINE)
        w("AVISOS")
        w(LINE)
        for wmsg in warnings:
            w(f"  - {wmsg}")
        w("")
    w(DBLINE)
    return "\n".join(out)


def _dec(v: Any) -> Decimal | None:
    if v in (None, ""):
        return None
    try:
        return Decimal(str(v))
    except Exception:
        return None


def _render_wage_link(w: Any, params: AnalysisParams, link: WageTypeLinkReport) -> None:
    w(LINE)
    w(f"FASE 2 — COMPOSIÇÃO DA LINHA {params.conta}  (run {link.run_id}, empresa {link.company})")
    w(LINE)
    w("  Cadeia: PPOIX.TSLIN = PPDIX.LINUM ; PPDIX.(DOCNUM,DOCLIN) = PPDIT.(DOCNUM,DOCLIN)")
    if not link.resolved:
        w("  (não resolvido — ver avisos)")
        w("")
        return
    w(f"  Linha(s) de posting......... {', '.join(f'{d}/{l}' for d, l in link.posting_doc_lines)}")
    w(f"  Linhas de transferência..... LINUM {', '.join(link.transfer_linums)}")
    w(f"  Contas simbólicas (KOMOK)... {', '.join(link.komok_set) or '—'}")
    w(f"  Linhas PPOIX ligadas........ {link.ppoix_rows}")
    w("")
    w(f"  {'Rubrica':<10}{'Linhas':>8}{'  Montante ' + params.moeda:>20}")
    for lg, info in link.by_wage_type.items():
        tag = "  (ref.)" if lg in params.wage_types_referencia else ""
        w(f"  {lg:<10}{info['rows']:>8}{fmt_money(Decimal(info['amount'])):>20}{tag}")
    w(f"  {'-' * 38:>38}")
    w(f"  {'/558 + /559':<18}{fmt_money(link.reference_total):>20}")
    w(f"  {'outras rubricas':<18}{fmt_money(link.other_total):>20}")
    w(f"  {'TOTAL PPOIX ligado':<18}{fmt_money(link.ppoix_total):>20}")
    w(f"  {'Linha PPDIT (FI)':<18}{fmt_money(link.posting_line_amount):>20}")
    w(f"  {'RESÍDUO (PPOIX - PPDIT)':<18}{fmt_money(link.residual_vs_posting):>20}")
    w("")
    ad = link.account_determination or {}
    w("  DETERMINAÇÃO DE CONTAS (customizing, só leitura):")
    if ad.get("t52ek"):
        koarts = sorted({r.get("KOART", "") for r in ad["t52ek"]})
        w(f"    T52EK  conta(s) simbólica(s) {', '.join(link.komok_set)} -> KOART {', '.join(koarts)}")
    if ad.get("t52el"):
        pos = sorted({r["LGART"] for r in ad["t52el"] if r.get("SIGN") == "+"})
        neg = sorted({r["LGART"] for r in ad["t52el"] if r.get("SIGN") == "-"})
        w(f"    T52EL  rubricas -> {', '.join(link.komok_set)} :  (+) {', '.join(pos) or '—'}")
        w(f"    {'':11}{'':>0}(-) {', '.join(neg) or '—'}")
    if ad.get("t030"):
        hits = sorted({r["BWMOD"] for r in ad["t030"]
                       if params.conta.lstrip('0') in {(r.get('KONTS') or '').lstrip('0'),
                                                       (r.get('KONTH') or '').lstrip('0')}})
        w(f"    T030   KTOPL/{ad.get('ktosl', '?')}/BWMOD -> {params.conta} para: {', '.join(hits) or '—'}")
    if ad.get("conclusion"):
        w(f"    => {ad['conclusion']}")
    if link.link_sample:
        w("")
        w("  PPOIX LINK ANALYSIS (amostra):")
        w(f"    {'PERNR':<10}{'POSTNUM':<9}{'RTLINE':<8}{'LGART':<7}{'KOMOK':<7}"
          f"{'BETRG':>15}  {'TSLIN':<11}{'DOCNUM':<11}{'DOCLIN'}")
        for r in link.link_sample[:25]:
            w(f"    {r['PERNR']:<10}{r['POSTNUM']:<9}{r['RTLINE']:<8}{r['LGART']:<7}{r['KOMOK']:<7}"
              f"{r['BETRG']:>15}  {r['TSLIN']:<11}{r['DOCNUM']:<11}{r['DOCLIN']}")
        if len(link.link_sample) > 25:
            w(f"    ... (+{len(link.link_sample) - 25}; ver CSV)")
    w("")


def _render_cluster(w: Any, params: AnalysisParams, cl: PayrollClusterReport) -> None:
    w(LINE)
    w(f"FASE 3 — RT -> PPOIX -> PPDIT   (run {cl.run_id}, empresa {cl.company})")
    w(LINE)
    if not cl.resolved:
        w("  (não resolvido — ver avisos)")
        w("")
        return
    w(f"  Contexto Payroll..... MOLGA {cl.molga or '?'}  ABKRS {cl.abkrs or '?'}  "
      f"PERMO {cl.permo or '?'}  RELID(PCL2) {cl.relid or '?'}  IN-period {cl.period}")
    w(f"  PERNR no posting..... {cl.pernr_count}  "
      f"(c/ componente do período {len(cl.current_pernr)} / c/ componente retro {len(cl.retro_pernr)})")
    w(f"  RGDIR INPER={cl.period} por FOR-period: "
      + ", ".join(f"{k}:{v}" for k, v in sorted(cl.fpper_distribution.items())))
    w(f"  RGDIR classificação (INPER={cl.period}): "
      + ", ".join(f"{k}={v}" for k, v in sorted(cl.classification_distribution.items())))
    if cl.retro_months_hist:
        w("  Meses de retro por PERNR: "
          + ", ".join(f"{k}m×{v}" for k, v in sorted(cl.retro_months_hist.items())))
    if cl.residual_notes.get("retro_lag_vs_corr"):
        w(f"  {cl.residual_notes['retro_lag_vs_corr']}")
    w("")
    w("  CATÁLOGO DE TABELAS DE RESULTADO DE PAYROLL (descoberta automática):")
    key = {"P2RX_RT", "P2RX_CRT", "P2RX_BT", "P2RX_RT_PERSON", "HRPADNLP_P2RX_RT",
           "HRPY_RGDIR", "HRPY_WPBP"}
    shown = 0
    for t in cl.hrpy_catalog:
        if t.table not in key and not t.accessible:
            continue
        pop = "com dados" if t.populated else ("VAZIA" if t.populated is False else "n/test")
        acc = "SIM" if t.accessible else ("—" if t.accessible is False else "n/test")
        w(f"    {t.table:<20} {t.table_class:<6} campos={t.field_count:<4} "
          f"acessível={acc:<6} {pop:<9} {t.description[:32]}")
        shown += 1
    rest = len(cl.hrpy_catalog) - shown
    if rest > 0:
        w(f"    ... (+{rest} tabelas HRPY_/P2RX_ descobertas; ver hrpy_catalog_*.csv)")
    w("")
    w("  RT (montantes por rubrica — cluster PCL2):")
    a = cl.rt_attempt
    w(f"    {a.function}: {'LIDA' if a.ok else 'NÃO LEGÍVEL (MANUAL_REQUIRED)'} — {a.reason}")
    if a.detail:
        w(f"      {a.detail}")
    if a.ok and a.sample:
        w(f"    amostra RT ({len(a.sample)}): {a.sample[:3]}")
    w("")
    n_retro_only = len([p for p in cl.retro_pernr if p not in cl.current_pernr])
    n_cur_only = len([p for p in cl.current_pernr if p not in cl.retro_pernr])
    n_both = len([p for p in cl.current_pernr if p in cl.retro_pernr])
    mixed = _dec(cl.residual_notes.get("ppoix_ref_mixed_current_and_retro")) or Decimal("0")
    w("  PPOIX /558+/559 por natureza do PERNR (proxy da RT transferida):")
    w(f"    PERNR sem retro (só período).. {fmt_money(cl.ppoix_ref_current_total)}  "
      f"({n_cur_only} PERNR)")
    w(f"    PERNR c/ retro (período+retro) {fmt_money(mixed + cl.ppoix_ref_retro_total)}  "
      f"({n_both + n_retro_only} PERNR)")
    if cl.ppoix_ref_unclassified_total:
        w(f"    PERNR sem entrada RGDIR....... {fmt_money(cl.ppoix_ref_unclassified_total)}")
    w(f"    total /558+/559............... {fmt_money(cl.ppoix_ref_total)}")
    w("")
    n_recalc = sum(1 for p in cl.recalc_pairs
                   if p.get("status") == "RESULT_RECALCULATED" and p.get("contributes_to_run"))
    w(f"  PARES ORIGINAL->RECALCULADO que alimentam o run: {n_recalc}")
    for p in [x for x in cl.recalc_pairs if x.get("contributes_to_run")][:8]:
        w(f"    PERNR {p['pernr']} FPPER {p['fpper']}: IN {'|'.join(p['in_periods'])}  "
          f"orig SEQNR {p['original_seqnr']} -> actual SEQNR {p['current_seqnr']}  [{p['status']}]")
    w("")
    rn = cl.residual_notes
    w("  RESÍDUOS:")
    w(f"    PPOIX vs PPDIT............. {rn.get('ppoix_vs_ppdit', '?')} EUR")
    w(f"    PPOIX /558+/559 vs RH ref.. {rn.get('ppoix_ref_vs_rh', '?')} EUR")
    w(f"    {rn.get('retro_share', '')}")
    w(f"    => {rn.get('explicacao', '')}")
    w("")
    c99 = cl.run_1299_comparison
    if c99.get("ok"):
        w(f"  RUN 1299 (comparação, NÃO somar): {c99.get('classification')}")
        w(f"    valor linha {fmt_money(_dec(c99.get('posting_line_amount')))}  "
          f"PPOIX {fmt_money(_dec(c99.get('ppoix_total')))}  "
          f"PERNR {c99.get('pernr_count')}  "
          f"mesmo conjunto PERNR={c99.get('same_pernr_set')}  "
          f"mesmas rubricas={c99.get('same_lgart_totals')}")
    elif c99:
        w(f"  RUN 1299: não analisável ({c99.get('error', '')[:80]})")
    w("")


# ---------------------------------------------------------------------------
# Output: JSON / CSV
# ---------------------------------------------------------------------------

NEXT_STEPS: list[str] = [
    "RT (cluster PCL2/RP) não é legível por RFC read-only neste sistema "
    "(PYXX_READ_PAYROLL_RESULT/CU_READ_RGDIR -> DA300 «No active nametab»; "
    "HR_GET_PAYROLL_RESULTS não é RFC-enabled; não existe HRPY_RT transparente). "
    "Para fechar os resíduos ao cêntimo é preciso um wrapper Z read-only no SAP "
    "que exponha a RT (RT do período actual + RT dos períodos recalculados).",
    "Alternativa sem desenvolvimento: extrair a RT por relatório standard "
    "(PC00_M99_CWTR / PC_PAYRESULT) e importar o ficheiro para comparação.",
    "Retroactividade é a causa estrutural: 95% dos PERNR do run têm resultado "
    "retro de 05/2026. Comparar RT(FOR 06/2026) com RT(FOR 05/2026) por PERNR "
    "para atribuir os 427,74 e os 265,65.",
    "Run 1299 = repetição do 1298 (mesmos PERNR e mesmos totais por rubrica): "
    "confirmar em PCP0 qual foi efectivamente transferido para FI.",
]


def build_result(
    params: AnalysisParams,
    payroll: PayrollPostingReport,
    fi: FIReport,
    recon: list[ReconLine],
    connection_info: dict[str, Any],
    wt: WageTypeReport | None = None,
    link: WageTypeLinkReport | None = None,
    cluster: PayrollClusterReport | None = None,
    extra_warnings: list[str] | None = None,
) -> AnalysisResult:
    res = AnalysisResult()
    res.parameters = {
        "empresa": params.empresa, "ano": params.ano, "mes": params.mes,
        "conta": params.conta, "conta_10": params.conta_10, "moeda": params.moeda,
        "posting_runs": list(params.posting_runs),
        "wage_types_referencia": list(params.wage_types_referencia),
    }
    res.reference_values = {
        "payroll_558_559": str(params.valor_rh_referencia),
        "fi_informado": str(params.valor_fi_referencia),
        "diferenca": str(params.diferenca_referencia),
    }
    res.connection = connection_info
    res.tables = {name: diag.as_dict() for name, diag in {**payroll.table_diags, **fi.table_diags}.items()}
    res.field_guesses = {
        table: {c: g.as_dict() for c, g in guesses.items()}
        for table, guesses in payroll.field_guesses.items()
    }
    res.posting_runs = [
        {
            "run": run,
            "documentos": payroll.run_to_docs.get(run.zfill(10), []),
            "encontrado": payroll.runs_found.get(run.zfill(10)),
            "tem_conta_alvo": payroll.runs_with_account.get(run.zfill(10)),
        }
        for run in params.posting_runs
    ]
    eff, eff_company = effective_payroll_total(payroll, params)
    res.payroll_posting = {
        "resolved": payroll.resolved,
        "resolved_fields": payroll.resolved_fields,
        "headers": [h.__dict__ for h in payroll.headers],
        "by_company_account": payroll.by_company_account,
        "target_totals_by_run_company": {
            f"{run}|{comp}": str(val) for (run, comp), val in payroll.totals_by_run_company.items()
        },
        "total_signed": str(payroll.total),
        "reconciliation_total": None if eff is None else str(eff),
        "reconciliation_company": eff_company,
        "companies_with_account": payroll.companies_with_account,
        "match_company": payroll.match_company,
        "match_runs": payroll.match_runs,
        "duplicate_run_groups": payroll.duplicate_run_groups,
        "item_count": len(payroll.items),
        "items": [i.as_dict() for i in payroll.items],
    }
    res.fi = {
        "source": fi.source,
        "company_used": fi.company_used,
        "companies_tried": fi.companies_tried,
        "resolved": fi.resolved,
        "resolved_fields": fi.resolved_fields,
        "bapi_period_balance": {k: v for k, v in fi.bapi_period_balance.items() if k != "raw_rows"},
        "bapi_messages": fi.bapi_messages,
        "total_signed": str(fi.total),
        "total_debit": str(fi.total_debit),
        "total_credit": str(fi.total_credit),
        "item_count": len(fi.items),
        "items": [i.as_dict() for i in fi.items],
    }
    if wt is not None:
        res.payroll_posting["wage_types"] = {
            "resolved": wt.resolved,
            "resolved_fields": wt.resolved_fields,
            "by_wage_type": wt.by_wage_type,
            "by_wage_type_symbolic": wt.by_wage_type_symbolic,
            "reference_total_all_companies": str(wt.reference_total),
            "reference_rows": wt.reference_rows,
            "symbolic_accounts": wt.symbolic_accounts,
            "truncated": wt.truncated,
        }
    if link is not None:
        res.wage_link = link.as_dict()
    if cluster is not None:
        res.payroll_cluster = cluster.as_dict()
    res.reconciliation = [l.as_dict() for l in recon]
    res.next_steps = NEXT_STEPS
    for wmsg in [*payroll.warnings, *fi.warnings, *((wt.warnings if wt else [])),
                 *((link.warnings if link else [])), *((cluster.warnings if cluster else [])),
                 *(extra_warnings or [])]:
        res.add_warning(wmsg)
    return res


def write_json(result: AnalysisResult, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result.as_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info("JSON escrito: %s", path)
    return path


def _csv(path: Path, header: list[str], rows: list[list[Any]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(header)
        writer.writerows(rows)
    logger.info("CSV escrito: %s", path)
    return path


def write_csv_fi(fi: FIReport, path: Path) -> Path:
    return _csv(
        path,
        ["source", "document", "fiscal_year", "line", "period", "posting_date", "account",
         "company", "currency", "debit_credit", "amount", "signed_amount", "doc_type", "reference", "text"],
        [[i.source, i.document, i.fiscal_year, i.line, i.period, i.posting_date, i.account, i.company,
          i.currency, i.debit_credit, f"{i.amount}", f"{i.signed_amount}", i.doc_type, i.reference, i.text]
         for i in fi.items],
    )


def write_csv_payroll(payroll: PayrollPostingReport, path: Path) -> Path:
    return _csv(
        path,
        ["run_id", "doc_number", "line", "account", "company", "currency", "debit_credit",
         "amount", "signed_amount"],
        [[i.run_id, i.doc_number, i.line, i.account, i.company, i.currency, i.debit_credit,
          f"{i.amount}", f"{i.signed_amount}"] for i in payroll.all_items],
    )


def write_csv_rgdir(cluster: PayrollClusterReport, path: Path) -> Path:
    """RGDIR completo (todos os SEQNR) dos PERNR do posting, com classificação."""
    rows = []
    for tl in cluster.timelines:
        for e in tl.entries:
            rows.append([e.pernr, e.seqnr, e.abkrs, e.fpper, e.fpbeg, e.fpend, e.inper, e.ipend,
                         e.srtza, e.payty, e.payid, e.void, e.reversal, e.outofseq, e.ocrsn,
                         e.rundt, e.classify()])
    if not rows:  # timelines não construídas -> usa as entradas do run
        rows = [[e.pernr, e.seqnr, e.abkrs, e.fpper, e.fpbeg, e.fpend, e.inper, e.ipend,
                 e.srtza, e.payty, e.payid, e.void, e.reversal, e.outofseq, e.ocrsn,
                 e.rundt, e.classify()] for e in cluster.rgdir_for_inper]
    return _csv(
        path,
        ["pernr", "seqnr", "abkrs", "fpper", "fpbeg", "fpend", "inper", "ipend",
         "srtza", "payty", "payid", "void", "reversal", "outofseq", "ocrsn", "rundt", "classification"],
        rows,
    )


def write_csv_timeline_pairs(cluster: PayrollClusterReport, path: Path) -> Path:
    """Pares original->recalculado por (PERNR, FPPER)."""
    return _csv(
        path,
        ["pernr", "fpper", "in_periods", "recalc_count", "original_seqnr", "original_inper",
         "current_seqnr", "current_inper", "status", "contributes_to_run"],
        [[p["pernr"], p["fpper"], "|".join(p["in_periods"]), p["recalc_count"],
          p["original_seqnr"], p["original_inper"], p["current_seqnr"], p["current_inper"],
          p["status"], p.get("contributes_to_run", "")]
         for p in cluster.recalc_pairs],
    )


def write_csv_ppoix_rgdir_view(cluster: PayrollClusterReport, path: Path) -> Path:
    """Vista única: PPOIX (por rubrica) + factos RGDIR do run, por PERNR."""
    return _csv(
        path,
        ["pernr", "ppoix_558", "ppoix_559", "ppoix_561", "ppoix_563", "ppoix_0029",
         "ppoix_ref_558_559", "rgdir_run_fppers", "retro_months", "has_current",
         "current_seqnr", "classes"],
        [[r["pernr"], r["ppoix_558"], r["ppoix_559"], r["ppoix_561"], r["ppoix_563"],
          r["ppoix_0029"], r["ppoix_ref"], "|".join(r["rgdir_run_fppers"]), r["retro_months"],
          r["has_current"], r["current_seqnr"], "|".join(r["classes"])]
         for r in cluster.ppoix_rgdir_view],
    )


def write_csv_hrpy_catalog(cluster: PayrollClusterReport, path: Path) -> Path:
    return _csv(
        path,
        ["table", "class", "field_count", "exists", "accessible", "populated", "description", "note"],
        [[t.table, t.table_class, t.field_count, t.exists, t.accessible,
          "" if t.populated is None else t.populated, t.description, t.note]
         for t in cluster.hrpy_catalog],
    )


def write_csv_wage_link(link: WageTypeLinkReport, path: Path) -> Path:
    """Amostra PPOIX LINK ANALYSIS + agregado por rubrica."""
    rows = [["#SAMPLE", r["PERNR"], r["POSTNUM"], r["RTLINE"], r["LGART"], r["KOMOK"],
             r["BETRG"], r["TSLIN"], r["DOCNUM"], r["DOCLIN"]] for r in link.link_sample]
    for lg, info in link.by_wage_type.items():
        rows.append(["#BY_LGART", lg, "", "", "", "", info["amount"], str(info["rows"]), "", ""])
    return _csv(
        path,
        ["kind", "pernr_or_lgart", "postnum", "rtline", "lgart", "komok",
         "betrg_or_amount", "tslin_or_rows", "docnum", "doclin"],
        rows,
    )
