"""Modo dedicado da Fase 3 (`--explain-cluster` / `--rt-link-diagnostic` /
`--payroll-timeline`).

Reutiliza as Fases 1 e 2 e corre só a análise de cluster. Sem tocar em FI.
Tudo automático — nenhum PERNR é pedido ao utilizador.
"""

from __future__ import annotations

import logging
from decimal import Decimal
from typing import Any

from .config import AnalysisParams
from .payroll_cluster import analyse_cluster, build_timeline
from .payroll_posting import analyze as analyze_payroll
from .payroll_wagetypes import link_wage_types_to_posting_line
from .report import _render_cluster, fmt_money

logger = logging.getLogger(__name__)


def run_cluster_only(
    connection: Any,
    params: AnalysisParams,
    *,
    pernr: str | None = None,
    rt_link_diagnostic: bool = False,
    payroll_timeline: bool = False,
) -> None:
    payroll = analyze_payroll(connection, params)
    link = link_wage_types_to_posting_line(connection, params, payroll)
    if not link.resolved:
        print("Fase 2 não resolvida — impossível correr a fase 3.")
        for w in link.warnings:
            print("  -", w)
        return

    cluster = analyse_cluster(connection, params, payroll, link)

    out: list[str] = []
    _render_cluster(out.append, params, cluster)
    print("\n".join(out))

    if rt_link_diagnostic:
        _print_rt_link_diag(params, cluster, pernr)
    if payroll_timeline:
        _print_timeline(cluster, pernr)


def _print_rt_link_diag(params: AnalysisParams, cluster: Any, pernr: str | None) -> None:
    rows = cluster.per_pernr_diag
    if pernr:
        p10 = pernr.strip().zfill(8)
        rows = [r for r in rows if r["pernr"] in {pernr.strip(), p10}]
    ordered = sorted(rows, key=lambda d: abs(Decimal(d["ppoix_ref"] or "0")), reverse=True)
    cap = len(ordered) if pernr else 40
    print("-" * 72)
    print("RT <-> PPOIX LINK DIAGNOSTIC (por PERNR)  [RT n/d neste sistema]")
    print("-" * 72)
    print(f"  {'PERNR':<10}{'PPOIX /558+/559':>16}  {'RT':>6}  {'retro m':>7}  "
          f"{'FOR-periods do run':<26}{'classe'}")
    tot = Decimal("0")
    for i, r in enumerate(ordered):
        amt = Decimal(r["ppoix_ref"] or "0")
        tot += amt
        if i < cap:
            klass = "RETRO+CUR" if (r["retro"] and r["current"]) else \
                    ("RETRO" if r["retro"] else ("CURRENT" if r["current"] else "SEM-RGDIR"))
            fp = "|".join(r["rgdir_run_fppers"])[:25]
            print(f"  {r['pernr']:<10}{fmt_money(amt):>16}  {'n/d':>6}  "
                  f"{r['retro_months']:>7}  {fp:<26}{klass}")
    if len(ordered) > cap:
        print(f"  ... (+{len(ordered) - cap} PERNR; ver ppoix_rgdir_view_*.csv)")
    print("-" * 72)
    print(f"  TOTAL /558+/559 (PPOIX) = {fmt_money(tot)}   |   RT = n/d (cluster PCL2)")


def _print_timeline(cluster: Any, pernr: str | None) -> None:
    tls = cluster.timelines
    if pernr:
        p10 = pernr.strip().zfill(8)
        tls = [t for t in tls if t.pernr in {pernr.strip(), p10}]
    else:
        tls = tls[:3]  # amostra
    for tl in tls:
        print("-" * 72)
        print(f"TIMELINE DE PAYROLL — PERNR {tl.pernr}")
        print("-" * 72)
        print(f"  {'SEQNR':<7}{'FPPER':<8}{'INPER':<8}{'SRTZA':<6}{'RUNDT':<10}classificação")
        for e in tl.entries:
            print(f"  {e.seqnr:<7}{e.fpper:<8}{e.inper:<8}{e.srtza:<6}{e.rundt:<10}{e.classify()}")
        print("  PARES (PERNR+FPPER):")
        for p in tl.pairs:
            if p.status == "RESULT_UNCHANGED":
                continue
            d = p.as_dict()
            print(f"    FPPER {d['fpper']}: IN {'|'.join(d['in_periods'])}  "
                  f"orig SEQNR {d['original_seqnr']} -> actual SEQNR {d['current_seqnr']}  [{d['status']}]")
