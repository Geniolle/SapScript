"""Orquestração da análise completa Payroll -> FI."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from .config import AnalysisParams
from .fi_analysis import analyze as analyze_fi
from .models import AnalysisResult
from .payroll_cluster import analyse_cluster
from .payroll_posting import analyze as analyze_payroll
from .payroll_wagetypes import link_wage_types_to_posting_line, probe_reference_wage_types
from .report import (
    build_result,
    reconcile,
    render_terminal,
    write_csv_fi,
    write_csv_hrpy_catalog,
    write_csv_payroll,
    write_csv_ppoix_rgdir_view,
    write_csv_rgdir,
    write_csv_timeline_pairs,
    write_csv_wage_link,
    write_json,
)

logger = logging.getLogger(__name__)


def run_analysis(
    connection: Any,
    params: AnalysisParams,
    connection_info: dict[str, Any],
    output_dir: Path | None = None,
    write_files: bool = True,
    with_cluster: bool = True,
) -> AnalysisResult:
    logger.info("Análise: empresa=%s período=%s conta=%s",
                params.empresa, params.periodo_label, params.conta)

    payroll = analyze_payroll(connection, params)
    fi = analyze_fi(connection, params, extra_companies=payroll.companies_with_account)
    wt = probe_reference_wage_types(connection, params)
    link = link_wage_types_to_posting_line(connection, params, payroll)
    cluster = analyse_cluster(connection, params, payroll, link) if with_cluster else None
    recon = reconcile(params, payroll, fi, wt, link)

    text = render_terminal(
        params, payroll, fi, recon, wt, link, cluster,
        [*payroll.warnings, *fi.warnings, *wt.warnings, *link.warnings,
         *(cluster.warnings if cluster else [])],
    )
    print(text)

    result = build_result(params, payroll, fi, recon, connection_info, wt, link, cluster)

    if write_files:
        out = output_dir or (Path.cwd() / "output")
        slug = f"{params.empresa}_{params.ano}_{params.mes:02d}_{params.conta}"
        run_slug = f"{params.empresa}_{params.ano}_{params.mes:02d}_run{link.run_id}"
        write_json(result, out / f"payroll_fi_{slug}.json")
        write_csv_fi(fi, out / f"fi_items_{slug}.csv")
        write_csv_payroll(payroll, out / f"payroll_posting_items_{slug}.csv")
        write_csv_wage_link(link, out / f"wage_link_{slug}_run{link.run_id}.csv")
        if cluster is not None and cluster.resolved:
            write_json(_wrap(result.payroll_cluster),
                       out / f"payroll_cluster_analysis_{run_slug}.json")
            write_csv_rgdir(cluster, out / f"rgdir_{run_slug}.csv")
            write_csv_timeline_pairs(cluster, out / f"rgdir_pairs_{run_slug}.csv")
            write_csv_ppoix_rgdir_view(cluster, out / f"ppoix_rgdir_view_{run_slug}.csv")
            write_csv_hrpy_catalog(cluster, out / f"hrpy_catalog_{run_slug}.csv")
        print(f"\nFicheiros de diagnóstico em: {out}")

    return result


class _wrap:
    """Adapta um dict ao `.as_dict()` esperado por `write_json`."""

    def __init__(self, payload: dict[str, Any]) -> None:
        self._payload = payload

    def as_dict(self) -> dict[str, Any]:
        return self._payload
