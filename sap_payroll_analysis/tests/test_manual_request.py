"""Shortlist mínima para consulta manual (PC00_M99_CWTR) — sem SAP.

Usa um JSON de cluster sintético (o mesmo formato de
`output/payroll_cluster_analysis_*.json`).
"""

import json
from decimal import Decimal
from pathlib import Path

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.manual_request import (
    ManualRtCase,
    build_manual_rt_shortlist,
)

PERIOD = "202606"


def _view(pernr, w558="0", w559="0", w561="0", w563="0", w0029="0", cur_seqnr="0300"):
    return {"pernr": pernr, "ppoix_558": w558, "ppoix_559": w559, "ppoix_561": w561,
            "ppoix_563": w563, "ppoix_0029": w0029, "ppoix_ref": str(Decimal(w558) + Decimal(w559)),
            "rgdir_run_fppers": [], "rgdir_run_seqnrs": [], "retro_months": 0,
            "has_current": True, "current_seqnr": cur_seqnr, "classes": []}


def _pair(pernr, fpper, orig, orig_inper, cur, cur_inper):
    return {"pernr": pernr, "fpper": fpper, "in_periods": [orig_inper, cur_inper],
            "recalc_count": 2, "original_seqnr": orig, "original_inper": orig_inper,
            "current_seqnr": cur, "current_inper": cur_inper, "status": "RESULT_RECALCULATED",
            "contributes_to_run": True}


def _fixture(path: Path) -> None:
    cluster = {
        "run_id": "0000001298", "company": "1010", "period": PERIOD,
        "ppoix_ref_total": "-724474.38",
        "residual_notes": {"ppoix_vs_ppdit": "265.65"},
        "ppoix_rgdir_view": [
            # B / prioridade 1 — correcção real (FPPER 202603 recalculado em 202606)
            _view("00000001", w559="-5000.00", cur_seqnr="0110"),
            # B / prioridade 2 — carry-forward retro presente
            _view("00000002", w559="-2000.00", w561="120.00", w563="-120.00", cur_seqnr="0210"),
            # E — recálculo de rotina (1 mês), sem evidência
            _view("00000003", w559="-1500.00", cur_seqnr="0310"),
            # A — retro de processamento sem par de versões
            _view("00000004", w559="-800.00", cur_seqnr="0410"),
        ],
        "recalc_pairs": [
            _pair("00000001", "202603", "0101", "202603", "0108", "202606"),
            _pair("00000001", "202605", "0105", "202605", "0109", "202606"),
            _pair("00000002", "202605", "0205", "202605", "0209", "202606"),
            _pair("00000003", "202605", "0305", "202605", "0309", "202606"),
            # 00000004 -> sem par válido (old == new)
            {"pernr": "00000004", "fpper": "202605", "in_periods": ["202605"],
             "recalc_count": 1, "original_seqnr": "0405", "original_inper": "202605",
             "current_seqnr": "0405", "current_inper": "202605", "status": "RESULT_UNCHANGED",
             "contributes_to_run": True},
        ],
    }
    path.write_text(json.dumps(cluster), encoding="utf-8")


def test_manual_request_categories_and_priority(tmp_path: Path):
    _fixture(tmp_path / "payroll_cluster_analysis_x.json")
    res = build_manual_rt_shortlist(DEFAULTS, output_dir=tmp_path, write_files=True)

    assert res.source == "json"
    assert res.total_pernrs == 4
    assert res.category_b_count == 2          # 00000001, 00000002
    assert res.category_e_count == 1          # 00000003 (rotina)
    assert res.category_a_count == 1          # 00000004 (sem par)
    assert res.category_c_count == 0 and res.category_d_count == 0

    # 00000001 tem 2 FPPER recalculados -> 2 casos; 00000002 -> 1 caso
    assert len(res.cases) == 3
    assert sorted({c.pernr for c in res.cases}) == ["00000001", "00000002"]

    p1 = [c for c in res.cases if c.priority == 1]
    p2 = [c for c in res.cases if c.priority == 2]
    assert {c.pernr for c in p1} == {"00000001"}   # correcção >=2 meses
    assert {c.pernr for c in p2} == {"00000002"}   # carry-forward

    for c in res.cases:
        assert c.old_seqnr and c.new_seqnr and c.old_seqnr != c.new_seqnr
        assert c.new_inper == PERIOD
        assert c.fpper < PERIOD

    # ordenação: prioridade asc, depois |/558+/559| desc
    assert [c.priority for c in res.cases] == sorted(c.priority for c in res.cases)

    assert res.pernrs_manual == ["00000001", "00000002"]
    assert res.gap_ppoix_vs_rh == Decimal("-724474.38").copy_abs() - DEFAULTS.valor_rh_referencia

    # ficheiros gerados
    assert res.csv_path.exists() and res.txt_path.exists()
    csv_txt = res.csv_path.read_text(encoding="utf-8-sig")
    assert "PERNR;FPPER;OLD_SEQNR" in csv_txt
    assert "00000001;202603;0101;202603" in csv_txt
    req = res.txt_path.read_text(encoding="utf-8")
    assert "Preciso da RT do SEQNR 0101 e do SEQNR 0108" in req
    assert "CASO 1" in req and "CASO 3" in req


def test_manual_request_needs_data():
    import pytest

    with pytest.raises(FileNotFoundError):
        build_manual_rt_shortlist(DEFAULTS, output_dir=Path("does-not-exist-xyz"),
                                  write_files=False)


def test_manual_rt_case_percentages():
    c = ManualRtCase(pernr="1", fpper="202605", old_seqnr="1", new_seqnr="2",
                     ppoix_558=Decimal("-10.00"), ppoix_559=Decimal("-417.74"))
    assert c.ppoix_ref == Decimal("-427.74")
    assert c.pct_of_gap(Decimal("427.74")) == "100%"
