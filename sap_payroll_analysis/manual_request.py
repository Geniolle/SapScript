"""Shortlist mínima para consulta manual no relatório PC00_M99_CWTR.

A RT efectiva está no cluster PCL2 e não é legível por RFC neste sistema
(comprovado pela descoberta DDIC). Em vez de pedir os 324 PERNR, este módulo
usa apenas os dados **já obtidos automaticamente** (HRPY_RGDIR + PPOIX) para
produzir a lista dos PERNR/SEQNR onde há **evidência estrutural de recálculo
real** — os únicos que vale a pena extrair à mão para fechar os resíduos.

Não consulta o cluster. `analyse_cluster(..., try_rt=False)` lê só tabelas
transparentes (RGDIR, PPOIX, PPDIT, …).
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AnalysisParams
from .payroll_cluster import _months_between  # helper YYYYMM -> nº meses

logger = logging.getLogger(__name__)

RT_FIELDS = ("LGART", "BETRG", "ANZHL", "RTE", "APZNR", "C1ZNR", "V0ZNR", "ALZNR")
PRIORITY_WAGE_TYPES = ("/558", "/559", "/561", "/563", "0029")


# ---------------------------------------------------------------------------

@dataclass
class ManualRtCase:
    pernr: str
    fpper: str
    old_seqnr: str = ""
    old_inper: str = ""
    old_srtza: str = ""
    new_seqnr: str = ""
    new_inper: str = ""
    new_srtza: str = ""
    ppoix_558: Decimal = Decimal("0")
    ppoix_559: Decimal = Decimal("0")
    ppoix_561: Decimal = Decimal("0")
    ppoix_563: Decimal = Decimal("0")
    ppoix_0029: Decimal = Decimal("0")
    retro_months: int = 0
    priority: int = 4
    reason: str = ""
    june_seqnr: str = ""          # SEQNR do próprio período (contexto)

    @property
    def ppoix_ref(self) -> Decimal:
        return self.ppoix_558 + self.ppoix_559

    @property
    def ppoix_total(self) -> Decimal:
        return self.ppoix_558 + self.ppoix_559 + self.ppoix_561 + self.ppoix_563 + self.ppoix_0029

    def pct_of_gap(self, gap: Decimal) -> str:
        """|/558+/559| deste PERNR em % do resíduo a fechar (>100% = o PERNR é
        grande o suficiente para, sozinho, esconder um erro do tamanho do
        resíduo — NÃO significa que o explica)."""
        if not gap:
            return ""
        return f"{abs(self.ppoix_ref) / abs(gap) * 100:.0f}%"

    def as_row(self, gap: Decimal) -> list[str]:
        return [
            self.pernr, self.fpper,
            self.old_seqnr, self.old_inper, self.old_srtza,
            self.new_seqnr, self.new_inper, self.new_srtza,
            f"{self.ppoix_558}", f"{self.ppoix_559}", f"{self.ppoix_561}",
            f"{self.ppoix_563}", f"{self.ppoix_0029}", f"{self.ppoix_total}",
            str(self.retro_months), str(self.priority),
            f"{self.reason} | |558+559|/residuo={self.pct_of_gap(gap)}",
        ]


@dataclass
class ManualRtResult:
    run_id: str = ""
    company: str = ""
    period: str = ""
    account: str = ""
    gap_ppoix_vs_rh: Decimal = Decimal("0")        # 427.74
    gap_ppoix_vs_ppdit: Decimal = Decimal("0")     # 265.65
    ppoix_ref_total: Decimal = Decimal("0")
    total_pernrs: int = 0
    cases: list[ManualRtCase] = field(default_factory=list)     # só categoria B
    category_counts: dict[str, int] = field(default_factory=dict)
    source: str = ""                                            # "rfc" | "json"
    warnings: list[str] = field(default_factory=list)
    csv_path: Path | None = None
    txt_path: Path | None = None

    @property
    def category_a_count(self) -> int:
        return self.category_counts.get("A", 0)

    @property
    def category_b_count(self) -> int:
        return self.category_counts.get("B", 0)

    @property
    def category_c_count(self) -> int:
        return self.category_counts.get("C", 0)

    @property
    def category_d_count(self) -> int:
        return self.category_counts.get("D", 0)

    @property
    def category_e_count(self) -> int:
        return self.category_counts.get("E", 0)

    @property
    def pernrs_manual(self) -> list[str]:
        return sorted({c.pernr for c in self.cases if c.priority <= 2})

    @property
    def all_shortlist_pernrs(self) -> list[str]:
        return sorted({c.pernr for c in self.cases})


# ---------------------------------------------------------------------------
# Construção
# ---------------------------------------------------------------------------

def _cases_for_pernr(pernr: str, recalc_pairs: list[dict[str, Any]], ppoix: dict[str, Decimal],
                     period: str, june_seqnr: str) -> list[ManualRtCase]:
    """Um `ManualRtCase` por FPPER com recálculo real (categoria B).

    `recalc_pairs` já vem filtrado: `fpper < period`, old/new SEQNR presentes
    e diferentes. Devolve [] se não há evidência de recálculo.
    """
    real = [p for p in recalc_pairs
            if p["fpper"] < period and p["old_seqnr"] and p["new_seqnr"]
            and p["old_seqnr"] != p["new_seqnr"]]
    if not real:
        return []
    real.sort(key=lambda p: p["fpper"])

    months_late = max(_months_between(p["fpper"], period) for p in real)
    has_cf = any(ppoix.get(w) for w in ("/561", "/563", "0029"))

    # Categoria B só se houver evidência de recálculo REAL (não a finalização
    # mensal de rotina): correcção de período >=2 meses OU rubricas de
    # carry-forward retro no PPOIX. Caso contrário -> ambíguo (categoria E),
    # devolvido como [] aqui e contabilizado pelo chamador.
    if months_late >= 2:
        prio, why = 1, f"correcção real: FPPER recalculado {months_late} meses depois, em {period}"
    elif has_cf:
        prio, why = 2, "PPOIX c/ rubricas de carry-forward retro (/561 //563 /0029)"
    else:
        return []

    out: list[ManualRtCase] = []
    for p in real:
        out.append(ManualRtCase(
            pernr=pernr, fpper=p["fpper"],
            old_seqnr=p["old_seqnr"], old_inper=p["old_inper"], old_srtza=p["old_srtza"],
            new_seqnr=p["new_seqnr"], new_inper=p["new_inper"], new_srtza=p["new_srtza"],
            ppoix_558=ppoix.get("/558", Decimal("0")), ppoix_559=ppoix.get("/559", Decimal("0")),
            ppoix_561=ppoix.get("/561", Decimal("0")), ppoix_563=ppoix.get("/563", Decimal("0")),
            ppoix_0029=ppoix.get("0029", Decimal("0")),
            retro_months=len(real), priority=prio, reason=why, june_seqnr=june_seqnr,
        ))
    return out


def _cases_from_cluster(cluster: Any) -> ManualRtResult:
    period = cluster.period
    res = ManualRtResult(run_id=cluster.run_id, company=cluster.company, period=period)

    view_by_pernr = {r["pernr"]: r for r in cluster.ppoix_rgdir_view}
    cat = {"B": 0, "A": 0, "C": 0, "D": 0, "E": 0}

    for tl in cluster.timelines:
        pn = tl.pernr
        view = view_by_pernr.get(pn, {})
        ppoix = {w: _d(view.get(f"ppoix_{w.strip('/')}", "0")) for w in PRIORITY_WAGE_TYPES}

        run_entries = [e for e in tl.entries if e.inper == period]
        if not run_entries:
            cat["E"] += 1
            continue
        if any(e.is_void for e in run_entries):
            cat["D"] += 1
            continue
        if any(e.is_offcycle for e in run_entries):
            cat["C"] += 1
            continue

        june = next((e for e in run_entries if e.fpper == period), None)
        june_seqnr = june.seqnr if june else ""
        run_fppers = {e.fpper for e in run_entries}

        # par por FPPER: ORIGINAL (INPER=FPPER) -> ACTUAL/definitivo (SRTZA=A).
        # Só FPPER que o run transferiu (tem entrada com INPER == período).
        pairs: list[dict[str, Any]] = []
        for pr in tl.pairs:
            if pr.fpper not in run_fppers:
                continue
            o, c = pr.original, pr.current
            if not (o and c):
                continue
            pairs.append({
                "fpper": pr.fpper,
                "old_seqnr": o.seqnr, "old_inper": o.inper, "old_srtza": o.srtza,
                "new_seqnr": c.seqnr, "new_inper": c.inper, "new_srtza": c.srtza,
            })

        cases = _cases_for_pernr(pn, pairs, ppoix, period, june_seqnr)
        if cases:
            cat["B"] += 1
            res.cases.extend(cases)
        elif _has_recalc_pair(pairs, period):
            cat["E"] += 1          # recálculo de rotina — não confirmável sem RT
        else:
            cat["A"] += 1          # retro de processamento sem par de versões

    res.total_pernrs = len(cluster.timelines)
    res.category_counts = cat
    return res


def _has_recalc_pair(pairs: list[dict[str, Any]], period: str) -> bool:
    return any(p["fpper"] < period and p["old_seqnr"] and p["new_seqnr"]
              and p["old_seqnr"] != p["new_seqnr"] for p in pairs)




def _cases_from_json(data: dict[str, Any]) -> ManualRtResult:
    # o ficheiro payroll_cluster_analysis_*.json tem o cluster no topo;
    # o payroll_fi_*.json tem-no em ["payroll_cluster"].
    pc = data.get("payroll_cluster") or data
    if "ppoix_rgdir_view" not in pc and "recalc_pairs" not in pc:
        raise ValueError("JSON sem dados de cluster (ppoix_rgdir_view / recalc_pairs)")
    period = pc.get("period", "")
    res = ManualRtResult(run_id=pc.get("run_id", ""), company=pc.get("company", ""),
                         period=period, source="json")

    view_rows = pc.get("ppoix_rgdir_view", [])
    view_by_pernr = {r["pernr"]: r for r in view_rows}
    pairs_by_pernr: dict[str, list[dict[str, Any]]] = {}
    for p in pc.get("recalc_pairs", []):
        pairs_by_pernr.setdefault(p["pernr"], []).append({
            "fpper": p["fpper"],
            "old_seqnr": p.get("original_seqnr") or "",
            "old_inper": p.get("original_inper") or "",
            "old_srtza": "P/O",
            "new_seqnr": p.get("current_seqnr") or "",
            "new_inper": p.get("current_inper") or "",
            "new_srtza": "A",
        })

    cat = {"B": 0, "A": 0, "C": 0, "D": 0, "E": 0}
    for pn, view in view_by_pernr.items():
        ppoix = {w: _d(view.get(f"ppoix_{w.strip('/')}", "0")) for w in PRIORITY_WAGE_TYPES}
        june_seqnr = view.get("current_seqnr", "")
        pp = pairs_by_pernr.get(pn, [])
        cases = _cases_for_pernr(pn, pp, ppoix, period, june_seqnr)
        if cases:
            cat["B"] += 1
            res.cases.extend(cases)
        elif _has_recalc_pair(pp, period):
            cat["E"] += 1
        else:
            cat["A"] += 1
    res.total_pernrs = len(view_by_pernr)
    res.ppoix_ref_total = _d(pc.get("ppoix_ref_total"))
    res.gap_ppoix_vs_ppdit = _d(pc.get("residual_notes", {}).get("ppoix_vs_ppdit"))
    res.category_counts = cat
    res.warnings.append("fonte: JSON guardado — OLD_SEQNR = 1ª execução do FPPER "
                        "(com RFC obtém-se a versão imediatamente anterior ao run).")
    return res


def _d(v: Any) -> Decimal:
    try:
        return Decimal(str(v or "0"))
    except Exception:
        return Decimal("0")


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

def build_manual_rt_shortlist(
    params: AnalysisParams,
    *,
    output_dir: Path,
    write_files: bool = True,
    connection: Any = None,
) -> ManualRtResult:
    """Sem `connection`: reconstrói do `payroll_cluster_analysis_*.json`.
    Com `connection`: corre as fases 1-3 (transparentes, sem tocar no cluster).
    """
    if connection is None:
        slug = f"{params.empresa}_{params.ano}_{params.mes:02d}_run{params.primary_run.zfill(10)}"
        candidates = [output_dir / f"payroll_cluster_analysis_{slug}.json",
                      *output_dir.glob("payroll_cluster_analysis_*.json")]
        for path in candidates:
            if path.exists():
                res = _cases_from_json(json.loads(path.read_text(encoding="utf-8")))
                res.warnings.append(f"fonte: {path.name}")
                break
        else:
            raise FileNotFoundError(
                "Sem payroll_cluster_analysis_*.json em output/ — correr a análise "
                "completa uma vez, ou usar --manual-rt-request com ligação RFC.")
    else:
        from .payroll_posting import analyze as analyze_payroll
        from .payroll_wagetypes import link_wage_types_to_posting_line
        from .payroll_cluster import analyse_cluster

        payroll = analyze_payroll(connection, params)
        link = link_wage_types_to_posting_line(connection, params, payroll)
        if not link.resolved:
            raise RuntimeError("Fase 2 não resolvida: " + "; ".join(link.warnings))
        cluster = analyse_cluster(connection, params, payroll, link, try_rt=False)
        res = _cases_from_cluster(cluster)
        res.source = "rfc"
        res.gap_ppoix_vs_ppdit = _d(cluster.residual_notes.get("ppoix_vs_ppdit"))
        res.ppoix_ref_total = abs(cluster.ppoix_ref_total)

    # gap PPOIX /558+/559 vs referência RH
    res.gap_ppoix_vs_rh = (abs(res.ppoix_ref_total or _sum_ref(res)) - params.valor_rh_referencia)
    if not res.ppoix_ref_total:
        res.ppoix_ref_total = _sum_ref(res)
    res.account = params.conta

    # ordenação: prioridade asc, impacto /558+/559 desc, nº recálculos desc
    res.cases.sort(key=lambda c: (c.priority, -abs(c.ppoix_ref), -c.retro_months))

    if write_files:
        output_dir.mkdir(parents=True, exist_ok=True)
        run = res.run_id or params.primary_run.zfill(10)
        res.csv_path = _write_csv(res, output_dir / f"manual_rt_shortlist_run{run}.csv")
        res.txt_path = _write_txt(res, output_dir / f"manual_rt_request_run{run}.txt")
    return res


def _sum_ref(res: ManualRtResult) -> Decimal:
    return abs(sum((c.ppoix_ref for c in res.cases), Decimal("0")))


# ---------------------------------------------------------------------------
# Escrita
# ---------------------------------------------------------------------------

_CSV_HEADER = ["PERNR", "FPPER", "OLD_SEQNR", "OLD_INPER", "OLD_SRTZA",
               "NEW_SEQNR", "NEW_INPER", "NEW_SRTZA", "PPOIX_558", "PPOIX_559",
               "PPOIX_561", "PPOIX_563", "PPOIX_0029", "PPOIX_TOTAL",
               "RETRO_MONTHS", "PRIORITY", "REASON"]


def _write_csv(res: ManualRtResult, path: Path) -> Path:
    import csv

    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(_CSV_HEADER)
        for c in res.cases:
            w.writerow(c.as_row(res.gap_ppoix_vs_rh))
    logger.info("CSV escrito: %s", path)
    return path


def _write_txt(res: ManualRtResult, path: Path) -> Path:
    sep = "=" * 60
    lines: list[str] = [
        sep,
        "PEDIDO DE EXTRACÇÃO MANUAL — PC00_M99_CWTR (Display Payroll Results)",
        sep,
        f"Empresa {res.company}   Run {res.run_id}   IN-period {res.period}   "
        f"Conta {res.account}",
        f"Objectivo: fechar o resíduo PPOIX /558+/559 vs RH = "
        f"{_fmt(res.gap_ppoix_vs_rh)} EUR (e PPOIX vs PPDIT = "
        f"{_fmt(res.gap_ppoix_vs_ppdit)} EUR).",
        "",
        "Extrair APENAS os casos abaixo (prioridade 1 e 2). Não pedir mais "
        "PERNR enquanto estes não forem analisados.",
        "",
        f"Campos a exportar por resultado: {', '.join(RT_FIELDS)}",
        f"Rubricas prioritárias: {', '.join(PRIORITY_WAGE_TYPES)}",
        "Preferência: exportar a RT completa dos dois SEQNR de cada caso.",
        "",
    ]
    prio_cases = [c for c in res.cases if c.priority <= 2]
    if not prio_cases:
        lines.append("(nenhum caso de prioridade 1/2 — ver a shortlist CSV para os "
                     "casos de prioridade 3.)")
    for i, c in enumerate(prio_cases, 1):
        lines += [
            sep,
            f"CASO {i}   [prioridade {c.priority}]",
            sep,
            f"PERNR: {c.pernr}",
            f"FPPER (período recalculado): {c.fpper}",
            "",
            "Resultado ORIGINAL do FPPER (provisório):",
            f"  SEQNR: {c.old_seqnr or '(primeira execução — não há anterior)'}",
            f"  INPER: {c.old_inper}",
            f"  SRTZA: {c.old_srtza}",
            "",
            "Resultado RECALCULADO do FPPER (definitivo actual):",
            f"  SEQNR: {c.new_seqnr}",
            f"  INPER: {c.new_inper}"
            + ("   (posterior ao run — verificar também a versão de "
               f"{res.period} se necessário)" if c.new_inper > res.period else ""),
            f"  SRTZA: {c.new_srtza}",
            "",
            f"Resultado do próprio período {res.period} deste PERNR (contexto): "
            f"SEQNR {c.june_seqnr or 'n/d'}",
            "",
            "PPOIX deste PERNR (transferido para a conta):",
            f"  /558  = {_fmt(c.ppoix_558)}",
            f"  /559  = {_fmt(c.ppoix_559)}",
            f"  /561  = {_fmt(c.ppoix_561)}",
            f"  /563  = {_fmt(c.ppoix_563)}",
            f"  0029  = {_fmt(c.ppoix_0029)}",
            f"  TOTAL = {_fmt(c.ppoix_total)}   "
            f"(/558+/559 = {_fmt(c.ppoix_ref)}; |558+559| = "
            f"{c.pct_of_gap(res.gap_ppoix_vs_rh)} do resíduo {_fmt(res.gap_ppoix_vs_rh)})",
            "",
            f"Motivo: {c.reason}",
            "",
            f"=> Preciso da RT do SEQNR {c.old_seqnr or '(n/a)'} e do SEQNR {c.new_seqnr} "
            f"para o PERNR {c.pernr}.",
            "",
        ]
    lines.append(sep)
    path.write_text("\n".join(lines), encoding="utf-8")
    logger.info("TXT escrito: %s", path)
    return path


def _fmt(v: Decimal | None) -> str:
    if v is None:
        return "(n/d)"
    q = Decimal(v).quantize(Decimal("0.01"))
    s = f"{abs(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if q < 0 else ''}{s}"


# ---------------------------------------------------------------------------
# Relatório terminal
# ---------------------------------------------------------------------------

def print_manual_rt_report(res: ManualRtResult) -> None:
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # consola Windows cp1252
    except Exception:  # pragma: no cover
        pass
    line = "-" * 72
    print("=" * 72)
    print(f"MANUAL RT REQUEST — run {res.run_id}  empresa {res.company}  "
          f"período {res.period}  (fonte: {res.source})")
    print("=" * 72)
    for w in res.warnings:
        print(f"  aviso: {w}")
    cc = res.category_counts
    tot = res.total_pernrs
    print(f"  {tot} PERNR no run. Categorias RGDIR (evidência sem RT):")
    print(f"    B recálculo real ...... {cc.get('B', 0):>4}   (correcção >=2 meses OU /561//563/0029)")
    print(f"    E rotina (ambíguo) .... {cc.get('E', 0):>4}   (finalização mensal de rotina — não confirmável sem RT)")
    print(f"    A retro s/ par ........ {cc.get('A', 0):>4}")
    print(f"    C off-cycle ........... {cc.get('C', 0):>4}     D void/reversal ... {cc.get('D', 0)}")
    print(f"  Resíduo a fechar: PPOIX /558+/559 vs RH = {_fmt(res.gap_ppoix_vs_rh)} EUR  |  "
          f"PPOIX vs PPDIT = {_fmt(res.gap_ppoix_vs_ppdit)} EUR")
    print("")
    manual = res.pernrs_manual
    p1 = sorted({c.pernr for c in res.cases if c.priority == 1})
    p2 = sorted({c.pernr for c in res.cases if c.priority == 2})
    n_cases = len(res.cases)
    print(f"  >>> CONSULTA MANUAL NECESSÁRIA: {len(manual)} PERNR  /  {n_cases} pares OLD->NEW SEQNR")
    print(f"      prioridade 1 (correcção real, {len(p1)} PERNR): {', '.join(p1) or '—'}")
    print(f"      prioridade 2 (carry-forward, {len(p2)} PERNR): {', '.join(p2) or '—'}")
    print(f"  ({cc.get('E', 0)} PERNR de rotina NÃO entram — só analisar se estes não fecharem.)")
    print("")
    print(line)
    print(f"  {'PERNR':<10}{'FPPER':<7}{'OLD SEQNR':<10}{'NEW SEQNR':<10}"
          f"{'/558+/559':>12}{'/561':>9}{'/563':>10}{'0029':>10}{' P':>3}  #FPPER")
    print(line)
    for c in res.cases:
        print(f"  {c.pernr:<10}{c.fpper:<7}{c.old_seqnr or '-':<10}{c.new_seqnr:<10}"
              f"{_fmt(c.ppoix_ref):>12}{_fmt(c.ppoix_561):>9}{_fmt(c.ppoix_563):>10}"
              f"{_fmt(c.ppoix_0029):>10}{c.priority:>3}  {c.retro_months}")
    print(line)
    if res.csv_path:
        print(f"  shortlist : {res.csv_path}")
    if res.txt_path:
        print(f"  pacote    : {res.txt_path}   ({n_cases} CASOS, formato PC00_M99_CWTR)")
