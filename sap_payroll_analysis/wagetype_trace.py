"""Fase 4.1 — rastreio contabilístico de uma rubrica (LGART) de um PERNR,
de PPOIX até PPDIT, no posting run.

    PPOIX  --(TSLIN == PPDIX.LINUM)-->  PPDIX  --(DOCNUM,DOCLIN)-->  PPDIT --> HKONT

100 % read-only (RFC_READ_TABLE via `safe_rfc_call`). Não lê o cluster PCL2,
não chama `PYXX_READ_PAYROLL_RESULT`. Genérico: `--run/--pernr/--lgart`.
"""

from __future__ import annotations

import csv
import json
import logging
from dataclasses import dataclass, field
from decimal import Decimal
from itertools import combinations
from pathlib import Path
from typing import Any

from .config import AnalysisParams
from .payroll_wagetypes import resolve_account_determination
from .sap_reader import NoData, RfcReadError, opt_and, opt_eq, opt_in, read_table, sap_str_to_decimal

logger = logging.getLogger(__name__)

_PPOIX_FIELDS = ["RUNID", "PERNR", "SEQNO", "ACTSIGN", "POSTNUM", "SPPRC", "RTLINE",
                "KOART", "MOMAG", "KOMOK", "TSLIN", "LGART", "BETRG",
                "SWRETROACT", "SWPER", "NEG_POSTNG"]
#: subconjunto de recurso se RFC_READ_TABLE lançar SAPSQL_DATA_LOSS (quirk
#: intermitente de RFC_READ_TABLE sobre PPOIX).
_PPOIX_FIELDS_MIN = ["RUNID", "PERNR", "SEQNO", "LGART", "BETRG", "TSLIN", "KOMOK",
                     "MOMAG", "POSTNUM", "RTLINE", "ACTSIGN", "NEG_POSTNG"]
_PPDIT_FIELDS = ["DOCNUM", "DOCLIN", "BUKRS", "HKONT", "KTOSL", "WRBTR", "WAERS",
                 "PERNR", "NEG_POSTNG", "ITTYP", "SGTXT"]


def _signed(betrg_raw: str, neg_postng: str) -> Decimal:
    """Valor PPOIX com sinal: BETRG já traz sinal SAP; NEG_POSTNG='X' inverte."""
    v = sap_str_to_decimal(betrg_raw)
    return -v if str(neg_postng).strip().upper() == "X" else v


# ---------------------------------------------------------------------------
# Modelos
# ---------------------------------------------------------------------------

@dataclass
class PpoixRow:
    runid: str = ""
    pernr: str = ""
    seqno: str = ""
    lgart: str = ""
    betrg_raw: str = ""
    betrg: Decimal = Decimal("0")
    komok: str = ""
    koart: str = ""
    momag: str = ""
    postnum: str = ""
    rtline: str = ""
    tslin: str = ""
    actsign: str = ""
    spprc: str = ""
    swretroact: str = ""
    swper: str = ""
    neg_postng: str = ""
    ppdix_dest: list[tuple[str, str]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "runid": self.runid, "pernr": self.pernr, "seqno": self.seqno, "lgart": self.lgart,
            "betrg_raw": self.betrg_raw, "betrg": str(self.betrg), "komok": self.komok,
            "koart": self.koart, "momag": self.momag, "postnum": self.postnum,
            "rtline": self.rtline, "tslin": self.tslin, "actsign": self.actsign,
            "spprc": self.spprc, "swretroact": self.swretroact, "swper": self.swper,
            "neg_postng": self.neg_postng,
            "ppdix_dest": [f"{d}/{l}" for d, l in self.ppdix_dest],
        }


@dataclass
class PpdixRow:
    runid: str = ""
    linum: str = ""
    docnum: str = ""
    doclin: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {"runid": self.runid, "linum": self.linum, "docnum": self.docnum, "doclin": self.doclin}


@dataclass
class PpditRow:
    docnum: str = ""
    doclin: str = ""
    bukrs: str = ""
    hkont: str = ""
    ktosl: str = ""
    wrbtr_raw: str = ""
    wrbtr: Decimal = Decimal("0")
    waers: str = ""
    pernr: str = ""
    neg_postng: str = ""
    is_target_account: bool = False

    def as_dict(self) -> dict[str, Any]:
        return {
            "docnum": self.docnum, "doclin": self.doclin, "bukrs": self.bukrs,
            "hkont": self.hkont, "ktosl": self.ktosl, "wrbtr_raw": self.wrbtr_raw,
            "wrbtr": str(self.wrbtr), "waers": self.waers, "pernr": self.pernr,
            "neg_postng": self.neg_postng, "is_target_account": self.is_target_account,
        }


@dataclass
class SignPath:
    lgart: str = ""
    ppoix_betrg_raw: str = ""
    ppoix_signed: Decimal = Decimal("0")
    actsign: str = ""
    neg_postng: str = ""
    ppdit_wrbtr_raw: str = ""
    ppdit_signed: Decimal = Decimal("0")
    accounting_effect: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "lgart": self.lgart,
            "ppoix_betrg_raw": self.ppoix_betrg_raw, "ppoix_signed": str(self.ppoix_signed),
            "actsign": self.actsign, "neg_postng": self.neg_postng,
            "ppdit_wrbtr_raw": self.ppdit_wrbtr_raw, "ppdit_signed": str(self.ppdit_signed),
            "accounting_effect": self.accounting_effect,
        }


@dataclass
class WageTypeTrace:
    run_id: str = ""
    pernr: str = ""
    lgart: str = ""
    company: str = ""
    account: str = ""

    ppoix: list[PpoixRow] = field(default_factory=list)
    ppdix: list[PpdixRow] = field(default_factory=list)
    ppdit: list[PpditRow] = field(default_factory=list)

    target_doc_lines: list[tuple[str, str]] = field(default_factory=list)
    reaches_account: bool = False
    reaches_target_line: bool = False
    transferred_tslins: list[str] = field(default_factory=list)
    non_transferred_tslins: list[str] = field(default_factory=list)

    same_transfer_line_components: list[dict[str, Any]] = field(default_factory=list)
    transfer_line_by_lgart: dict[str, dict[str, Any]] = field(default_factory=dict)
    transfer_line_by_momag: dict[str, dict[str, Any]] = field(default_factory=dict)
    transfer_line_by_tslin: dict[str, dict[str, Any]] = field(default_factory=dict)

    account_determination: dict[str, Any] = field(default_factory=dict)
    sign_path: SignPath = field(default_factory=SignPath)
    compare: "WageTypeTrace | None" = None

    reconciliation: dict[str, Any] = field(default_factory=dict)
    residual_investigation: dict[str, Any] = field(default_factory=dict)
    conclusion: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def warn(self, m: str) -> None:
        if m not in self.warnings:
            self.warnings.append(m)
            logger.warning(m)

    def as_dict(self) -> dict[str, Any]:
        d = {
            "input": {"run": self.run_id, "pernr": self.pernr, "lgart": self.lgart,
                      "company": self.company, "account": self.account},
            "ppoix": [r.as_dict() for r in self.ppoix],
            "ppdix": [r.as_dict() for r in self.ppdix],
            "ppdit": [r.as_dict() for r in self.ppdit],
            "target_doc_lines": [f"{d}/{l}" for d, l in self.target_doc_lines],
            "reaches_account": self.reaches_account,
            "reaches_target_line": self.reaches_target_line,
            "transferred_tslins": self.transferred_tslins,
            "non_transferred_tslins": self.non_transferred_tslins,
            "same_transfer_line_components": self.same_transfer_line_components,
            "transfer_line_by_lgart": self.transfer_line_by_lgart,
            "transfer_line_by_momag": self.transfer_line_by_momag,
            "transfer_line_by_tslin": self.transfer_line_by_tslin,
            "account_determination": self.account_determination,
            "sign_path": self.sign_path.as_dict(),
            "reconciliation": self.reconciliation,
            "residual_investigation": self.residual_investigation,
            "conclusion": self.conclusion,
            "warnings": self.warnings,
        }
        if self.compare is not None:
            d["compare"] = self.compare.as_dict()
        return d


# ---------------------------------------------------------------------------
# Trace
# ---------------------------------------------------------------------------

def _read_ppoix_options(connection: Any, params: AnalysisParams, options: list[dict[str, str]],
                        *, max_rows: int | None = None) -> list[dict[str, str]]:
    """PPOIX com recurso: se `SAPSQL_DATA_LOSS` (quirk RFC_READ_TABLE), repete
    com um subconjunto de campos mais estreito."""
    for fields in (_PPOIX_FIELDS, _PPOIX_FIELDS_MIN):
        try:
            return read_table(connection, "PPOIX", fields=fields, options=options,
                              page_size=min(params.page_size, 20_000), max_rows=max_rows).rows
        except NoData:
            return []
        except RfcReadError as exc:
            if fields is _PPOIX_FIELDS_MIN or "DATA_LOSS" not in exc.kind and "DATA WAS LOST" not in str(exc).upper():
                raise
            logger.warning("PPOIX SAPSQL_DATA_LOSS — a repetir com campos reduzidos.")
    return []


def _read_ppoix(connection: Any, params: AnalysisParams, run: str,
                pernr: str, lgart: str) -> list[dict[str, str]]:
    """Lê PPOIX por RUNID+PERNR (resultado pequeno) e filtra LGART em Python.

    Nota: um filtro `LGART = '/559'` em OPTIONS de RFC_READ_TABLE provoca
    `SAPSQL_DATA_LOSS` neste sistema (valor a começar por '/'). Evita-se.
    """
    rows = _read_ppoix_options(
        connection, params,
        opt_and(opt_eq("RUNID", run), opt_eq("PERNR", pernr.zfill(8))),
    )
    if lgart:
        rows = [r for r in rows if r.get("LGART") == lgart]
    return rows


def _ppdix_map(connection: Any, params: AnalysisParams, run: str) -> dict[str, list[tuple[str, str]]]:
    rows = read_table(connection, "PPDIX", fields=["RUNID", "LINUM", "DOCNUM", "DOCLIN"],
                      options=opt_and(opt_eq("RUNID", run)), page_size=params.page_size).rows
    out: dict[str, list[tuple[str, str]]] = {}
    for r in rows:
        out.setdefault(r["LINUM"], []).append((r["DOCNUM"], r["DOCLIN"]))
    return out


def _target_account_lines(connection: Any, params: AnalysisParams, run: str,
                          ) -> tuple[list[tuple[str, str]], list[PpditRow]]:
    """Linha(s) PPDIT da conta `params.conta` / empresa `params.empresa` no run."""
    hdr = read_table(connection, "PPDHD", fields=["DOCNUM", "RUNID"],
                     options=opt_and(opt_eq("RUNID", run)), page_size=params.page_size).rows
    docs = sorted({r["DOCNUM"] for r in hdr})
    if not docs:
        return [], []
    acct = {params.conta_10, params.conta.strip()}
    rows: list[dict[str, str]] = []
    for s in range(0, len(docs), 100):
        try:
            rows += read_table(connection, "PPDIT", fields=_PPDIT_FIELDS,
                               options=opt_and(opt_in("DOCNUM", docs[s:s + 100])),
                               page_size=params.page_size).rows
        except NoData:
            pass
    tgt: list[PpditRow] = []
    for r in rows:
        if r.get("BUKRS") == params.empresa and (
            r.get("HKONT") in acct or r.get("HKONT", "").lstrip("0") == params.conta.lstrip("0")
        ):
            tgt.append(_ppdit_row(r, is_target=True))
    return sorted({(t.docnum, t.doclin) for t in tgt}), tgt


def _ppdit_row(r: dict[str, str], *, is_target: bool = False) -> PpditRow:
    return PpditRow(
        docnum=r.get("DOCNUM", ""), doclin=r.get("DOCLIN", ""), bukrs=r.get("BUKRS", ""),
        hkont=r.get("HKONT", ""), ktosl=r.get("KTOSL", ""), wrbtr_raw=r.get("WRBTR", ""),
        wrbtr=_signed(r.get("WRBTR", "0"), r.get("NEG_POSTNG", "")), waers=r.get("WAERS", ""),
        pernr=r.get("PERNR", ""), neg_postng=r.get("NEG_POSTNG", ""), is_target_account=is_target,
    )


def explain_amount_sign_path(ppoix: PpoixRow, ppdit: PpditRow | None) -> SignPath:
    """Documenta cada sinal em separado; não deduz D/C só do texto do BETRG."""
    sp = SignPath(
        lgart=ppoix.lgart, ppoix_betrg_raw=ppoix.betrg_raw, ppoix_signed=ppoix.betrg,
        actsign=ppoix.actsign, neg_postng=ppoix.neg_postng or "(vazio)",
    )
    if ppdit is not None:
        sp.ppdit_wrbtr_raw = ppdit.wrbtr_raw
        sp.ppdit_signed = ppdit.wrbtr
        # 23120000 é conta de passivo (KOART F). BETRG/WRBTR negativo => credita
        # o passivo (aumenta o valor a pagar). Positivo => debita (reduz).
        if ppdit.wrbtr < 0:
            sp.accounting_effect = f"crédito em {ppdit.hkont.lstrip('0')} (aumenta o passivo)"
        elif ppdit.wrbtr > 0:
            sp.accounting_effect = f"débito em {ppdit.hkont.lstrip('0')} (reduz o passivo)"
        else:
            sp.accounting_effect = "nulo"
    else:
        sp.accounting_effect = "não transferido (TSLIN sem destino PPDIX)"
    return sp


def trace_wagetype(connection: Any, params: AnalysisParams, *, pernr: str, lgart: str,
                   compare_lgart: str | None = None, _is_compare: bool = False) -> WageTypeTrace:
    run = params.primary_run_10
    tr = WageTypeTrace(run_id=run, pernr=pernr.zfill(8), lgart=lgart,
                       company=params.empresa, account=params.conta_10)

    dixmap = _ppdix_map(connection, params, run)
    tgt_keys, tgt_ppdit = _target_account_lines(connection, params, run)
    tr.target_doc_lines = tgt_keys
    tgt_set = set(tgt_keys)

    # --- PPOIX do PERNR/LGART (filtro reforçado no cliente) ---
    pset = {pernr, pernr.zfill(8), pernr.lstrip("0")}
    for r in _read_ppoix(connection, params, run, pernr, lgart):
        if r.get("RUNID") != run or r.get("PERNR") not in pset:
            continue
        if lgart and r.get("LGART") != lgart:
            continue
        row = PpoixRow(
            runid=r.get("RUNID", ""), pernr=r.get("PERNR", ""), seqno=r.get("SEQNO", ""),
            lgart=r.get("LGART", ""), betrg_raw=r.get("BETRG", ""),
            betrg=_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")),
            komok=r.get("KOMOK", ""), koart=r.get("KOART", ""), momag=r.get("MOMAG", ""),
            postnum=r.get("POSTNUM", ""), rtline=r.get("RTLINE", ""), tslin=r.get("TSLIN", ""),
            actsign=r.get("ACTSIGN", ""), spprc=r.get("SPPRC", ""),
            swretroact=r.get("SWRETROACT", ""), swper=r.get("SWPER", ""),
            neg_postng=r.get("NEG_POSTNG", ""),
        )
        row.ppdix_dest = dixmap.get(row.tslin, [])
        tr.ppoix.append(row)

    if not tr.ppoix:
        tr.warn(f"Sem PPOIX para run {run} / PERNR {pernr} / LGART {lgart}.")
        tr.conclusion = {"reaches_account": False, "reaches_target_line": False,
                         "residual_class": "N/A"}
        return tr

    # --- PPDIX + PPDIT dos destinos ---
    dest_keys: set[tuple[str, str]] = set()
    for row in tr.ppoix:
        if row.ppdix_dest:
            tr.transferred_tslins.append(row.tslin)
            for d, l in row.ppdix_dest:
                dest_keys.add((d, l))
                tr.ppdix.append(PpdixRow(runid=run, linum=row.tslin, docnum=d, doclin=l))
        elif row.tslin in {"", "0000000000"} or row.tslin.strip("0") == "":
            tr.non_transferred_tslins.append(row.tslin)
        else:
            tr.non_transferred_tslins.append(row.tslin)
    tr.transferred_tslins = sorted(set(tr.transferred_tslins))
    tr.non_transferred_tslins = sorted(set(tr.non_transferred_tslins))

    if dest_keys:
        docs = sorted({d for d, _ in dest_keys})
        prows: list[dict[str, str]] = []
        for s in range(0, len(docs), 100):
            try:
                prows += read_table(connection, "PPDIT", fields=_PPDIT_FIELDS,
                                    options=opt_and(opt_in("DOCNUM", docs[s:s + 100])),
                                    page_size=params.page_size).rows
            except NoData:
                pass
        for r in prows:
            if (r["DOCNUM"], r["DOCLIN"]) in dest_keys:
                pr = _ppdit_row(r, is_target=(r["DOCNUM"], r["DOCLIN"]) in tgt_set)
                tr.ppdit.append(pr)

    tr.reaches_account = any(p.is_target_account for p in tr.ppdit) or bool(dest_keys & tgt_set)
    tr.reaches_target_line = bool(dest_keys & tgt_set)

    # --- TODOS os TSLIN/LINUM que alimentam a linha PPDIT alvo (não só os do
    #     wage type rastreado): a linha colectora agrega vários LINUM. ---
    target_tslins = sorted({ln for ln, dests in dixmap.items()
                            if any((d, l) in tgt_set for d, l in dests)})
    if target_tslins and tr.reaches_target_line:
        _aggregate_transfer_line(connection, params, run, target_tslins, tgt_keys, tgt_ppdit, tr)

    # --- determinação de contas para a LGART ---
    symkos = sorted({row.komok for row in tr.ppoix if row.komok})
    ktosl = next((p.ktosl for p in tr.ppdit if p.is_target_account), params.hr_posting_key)
    try:
        tr.account_determination = resolve_account_determination(
            connection, params, symkos=symkos, wage_types=[lgart], ktosl=ktosl)
    except (RfcReadError, Exception) as exc:  # noqa: BLE001
        tr.warn(f"Determinação de contas indisponível: {exc}")

    # --- sign path (usa a linha PPOIX transferida para a conta alvo) ---
    ppoix_to_target = next((r for r in tr.ppoix
                            if any((d, l) in tgt_set for d, l in r.ppdix_dest)), None)
    ppdit_target = next((p for p in tr.ppdit if p.is_target_account), None)
    tr.sign_path = explain_amount_sign_path(ppoix_to_target or tr.ppoix[0], ppdit_target)

    # --- comparação com outra LGART do mesmo PERNR ---
    if compare_lgart and not _is_compare and compare_lgart != lgart:
        tr.compare = trace_wagetype(connection, params, pernr=pernr, lgart=compare_lgart,
                                    _is_compare=True)

    _build_conclusion(params, tr)
    return tr


def _aggregate_transfer_line(connection: Any, params: AnalysisParams, run: str,
                             tslins: list[str], tgt_keys: list[tuple[str, str]],
                             tgt_ppdit: list[PpditRow], tr: WageTypeTrace) -> None:
    """Lê TODOS os PPOIX com RUNID + TSLIN in tslins e agrega."""
    try:
        rows = _read_ppoix_options(connection, params, opt_and(opt_eq("RUNID", run)),
                                   max_rows=500_000)
    except (NoData, RfcReadError) as exc:
        tr.warn(f"Falha a agregar o TSLIN: {exc}")
        return
    tset = set(tslins)
    comp = [r for r in rows if r.get("TSLIN") in tset]

    by_lgart: dict[str, list[Any]] = {}
    by_momag: dict[str, list[Any]] = {}
    by_tslin: dict[str, list[Any]] = {}
    total = Decimal("0")
    for r in comp:
        v = _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))
        total += v
        for key, dic in ((r.get("LGART", ""), by_lgart), (r.get("MOMAG", ""), by_momag),
                         (r.get("TSLIN", ""), by_tslin)):
            e = dic.setdefault(key, [0, Decimal("0")])
            e[0] += 1
            e[1] += v
        tr.same_transfer_line_components.append({
            "pernr": r.get("PERNR", ""), "seqno": r.get("SEQNO", ""), "lgart": r.get("LGART", ""),
            "betrg": str(v), "komok": r.get("KOMOK", ""), "tslin": r.get("TSLIN", ""),
            "postnum": r.get("POSTNUM", ""), "rtline": r.get("RTLINE", ""),
            "actsign": r.get("ACTSIGN", ""), "neg_postng": r.get("NEG_POSTNG", ""),
            "momag": r.get("MOMAG", ""), "swretroact": r.get("SWRETROACT", ""),
        })
    tr.transfer_line_by_lgart = {k: {"rows": v[0], "sum": str(v[1])} for k, v in sorted(by_lgart.items())}
    tr.transfer_line_by_momag = {k: {"rows": v[0], "sum": str(v[1])} for k, v in sorted(by_momag.items())}
    tr.transfer_line_by_tslin = {k: {"rows": v[0], "sum": str(v[1])} for k, v in sorted(by_tslin.items())}

    ppdit_sum = sum((p.wrbtr for p in tgt_ppdit if (p.docnum, p.doclin) in set(tgt_keys)), Decimal("0"))
    delta = total - ppdit_sum
    # contribuição da rubrica rastreada (deste PERNR) dentro da linha
    traced = sum((_signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", ""))
                  for r in comp if r.get("LGART") == tr.lgart and r.get("PERNR") == tr.pernr),
                 Decimal("0"))
    tr.reconciliation = {
        "transfer_line_tslins": tslins,
        "ppoix_rows": len(comp),
        "ppoix_sum": str(total),
        "ppdit_target_line": [f"{d}/{l}" for d, l in tgt_keys],
        "ppdit_wrbtr": str(ppdit_sum),
        "delta": str(delta),
        "traced_row_in_line": str(traced),
        "leftover_if_traced_row_excluded": str(delta - traced),
    }
    _investigate_residual(comp, delta, tr, traced)


def _investigate_residual(comp: list[dict[str, str]], delta: Decimal, tr: WageTypeTrace,
                          traced: Decimal = Decimal("0")) -> None:
    """Procura evidência OBJECTIVA para o `delta`. Nunca classifica como
    arredondamento sem prova nos dados."""
    leftover = delta - traced
    close = traced != 0 and abs(abs(delta) - abs(traced)) <= Decimal("5.00")
    note = (
        (f"delta ({delta}) ≈ a rubrica rastreada deste PERNR ({traced}); "
         if close else
         f"a rubrica rastreada deste PERNR vale {traced}; ")
        + f"MAS ela JÁ está incluída na soma PPOIX — excluí-la é arbitrário e "
          f"deixa {leftover}. Coincidência aritmética, não prova de causa."
    )
    inv: dict[str, Any] = {
        "delta": str(delta), "abs_delta": str(abs(delta)),
        "traced_row_in_line": str(traced),
        "leftover_if_traced_row_excluded": str(leftover),
        "traced_row_near_delta": bool(close),
        "note_coincidence": note,
    }
    rows = [(r.get("PERNR", ""), r.get("LGART", ""),
             _signed(r.get("BETRG", "0"), r.get("NEG_POSTNG", "")),
             r.get("ACTSIGN", ""), r.get("NEG_POSTNG", "")) for r in comp]

    # A. valores absolutos <= 1,00
    inv["rows_abs_le_1"] = [{"pernr": p, "lgart": l, "betrg": str(b)} for p, l, b, *_ in rows
                            if abs(b) <= Decimal("1.00")]
    # I. linha única == delta / -delta / leftover
    inv["single_row_equals_delta"] = [{"pernr": p, "lgart": l, "betrg": str(b)}
                                      for p, l, b, *_ in rows if b in (delta, -delta)]
    inv["single_row_equals_leftover"] = [{"pernr": p, "lgart": l, "betrg": str(b)}
                                         for p, l, b, *_ in rows if b in (leftover, -leftover)]
    # fractional cents (arredondamento observável) — só se existir mesmo
    inv["rows_with_sub_cent"] = [{"pernr": p, "lgart": l, "betrg": str(b)} for p, l, b, *_ in rows
                                 if b != b.quantize(Decimal("0.01"))]
    # F/G. ACTSIGN != 'A' ou NEG_POSTNG marcado
    inv["rows_actsign_not_A"] = [{"pernr": p, "lgart": l, "betrg": str(b), "actsign": a}
                                 for p, l, b, a, _ in rows if a != "A"]
    inv["rows_neg_postng"] = [{"pernr": p, "lgart": l, "betrg": str(b)} for p, l, b, _, n in rows
                              if str(n).strip().upper() == "X"]
    # B. combinação de pequenos valores que totalize o delta OU o leftover
    def _combos(target: Decimal) -> list[list[dict[str, str]]]:
        cap = max(abs(target) * 3, Decimal("50"))
        small = [(p, l, b) for p, l, b, *_ in rows if Decimal("0") < abs(b) <= cap]
        bs = [b for _, _, b in small]
        found: list[list[dict[str, str]]] = []
        for k in (1, 2, 3):
            for combo in combinations(range(len(bs)), k):
                if abs(sum(bs[i] for i in combo) - target) < Decimal("0.005"):
                    found.append([{"pernr": small[i][0], "lgart": small[i][1],
                                   "betrg": str(small[i][2])} for i in combo])
                    if len(found) >= 12:
                        return found
        return found

    inv["arithmetic_combos_for_delta"] = _combos(delta)
    inv["arithmetic_combos_for_leftover"] = _combos(leftover) if leftover != delta else []
    inv["arithmetic_combos_note"] = (
        "Combinações que somam o delta/leftover são coincidência aritmética — só "
        "têm valor se houver razão contabilística (mesma linha, par F/C, mesmo "
        "PERNR). Não são prova."
    )

    # classificação — nunca 'arredondamento' sem prova
    if inv["single_row_equals_delta"]:
        inv["classification"] = "EXPLAINED"
        inv["explanation"] = "Existe uma linha PPOIX cujo valor é exactamente o delta."
    elif inv["rows_with_sub_cent"]:
        inv["classification"] = "PARTIALLY_EXPLAINED"
        inv["explanation"] = ("Há valores PPOIX com fracção de cêntimo — possível "
                              "arredondamento na agregação (confirmar caso a caso).")
    else:
        inv["classification"] = "UNEXPLAINED"
        inv["explanation"] = (
            "O delta entre a soma dos itens PPOIX dos TSLIN e a linha PPDIT não "
            "corresponde a nenhuma linha isolada, marca de sinal (ACTSIGN/NEG_POSTNG), "
            "fracção de cêntimo ou split visível nas tabelas de posting. É produzido "
            "dentro do programa de posting HR ao construir a linha colectora. "
            "Hipótese (NÃO provada aqui): netting de deltas de retro já lançados em "
            "runs anteriores — o PPOIX carrega o resultado recalculado bruto e a "
            "PPDIT só a diferença. Demonstrável apenas com o trace do RPCIPE00 ou o "
            "documento FI, indisponíveis neste sistema."
        )
    tr.residual_investigation = inv


def _build_conclusion(params: AnalysisParams, tr: WageTypeTrace) -> None:
    ad = tr.account_determination or {}
    reason = ""
    hits = ad.get("wage_types_to_target") or []
    if tr.lgart in hits:
        koart = ",".join(sorted({r.get("KOART", "") for r in ad.get("t52ek", [])}))
        symkos = ",".join(ad.get("symkos", []))
        reason = (f"LGART {tr.lgart} -> SYMKO {symkos} (KOART {koart}); T030 "
                  f"KTOPL/{ad.get('ktosl', '?')}/BWMOD={tr.lgart} -> {params.conta}.")
    tr.conclusion = {
        "reaches_account": tr.reaches_account,
        "reaches_target_line": tr.reaches_target_line,
        "target_line": [f"{d}/{l}" for d, l in tr.target_doc_lines],
        "account_determination_reason": reason or "(não confirmada em T030 — ver account_determination)",
        "reconciliation_delta": tr.reconciliation.get("delta"),
        "residual_class": tr.residual_investigation.get("classification", "N/A"),
    }
    if tr.compare is not None:
        a = {(d, l) for d, l in tr.target_doc_lines}
        same = tr.reaches_target_line and tr.compare.reaches_target_line
        tr.conclusion["compare_lgart"] = tr.compare.lgart
        tr.conclusion["same_posting_line_as_compare"] = bool(same)


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_trace_json(tr: WageTypeTrace, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(tr.as_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info("JSON escrito: %s", path)
    return path


def write_trace_csv(tr: WageTypeTrace, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["kind", "pernr", "seqno", "lgart", "betrg", "komok", "momag",
                    "postnum", "rtline", "tslin", "actsign", "neg_postng", "docnum", "doclin", "hkont", "wrbtr"])
        for r in tr.ppoix:
            dest = r.ppdix_dest[0] if r.ppdix_dest else ("", "")
            w.writerow(["PPOIX", r.pernr, r.seqno, r.lgart, r.betrg, r.komok, r.momag,
                        r.postnum, r.rtline, r.tslin, r.actsign, r.neg_postng, dest[0], dest[1], "", ""])
        for r in tr.ppdix:
            w.writerow(["PPDIX", "", "", "", "", "", "", "", "", r.linum, "", "", r.docnum, r.doclin, "", ""])
        for r in tr.ppdit:
            w.writerow(["PPDIT", r.pernr, "", "", "", "", "", "", "", "", "", r.neg_postng,
                        r.docnum, r.doclin, r.hkont, r.wrbtr])
        for r in tr.same_transfer_line_components:
            w.writerow(["TSLIN_COMP", r["pernr"], r["seqno"], r["lgart"], r["betrg"], r["komok"],
                        r["momag"], r["postnum"], r["rtline"], r["tslin"], r["actsign"],
                        r["neg_postng"], "", "", "", ""])
        for lg, info in tr.transfer_line_by_lgart.items():
            w.writerow(["TSLIN_BY_LGART", "", "", lg, info["sum"], "", "", "", str(info["rows"]),
                        "", "", "", "", "", "", ""])
    logger.info("CSV escrito: %s", path)
    return path


def _fmt(v: Any) -> str:
    if v in (None, ""):
        return "(n/d)"
    try:
        q = Decimal(str(v)).quantize(Decimal("0.01"))
    except Exception:
        return str(v)
    s = f"{abs(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if q < 0 else ''}{s}"


def print_trace_report(tr: WageTypeTrace) -> None:
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # pragma: no cover
        pass
    L = "=" * 64
    print(L)
    print("TRACE WAGE TYPE")
    print(L)
    print(f"Run.............. {tr.run_id}")
    print(f"PERNR............ {tr.pernr}")
    print(f"LGART........... {tr.lgart}")
    print(f"Empresa/Conta... {tr.company} / {tr.account}")
    print("")
    for r in tr.ppoix:
        dest = ", ".join(f"{d}/{l}" for d, l in r.ppdix_dest) or "(sem destino — TSLIN não transferido)"
        print(f"  PPOIX  SEQNO {r.seqno}  BETRG {_fmt(r.betrg)}  KOMOK {r.komok}  MOMAG {r.momag}  "
              f"ACTSIGN {r.actsign}  NEG {r.neg_postng or '-'}  TSLIN {r.tslin}")
        print(f"         POSTNUM {r.postnum}  RTLINE {r.rtline}  ->  PPDIX/PPDIT {dest}")
    print("")
    for p in tr.ppdit:
        tag = "  <== CONTA ALVO" if p.is_target_account else ""
        print(f"  PPDIT  {p.docnum}/{p.doclin}  HKONT {p.hkont}  KTOSL {p.ktosl}  "
              f"WRBTR {_fmt(p.wrbtr)}{tag}")
    print("")
    print("-" * 64)
    print("DESTINO")
    print("-" * 64)
    tl = ", ".join(f"{d}/{l}" for d, l in tr.target_doc_lines) or "n/d"
    print(f"  {tr.lgart} -> conta {tr.account.lstrip('0')} : {'SIM' if tr.reaches_account else 'NÃO'}")
    print(f"  {tr.lgart} -> PPDIT {tl} : {'SIM' if tr.reaches_target_line else 'NÃO'}")
    ad = tr.account_determination or {}
    if ad:
        el = ad.get("t52el", [])
        cur = [x for x in el if x.get("ENDDA", "") in ("99991231", "")]
        print(f"  T52EL: {tr.lgart} -> SYMKO "
              + ", ".join(sorted({x.get('SYMKO', '') for x in (cur or el)}))
              + " (SIGN " + ", ".join(sorted({x.get('SIGN', '') for x in (cur or el)})) + ")")
        print(f"  T52EK: KOART " + ", ".join(sorted({x.get('KOART', '') for x in ad.get('t52ek', [])})))
        t030 = [r for r in ad.get("t030", []) if (r.get("KONTS") or "").strip("0")]
        for r in t030:
            print(f"  T030 : KTOPL {r['KTOPL']} / KTOSL {ad.get('ktosl')} / BWMOD {r['BWMOD']} / "
                  f"KOMOK {r['KOMOK']} -> KONTS {r['KONTS'].lstrip('0')}  KONTH {r['KONTH'].lstrip('0')}")
        if not t030:
            print("  T030 : (sem entrada BWMOD com conta directa — ver JSON)")
        print(f"  => {tr.conclusion.get('account_determination_reason', '')}")
    print("")
    sp = tr.sign_path
    print("-" * 64)
    print("SINAIS (payroll vs contabilístico)")
    print("-" * 64)
    print(f"  PPOIX BETRG (texto SAP) . {sp.ppoix_betrg_raw}   -> com sinal {_fmt(sp.ppoix_signed)}")
    print(f"  ACTSIGN ................ {sp.actsign}")
    print(f"  NEG_POSTNG ............. {sp.neg_postng}")
    print(f"  PPDIT WRBTR (texto SAP)  {sp.ppdit_wrbtr_raw or '(n/d)'}   -> com sinal {_fmt(sp.ppdit_signed)}")
    print(f"  efeito contabilístico .. {sp.accounting_effect}")
    print("")
    if tr.reconciliation:
        rec = tr.reconciliation
        print("-" * 64)
        print("AGREGAÇÃO DO TSLIN")
        print("-" * 64)
        print(f"  TSLIN(s) que atingem a linha alvo: {', '.join(rec['transfer_line_tslins'])}")
        for lg, info in tr.transfer_line_by_lgart.items():
            print(f"    {lg:<8} {info['rows']:>4} linhas  {_fmt(info['sum']):>16}")
        print(f"    {'por MOMAG':<8}")
        for mg, info in tr.transfer_line_by_momag.items():
            print(f"      MOMAG {mg or '-'}: {info['rows']} linhas  {_fmt(info['sum'])}")
        print(f"  SUM PPOIX (TSLIN) .... {_fmt(rec['ppoix_sum'])}")
        print(f"  PPDIT WRBTR .......... {_fmt(rec['ppdit_wrbtr'])}")
        print(f"  DELTA ............... {_fmt(rec['delta'])}")
        print("")
        inv = tr.residual_investigation
        print("-" * 64)
        print(f"RESÍDUO  (delta = {_fmt(inv.get('delta'))})")
        print("-" * 64)
        print(f"  rubrica {tr.lgart} deste PERNR na linha .. {_fmt(rec.get('traced_row_in_line'))}")
        print(f"  leftover se essa linha fosse excluída ... {_fmt(rec.get('leftover_if_traced_row_excluded'))}")
        print(f"    -> {inv.get('note_coincidence', '')}")
        print(f"  linhas |BETRG| <= 1,00 ................. {len(inv.get('rows_abs_le_1', []))}")
        print(f"  linha única == delta ................. {inv.get('single_row_equals_delta') or '—'}")
        print(f"  linha única == leftover .............. {inv.get('single_row_equals_leftover') or '—'}")
        print(f"  fracções de cêntimo ................. {len(inv.get('rows_with_sub_cent', []))}")
        print(f"  ACTSIGN != A / NEG_POSTNG='X' ....... {len(inv.get('rows_actsign_not_A', []))} / "
              f"{len(inv.get('rows_neg_postng', []))}")
        for tag, key in (("delta", "arithmetic_combos_for_delta"),
                         ("leftover", "arithmetic_combos_for_leftover")):
            combos = inv.get(key, [])
            if combos:
                print(f"  combinações que somam o {tag} (COINCIDÊNCIA, não prova):")
                for cmb in combos[:4]:
                    print("     " + " + ".join(f"{x['lgart']}/{x['pernr']}={x['betrg']}" for x in cmb))
        print(f"  => CLASSIFICAÇÃO DO RESÍDUO: {inv.get('classification')}")
        print(f"     {inv.get('explanation', '')}")
    if tr.compare is not None:
        print("")
        print("-" * 64)
        print(f"COMPARAÇÃO COM {tr.compare.lgart}")
        print("-" * 64)
        c = tr.compare
        print(f"  {c.lgart}: PPOIX rows={len(c.ppoix)}  -> conta={('SIM' if c.reaches_account else 'NÃO')}  "
              f"-> PPDIT alvo={('SIM' if c.reaches_target_line else 'NÃO')}")
        for r in c.ppoix:
            dest = ", ".join(f"{d}/{l}" for d, l in r.ppdix_dest) or "(não transferido)"
            print(f"    SEQNO {r.seqno}  BETRG {_fmt(r.betrg)}  TSLIN {r.tslin}  -> {dest}")
        same = tr.conclusion.get("same_posting_line_as_compare")
        print(f"  {tr.lgart} e {c.lgart} na MESMA linha PPDIT: {'SIM' if same else 'NÃO'}")
    print("")
    if tr.warnings:
        print("AVISOS:")
        for w in tr.warnings:
            print(f"  - {w}")
    print(L)
