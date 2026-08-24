# -*- coding: utf-8 -*-
"""
Orquestrador da simulacao UAT da F110/RFF110S.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from f110_uat_common import (
    call_bapi_ap_acc_getopenitems,
    load_project_dotenv,
    normalize_system_key,
    open_rfc_connection,
    read_table_with_fallbacks,
    zero_pad_if_numeric,
)


# =============================================================================
# (1) MODELOS DE DADOS
# =============================================================================

def _default_proposal_date() -> str:
    return date.today().strftime("%Y%m%d")


ALLOWED_SYSTEM_IDS = {"QAD", "S4Q"}
ALLOWED_SYSTEM_KEYS = {"QAD", "S4Q", "S4QCLNT100"}


@dataclass(frozen=True)
class SimulationInput:
    system_key: str
    company_code: str
    vendor: str
    document_number: str
    fiscal_year: str
    proposal_date: str = ""
    payment_method: str = "S"


@dataclass
class SimulationEvidence:
    name: str
    status: str
    details: str
    rows: list[dict[str, str]] = field(default_factory=list)


@dataclass
class SimulationResult:
    status: str
    eligible: bool
    inconclusive: bool
    system_id: str
    client: str
    company_code: str
    vendor: str
    document_number: str
    fiscal_year: str
    proposal_date: str
    payment_method: str
    reasons: list[str] = field(default_factory=list)
    evidences: list[SimulationEvidence] = field(default_factory=list)


# =============================================================================
# (2) SIMULADOR RFF110S
# =============================================================================

class F110UATSimulator:
    def __init__(self) -> None:
        self.conn = None
        self.user = ""
        self.system_id = ""
        self.client = ""

    def connect(self, system_key: str) -> None:
        self.conn, self.user = open_rfc_connection(system_key)
        self.conn.call("RFC_PING")
        info = self.conn.call("RFC_SYSTEM_INFO").get("RFCSI_EXPORT", {}) or {}
        self.system_id = str(info.get("RFCSYSID") or info.get("RFCSYSTEMID") or "").strip().upper()
        self.client = str(info.get("RFCCLIENT") or "").strip()

    def close(self) -> None:
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass

    def _add_evidence(self, evidences: list[SimulationEvidence], name: str, status: str, details: str, rows: list[dict[str, str]] | None = None) -> None:
        evidences.append(SimulationEvidence(name=name, status=status, details=details, rows=list(rows or [])))

    def _evaluate_payment_method(self, open_item: dict[str, str], vendor_row: dict[str, str], requested_method: str) -> tuple[bool | None, str]:
        row_method = str(open_item.get("ZLSCH", "")).strip().upper()
        vendor_methods = str(vendor_row.get("ZWELS", "")).strip().upper()
        requested = str(requested_method or "").strip().upper()

        if row_method:
            if row_method == requested:
                return True, f"Método de pagamento do item aberto já está definido como {row_method}."
            return False, f"Método de pagamento do item aberto é {row_method}, não {requested}."

        if vendor_methods:
            if requested in vendor_methods:
                return True, f"Método {requested} permitido pelo mestre do fornecedor ({vendor_methods})."
            return False, f"Método {requested} não aparece nos métodos permitidos do fornecedor ({vendor_methods})."

        return None, "Não foi possível confirmar o método de pagamento no item aberto nem no mestre do fornecedor."

    def _evaluate_due_date(self, open_item: dict[str, str], proposal_date: str) -> tuple[bool | None, str]:
        due_date = str(open_item.get("BLINE_DATE") or open_item.get("FAEDT") or "").strip()
        if not due_date:
            return None, "Data de vencimento não foi encontrada no item aberto."
        if due_date <= proposal_date:
            return True, f"Item vencido ou devido até a data da proposta ({due_date} <= {proposal_date})."
        return False, f"Item ainda não venceu para a proposta ({due_date} > {proposal_date})."

    def run(self, payload: SimulationInput) -> SimulationResult:
        evidences: list[SimulationEvidence] = []
        reasons: list[str] = []
        eligible = False
        inconclusive = False
        status = "inconclusive"

        try:
            self.connect(payload.system_key)

            if self.system_id not in ALLOWED_SYSTEM_IDS:
                reasons.append(
                    f"Execução bloqueada: sistema ligado '{self.system_id or '?'}' fora do escopo permitido para UAT F110/RFF110S."
                )
                return SimulationResult(
                    status="blocked",
                    eligible=False,
                    inconclusive=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    proposal_date=payload.proposal_date,
                    payment_method=payload.payment_method,
                    reasons=reasons,
                    evidences=evidences,
                )

            if normalize_system_key(payload.system_key) not in ALLOWED_SYSTEM_KEYS:
                reasons.append("A chave de sistema informada foi normalizada, mas a ligação foi aceite pelo system id do SAP.")

            # --------------------------------------------------------------
            # Documento FI no cabeçalho
            # --------------------------------------------------------------
            bkpf_rows = []
            try:
                bkpf_rows = read_table_with_fallbacks(
                    self.conn,
                    "BKPF",
                    [
                        ["BUKRS", "BELNR", "GJAHR", "BLART", "BUDAT", "CPUDT", "USNAM", "TCODE", "XBLNR", "WAERS", "BKTXT"],
                        ["BUKRS", "BELNR", "GJAHR", "BLART", "BUDAT", "USNAM", "TCODE"],
                    ],
                    options=[
                        f"BUKRS = '{payload.company_code}'",
                        f"AND BELNR = '{zero_pad_if_numeric(payload.document_number)}'",
                        f"AND GJAHR = '{payload.fiscal_year}'",
                    ],
                    rowcount=5,
                )[0]
                if bkpf_rows:
                    self._add_evidence(evidences, "BKPF", "ok", "Documento FI localizado no cabeçalho.", bkpf_rows)
                else:
                    self._add_evidence(evidences, "BKPF", "aviso", "Documento FI não foi encontrado na BKPF.")
                    reasons.append("Documento não encontrado na BKPF.")
                    inconclusive = True
            except Exception as exc:
                self._add_evidence(evidences, "BKPF", "erro", f"Não foi possível ler a BKPF: {exc}")
                reasons.append("Falha ao consultar BKPF.")
                inconclusive = True

            # --------------------------------------------------------------
            # Segmento do fornecedor
            # --------------------------------------------------------------
            lfb1_rows = []
            try:
                lfb1_rows = read_table_with_fallbacks(
                    self.conn,
                    "LFB1",
                    [
                        ["BUKRS", "LIFNR", "ZTERM", "ZWELS"],
                        ["BUKRS", "LIFNR", "ZTERM"],
                    ],
                    options=[
                        f"BUKRS = '{payload.company_code}'",
                        f"AND LIFNR = '{zero_pad_if_numeric(payload.vendor)}'",
                    ],
                    rowcount=5,
                )[0]
                if lfb1_rows:
                    self._add_evidence(evidences, "LFB1", "ok", "Segmento da empresa do fornecedor localizado.", lfb1_rows)
                else:
                    self._add_evidence(evidences, "LFB1", "aviso", "Não foi encontrado o segmento do fornecedor na LFB1.")
                    reasons.append("Fornecedor/empresa não encontrado na LFB1.")
                    inconclusive = True
            except Exception as exc:
                self._add_evidence(evidences, "LFB1", "erro", f"Não foi possível ler a LFB1: {exc}")
                reasons.append("Falha ao consultar LFB1.")
                inconclusive = True

            # --------------------------------------------------------------
            # Item em aberto do fornecedor
            # --------------------------------------------------------------
            open_item_rows = []
            open_item_error = ""
            try:
                open_item_rows, raw_result = call_bapi_ap_acc_getopenitems(
                    self.conn,
                    company_code=payload.company_code,
                    vendor=zero_pad_if_numeric(payload.vendor),
                    keydate=payload.proposal_date,
                )
                if open_item_rows:
                    field_signature = ",".join(sorted(open_item_rows[0].keys()))
                    filtered_rows = [
                        row for row in open_item_rows
                        if zero_pad_if_numeric(row.get("DOC_NO"), 10) == zero_pad_if_numeric(payload.document_number)
                        and str(row.get("FISC_YEAR", "")).strip() == payload.fiscal_year
                        and str(row.get("COMP_CODE", "")).strip() == payload.company_code
                    ]
                    if filtered_rows:
                        open_item_rows = filtered_rows
                    else:
                        reasons.append("A BAPI devolveu itens em aberto, mas nenhum coincidiu com o documento criado.")
                    self._add_evidence(
                        evidences,
                        "BAPI_AP_ACC_GETOPENITEMS",
                        "ok",
                        f"Itens em aberto devolvidos pela BAPI (campos: {field_signature}).",
                        open_item_rows,
                    )
                else:
                    open_item_error = "BAPI_AP_ACC_GETOPENITEMS: sem linhas retornadas."
            except Exception as exc:
                open_item_error = f"BAPI_AP_ACC_GETOPENITEMS: {exc}"

            if not open_item_rows:
                self._add_evidence(
                    evidences,
                    "BAPI_AP_ACC_GETOPENITEMS",
                    "aviso",
                    f"Não foi possível localizar item em aberto com os critérios informados. {open_item_error}".strip(),
                )
                reasons.append("Não foi possível confirmar o item em aberto para a proposta.")
                inconclusive = True

            # --------------------------------------------------------------
            # Avaliação de elegibilidade
            # --------------------------------------------------------------
            if bkpf_rows and lfb1_rows and open_item_rows:
                open_item = open_item_rows[0]
                vendor_row = lfb1_rows[0]

                payment_ok, payment_reason = self._evaluate_payment_method(open_item, vendor_row, payload.payment_method)
                if payment_ok is True:
                    reasons.append(payment_reason)
                elif payment_ok is False:
                    reasons.append(payment_reason)
                    eligible = False
                    status = "blocked"
                else:
                    reasons.append(payment_reason)
                    inconclusive = True

                due_ok, due_reason = self._evaluate_due_date(open_item, payload.proposal_date)
                if due_ok is True:
                    reasons.append(due_reason)
                elif due_ok is False:
                    reasons.append(due_reason)
                    eligible = False
                    status = "blocked"
                else:
                    reasons.append(due_reason)
                    inconclusive = True

                payment_block = str(open_item.get("ZLSPR", "")).strip()
                if payment_block:
                    reasons.append(f"Item possui bloqueio de pagamento '{payment_block}'.")
                    eligible = False
                    status = "blocked"

                if payment_ok is True and due_ok is True and not payment_block:
                    eligible = True
                    status = "eligible"
                elif status != "blocked":
                    status = "inconclusive"

            if inconclusive and not eligible and status != "blocked":
                status = "inconclusive"

            if not eligible and not inconclusive and status != "blocked":
                status = "blocked"

            return SimulationResult(
                status=status,
                eligible=eligible,
                inconclusive=inconclusive,
                system_id=self.system_id,
                client=self.client,
                company_code=payload.company_code,
                vendor=payload.vendor,
                document_number=payload.document_number,
                fiscal_year=payload.fiscal_year,
                proposal_date=payload.proposal_date,
                payment_method=payload.payment_method,
                reasons=reasons,
                evidences=evidences,
            )
        finally:
            self.close()


# =============================================================================
# (3) CLI / ENTRYPOINT
# =============================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Simular elegibilidade de documento FI para proposta RFF110S.")
    parser.add_argument("--sap-system", default="QAD", help="Chave de sistema SAP. Padrao: QAD")
    parser.add_argument("--company-code", default="2010", help="Codigo da empresa (BUKRS)")
    parser.add_argument("--vendor", default="0010000040", help="Numero do fornecedor")
    parser.add_argument(
        "--belnr",
        "--document-number",
        dest="document_number",
        default="6050000001",
        help="Numero do documento FI / BKPF-BELNR usado na selecao RFF110S",
    )
    parser.add_argument("--fiscal-year", default="2026", help="Exercicio do documento")
    parser.add_argument("--proposal-date", default=_default_proposal_date(), help="Data de proposta RFF110S em YYYYMMDD")
    parser.add_argument("--payment-method", default="S", help="Metodo de pagamento esperado")
    return parser


def format_result(result: SimulationResult) -> str:
    lines = []
    lines.append("=" * 84)
    lines.append("UAT SIMULACAO RFF110S")
    lines.append("=" * 84)
    lines.append(f"Sistema        : {result.system_id or '?'}")
    lines.append(f"Cliente        : {result.client or '?'}")
    lines.append(f"Empresa        : {result.company_code}")
    lines.append(f"Fornecedor     : {result.vendor}")
    lines.append(f"BKPF-BELNR     : {result.document_number}")
    lines.append(f"Exercicio      : {result.fiscal_year}")
    lines.append(f"Data proposta  : {result.proposal_date}")
    lines.append(f"Metodo esp.    : {result.payment_method}")
    lines.append(f"Selecao RFF110S: BKPF-BELNR = {result.document_number}")
    lines.append(f"Estado         : {result.status.upper()}")
    lines.append(f"Apto F110      : {'SIM' if result.eligible else 'NAO'}")
    lines.append(f"Inconclusivo   : {'SIM' if result.inconclusive else 'NAO'}")
    lines.append("")
    lines.append("Motivos")
    lines.append("-" * 84)
    if result.reasons:
        for index, reason in enumerate(result.reasons, start=1):
            lines.append(f"{index}. {reason}")
    else:
        lines.append("Sem observacoes adicionais.")
    lines.append("")
    lines.append("Evidencias")
    lines.append("-" * 84)
    for evidence in result.evidences:
        lines.append(f"[{evidence.status.upper()}] {evidence.name}: {evidence.details}")
    return "\n".join(lines)


def executar(ambiente_cockpit: str | None = None, **kwargs: Any) -> SimulationResult:
    load_project_dotenv()
    payload = SimulationInput(
        system_key=str(kwargs.get("system_key") or ambiente_cockpit or kwargs.get("sap_system") or "QAD").strip(),
        company_code=str(kwargs.get("company_code") or "2010").strip(),
        vendor=str(kwargs.get("vendor") or "0010000040").strip(),
        document_number=str(kwargs.get("document_number") or "6050000001").strip(),
        fiscal_year=str(kwargs.get("fiscal_year") or "2026").strip(),
        proposal_date=str(kwargs.get("proposal_date") or _default_proposal_date()).strip(),
        payment_method=str(kwargs.get("payment_method") or "S").strip().upper(),
    )
    simulator = F110UATSimulator()
    result = simulator.run(payload)
    print(format_result(result))
    return result


def main(argv: list[str] | None = None) -> int:
    load_project_dotenv()
    args = build_parser().parse_args(argv)
    simulator = F110UATSimulator()
    payload = SimulationInput(
        system_key=args.sap_system,
        company_code=args.company_code,
        vendor=args.vendor,
        document_number=args.document_number,
        fiscal_year=args.fiscal_year,
        proposal_date=args.proposal_date,
        payment_method=args.payment_method,
    )
    result = simulator.run(payload)
    print(format_result(result))
    return 0 if result.eligible or result.inconclusive else 1


if __name__ == "__main__":
    raise SystemExit(main())
