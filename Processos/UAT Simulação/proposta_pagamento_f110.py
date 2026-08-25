# -*- coding: utf-8 -*-
"""
Agenda e executa a proposta de pagamento F110 via RFC/XBP.
"""

from __future__ import annotations

import argparse
import re
import time
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from typing import Any

from f110_uat_common import (
    load_project_dotenv,
    normalize_system_key,
    open_rfc_connection,
    parse_yyyymmdd,
    read_table,
    read_table_with_fallbacks,
    zero_pad_if_numeric,
)


# =============================================================================
# (1) CONSTANTES E MODELOS
# =============================================================================

ALLOWED_SYSTEM_IDS = {"QAD", "S4Q"}
ALLOWED_SYSTEM_KEYS = {"QAD", "S4Q", "S4QCLNT100"}
DEFAULT_JOBCLASS = "C"
DEFAULT_PAYMENT_METHOD = "S"
DEFAULT_PROPOSAL_ONLY = "X"
DEFAULT_DOC_DATE = date.today().strftime("%Y%m%d")
DEFAULT_RUN_DATE = date.today().strftime("%Y%m%d")
DEFAULT_DOCS_ENTERED_UP_TO = (date.today() + timedelta(days=1)).strftime("%Y%m%d")
DEFAULT_IDENTIFICATION = "AUTO"
DEFAULT_REPORT = "RFF110S"
IDENTIFICATION_PREFIX = "UAT"
IDENTIFICATION_MIN = 1
IDENTIFICATION_MAX = 99
DEFAULT_ACCOUNT_RANGE_LOW = "10000000"
DEFAULT_ACCOUNT_RANGE_HIGH = "99999999"
_TRUE_VALUES = {"1", "true", "yes", "on", "sim", "s", "x"}


def _today_yyyymmdd() -> str:
    return date.today().strftime("%Y%m%d")


def _next_yyyymmdd() -> str:
    return (date.today() + timedelta(days=1)).strftime("%Y%m%d")


def _now_hhmmss() -> str:
    return datetime.now().strftime("%H%M%S")


def _format_bapi_return(message: dict[str, Any]) -> str:
    msg_type = str(message.get("TYPE") or message.get("MSGTYPE") or "").strip().upper() or "?"
    msg_id = str(message.get("ID") or message.get("MSGID") or "").strip()
    msg_no = str(message.get("NUMBER") or message.get("MSGNO") or "").strip()
    text = str(message.get("MESSAGE") or message.get("TEXT") or "").strip()
    parts = [part for part in [msg_type, msg_id, msg_no, text] if part]
    return " | ".join(parts)


def _has_fatal_return(message: dict[str, Any]) -> bool:
    msg_type = str(message.get("TYPE") or message.get("MSGTYPE") or "").strip().upper()
    return msg_type in {"A", "E", "X"}


def _coerce_bool(value: Any, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if not text:
        return default
    if text in _TRUE_VALUES:
        return True
    if text in {"0", "false", "no", "off", "nao", "não", "n"}:
        return False
    return default


@dataclass(frozen=True)
class ProposalInput:
    system_key: str
    company_code: str
    vendor: str
    document_number: str
    fiscal_year: str
    run_date: str
    identification: str
    posting_date: str
    docs_entered_up_to: str
    payment_method: str = DEFAULT_PAYMENT_METHOD
    proposal_only: bool = True
    jobclass: str = DEFAULT_JOBCLASS
    wait_seconds: int = 120


@dataclass
class ProposalEvidence:
    name: str
    status: str
    details: str
    rows: list[dict[str, str]] = field(default_factory=list)


@dataclass
class ProposalResult:
    status: str
    scheduled: bool
    finished: bool
    system_id: str
    client: str
    jobname: str
    jobcount: str
    run_date: str
    identification: str
    company_code: str
    vendor: str
    document_number: str
    fiscal_year: str
    posting_date: str
    docs_entered_up_to: str
    payment_method: str
    proposal_only: bool
    payment_document_number: str = ""
    reasons: list[str] = field(default_factory=list)
    evidences: list[ProposalEvidence] = field(default_factory=list)
    job_status: str = ""
    joblog: list[str] = field(default_factory=list)


# =============================================================================
# (2) UTILITARIOS RFC/XBP
# =============================================================================

class F110ProposalRunner:
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

    def _add_evidence(self, evidences: list[ProposalEvidence], name: str, status: str, details: str, rows: list[dict[str, str]] | None = None) -> None:
        evidences.append(ProposalEvidence(name=name, status=status, details=details, rows=list(rows or [])))

    def _next_identification(self, run_date: str) -> str:
        rows = read_table(
            self.conn,
            "REGUV",
            ["LAUFI"],
            options=[f"LAUFD = '{run_date}'"],
            rowcount=200,
        )
        used_sequences: set[int] = set()
        for row in rows:
            laufi = str(row.get("LAUFI") or "").strip().upper()
            match = re.fullmatch(rf"{IDENTIFICATION_PREFIX}(\d{{1,2}})", laufi)
            if match:
                seq = int(match.group(1))
                if IDENTIFICATION_MIN <= seq <= IDENTIFICATION_MAX:
                    used_sequences.add(seq)

        for seq in range(IDENTIFICATION_MIN, IDENTIFICATION_MAX + 1):
            if seq not in used_sequences:
                return f"{IDENTIFICATION_PREFIX}{seq:02d}"

        raise RuntimeError(
            f"Nao foi possivel gerar uma identificacao sequencial livre para {IDENTIFICATION_PREFIX}01-{IDENTIFICATION_PREFIX}99."
        )

    def _resolve_identification(self, payload: ProposalInput) -> str:
        requested = str(payload.identification or "").strip().upper()
        if not requested or requested == DEFAULT_IDENTIFICATION:
            return self._next_identification(payload.run_date)
        return requested

    def _build_selinfo(self, payload: ProposalInput) -> list[dict[str, str]]:
        company_code = str(payload.company_code).strip()
        vendor_no = zero_pad_if_numeric(payload.vendor)
        payment_method = str(payload.payment_method or DEFAULT_PAYMENT_METHOD).strip().upper()
        run_date = str(payload.run_date or _today_yyyymmdd()).strip()
        posting_date = str(payload.posting_date or _today_yyyymmdd()).strip()
        docs_entered_up_to = str(payload.docs_entered_up_to or _next_yyyymmdd()).strip()
        document_number = zero_pad_if_numeric(payload.document_number)

        return [
            {"SELNAME": "SEL_BUKR", "KIND": "S", "SIGN": "I", "OPTION": "EQ", "LOW": company_code, "HIGH": ""},
            {"SELNAME": "SEL_KRED", "KIND": "S", "SIGN": "I", "OPTION": "EQ", "LOW": vendor_no, "HIGH": ""},
            {"SELNAME": "SEL_KREP-LOW", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": DEFAULT_ACCOUNT_RANGE_LOW, "HIGH": ""},
            {"SELNAME": "SEL_KREP-HIGH", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": DEFAULT_ACCOUNT_RANGE_HIGH, "HIGH": ""},
            {"SELNAME": "SEL_DEBP-LOW", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": DEFAULT_ACCOUNT_RANGE_LOW, "HIGH": ""},
            {"SELNAME": "SEL_DEBP-HIGH", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": DEFAULT_ACCOUNT_RANGE_HIGH, "HIGH": ""},
            {"SELNAME": "PAR_ZWE", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": payment_method, "HIGH": ""},
            {"SELNAME": "PAR_XVL", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "X" if payload.proposal_only else "", "HIGH": ""},
            {"SELNAME": "PAR_LFD", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": run_date, "HIGH": ""},
            {"SELNAME": "PAR_LFID", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": str(payload.identification or DEFAULT_IDENTIFICATION).strip().upper(), "HIGH": ""},
            {"SELNAME": "PAR_NEDA", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": docs_entered_up_to, "HIGH": ""},
            {"SELNAME": "PAR_BUDA", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": posting_date, "HIGH": ""},
            {"SELNAME": "PAR_GRDA", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": docs_entered_up_to, "HIGH": ""},
            {"SELNAME": "PAR_XFA", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "X", "HIGH": ""},
            {"SELNAME": "PAR_XZE", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "X", "HIGH": ""},
            {"SELNAME": "PAR_XBL", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "X", "HIGH": ""},
            {"SELNAME": "PAR_MITD", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "X", "HIGH": ""},
            {"SELNAME": "PAR_TEX1", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": "BKPF-BELNR", "HIGH": ""},
            {"SELNAME": "PAR_LIS1", "KIND": "P", "SIGN": "I", "OPTION": "EQ", "LOW": document_number, "HIGH": ""},
        ]

    def _job_message(self, result: dict[str, Any]) -> dict[str, Any]:
        return dict(result.get("RETURN") or {})

    def _call_open_job(self, jobname: str, jobclass: str) -> tuple[str, dict[str, Any]]:
        result = self.conn.call(
            "BAPI_XBP_JOB_OPEN",
            JOBNAME=jobname,
            JOBCLASS=jobclass,
            EXTERNAL_USER_NAME=self.user,
        )
        return str(result.get("JOBCOUNT") or "").strip(), self._job_message(result)

    def _call_add_step(self, jobname: str, jobcount: str, selinfo: list[dict[str, str]]) -> tuple[dict[str, Any], int]:
        result = self.conn.call(
            "BAPI_XBP_JOB_ADD_ABAP_STEP",
            JOBNAME=jobname,
            JOBCOUNT=jobcount,
            ABAP_PROGRAM_NAME=DEFAULT_REPORT,
            EXTERNAL_USER_NAME=self.user,
            LANGUAGE="P",
            SELINFO=selinfo,
        )
        return self._job_message(result), int(result.get("STEP_NUMBER") or 0)

    def _call_close_job(self, jobname: str, jobcount: str) -> dict[str, Any]:
        result = self.conn.call(
            "BAPI_XBP_JOB_CLOSE",
            JOBNAME=jobname,
            JOBCOUNT=jobcount,
            EXTERNAL_USER_NAME=self.user,
        )
        return self._job_message(result)

    def _call_start_job(self, jobname: str, jobcount: str) -> dict[str, Any]:
        result = self.conn.call(
            "BAPI_XBP_JOB_START_IMMEDIATELY",
            JOBNAME=jobname,
            JOBCOUNT=jobcount,
            EXTERNAL_USER_NAME=self.user,
        )
        return self._job_message(result)

    def _call_job_status(self, jobname: str, jobcount: str) -> str:
        result = self.conn.call(
            "BAPI_XBP_JOB_STATUS_GET",
            JOBNAME=jobname,
            JOBCOUNT=jobcount,
            EXTERNAL_USER_NAME=self.user,
        )
        return str(result.get("STATUS") or "").strip().upper()

    def _read_joblog(self, jobname: str, jobcount: str, max_lines: int = 50) -> list[str]:
        result = self.conn.call(
            "BAPI_XBP_JOB_JOBLOG_READ",
            JOBNAME=jobname,
            JOBCOUNT=jobcount,
            EXTERNAL_USER_NAME=self.user,
            LINES=max_lines,
            PROT_NEW="X",
        )
        rows = result.get("JOB_PROTOCOL_NEW") or result.get("JOB_PROTOCOL") or []
        lines: list[str] = []
        for row in rows:
            row_dict = dict(row)
            text = str(row_dict.get("TEXT") or "").strip()
            msgid = str(row_dict.get("MSGID") or "").strip()
            msgno = str(row_dict.get("MSGNO") or "").strip()
            msgtype = str(row_dict.get("MSGTYPE") or "").strip().upper()
            if text:
                lines.append(f"[{msgtype or '?'}] {msgid} {msgno} {text}".strip())
        return lines

    def _read_payment_document_number(self, payload: ProposalInput) -> str:
        belnr = zero_pad_if_numeric(payload.document_number)
        company_code = str(payload.company_code).strip()
        fiscal_year = str(payload.fiscal_year).strip()
        vendor_no = zero_pad_if_numeric(payload.vendor)

        for table_name in ("BSAK", "BSEG"):
            try:
                rows, _ = read_table_with_fallbacks(
                    self.conn,
                    table_name,
                    [["BUKRS", "BELNR", "GJAHR", "LIFNR", "AUGBL"], ["BUKRS", "BELNR", "GJAHR", "AUGBL"], ["BUKRS", "BELNR", "GJAHR"]],
                    options=[
                        f"BUKRS = '{company_code}'",
                        f"AND BELNR = '{belnr}'",
                        f"AND GJAHR = '{fiscal_year}'",
                        f"AND LIFNR = '{vendor_no}'",
                    ],
                    rowcount=20,
                )
            except Exception:
                continue

            for row in rows:
                payment_doc = zero_pad_if_numeric(row.get("AUGBL"), 10)
                if payment_doc:
                    return payment_doc

        return ""

    def run(self, payload: ProposalInput) -> ProposalResult:
        evidences: list[ProposalEvidence] = []
        reasons: list[str] = []
        job_status = ""
        joblog: list[str] = []
        payment_document_number = ""
        jobname = f"Z_F110_{payload.company_code}_{str(payload.identification or DEFAULT_IDENTIFICATION).strip().upper()}_{_now_hhmmss()}"
        jobcount = ""
        finished = False
        scheduled = False

        try:
            if normalize_system_key(payload.system_key) not in ALLOWED_SYSTEM_KEYS:
                reasons.append(f"Chave de sistema '{payload.system_key}' fora do escopo permitido para UAT F110.")
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id="",
                    client="",
                    jobname=jobname,
                    jobcount="",
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    payment_document_number="",
                    reasons=reasons,
                    evidences=evidences,
                )

            self.connect(payload.system_key)
            payload = replace(payload, identification=self._resolve_identification(payload))
            jobname = f"Z_F110_{payload.company_code}_{payload.identification}_{_now_hhmmss()}"

            if self.system_id not in ALLOWED_SYSTEM_IDS:
                reasons.append(
                    f"Execucao bloqueada: sistema ligado '{self.system_id or '?'}' fora do escopo permitido para UAT F110."
                )
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id=self.system_id,
                    client=self.client,
                    jobname=jobname,
                    jobcount="",
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    payment_document_number="",
                    reasons=reasons,
                    evidences=evidences,
                )

            jobcount, open_return = self._call_open_job(jobname, payload.jobclass)
            if open_return:
                self._add_evidence(evidences, "BAPI_XBP_JOB_OPEN", "ok" if not _has_fatal_return(open_return) else "erro", "Job aberto.", [open_return])
            if _has_fatal_return(open_return):
                reasons.append(_format_bapi_return(open_return))
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id=self.system_id,
                    client=self.client,
                    jobname=jobname,
                    jobcount=jobcount,
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    payment_document_number="",
                    reasons=reasons,
                    evidences=evidences,
                )

            selinfo = self._build_selinfo(payload)
            add_return, step_number = self._call_add_step(jobname, jobcount, selinfo)
            self._add_evidence(
                evidences,
                "BAPI_XBP_JOB_ADD_ABAP_STEP",
                "ok" if not _has_fatal_return(add_return) else "erro",
                f"Step {step_number} criado para {DEFAULT_REPORT}.",
                [add_return],
            )
            if _has_fatal_return(add_return):
                reasons.append(_format_bapi_return(add_return))
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id=self.system_id,
                    client=self.client,
                    jobname=jobname,
                    jobcount=jobcount,
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    payment_document_number="",
                    reasons=reasons,
                    evidences=evidences,
                )

            close_return = self._call_close_job(jobname, jobcount)
            self._add_evidence(
                evidences,
                "BAPI_XBP_JOB_CLOSE",
                "ok" if not _has_fatal_return(close_return) else "erro",
                "Job fechado.",
                [close_return],
            )
            if _has_fatal_return(close_return):
                reasons.append(_format_bapi_return(close_return))
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id=self.system_id,
                    client=self.client,
                    jobname=jobname,
                    jobcount=jobcount,
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    reasons=reasons,
                    evidences=evidences,
                )

            start_return = self._call_start_job(jobname, jobcount)
            self._add_evidence(
                evidences,
                "BAPI_XBP_JOB_START_IMMEDIATELY",
                "ok" if not _has_fatal_return(start_return) else "erro",
                "Job enviado para execucao.",
                [start_return],
            )
            if _has_fatal_return(start_return):
                reasons.append(_format_bapi_return(start_return))
                return ProposalResult(
                    status="blocked",
                    scheduled=False,
                    finished=False,
                    system_id=self.system_id,
                    client=self.client,
                    jobname=jobname,
                    jobcount=jobcount,
                    run_date=payload.run_date,
                    identification=payload.identification,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    document_number=payload.document_number,
                    fiscal_year=payload.fiscal_year,
                    posting_date=payload.posting_date,
                    docs_entered_up_to=payload.docs_entered_up_to,
                    payment_method=payload.payment_method,
                    proposal_only=payload.proposal_only,
                    reasons=reasons,
                    evidences=evidences,
                )

            scheduled = True
            deadline = time.time() + max(5, int(payload.wait_seconds))
            while time.time() < deadline:
                job_status = self._call_job_status(jobname, jobcount)
                if job_status in {"F", "A", "C", "S"}:
                    break
                time.sleep(2)

            if job_status in {"F", "C"}:
                finished = True
            elif job_status in {"A", "X"}:
                reasons.append(f"Job terminou com status {job_status}.")
            else:
                reasons.append(f"Job nao terminou dentro do tempo de espera ({payload.wait_seconds}s).")

            joblog = self._read_joblog(jobname, jobcount)
            if joblog:
                self._add_evidence(evidences, "JOBLOG", "ok", "Log do job lido.", [{"TEXT": line} for line in joblog[:20]])

            if not payload.proposal_only:
                payment_document_number = self._read_payment_document_number(payload)
                if payment_document_number:
                    self._add_evidence(
                        evidences,
                        "PAYMENT_DOCUMENT",
                        "ok",
                        "Documento de pagamento localizado em BSAK/BSEG.",
                        [{"AUGBL": payment_document_number}],
                    )

            status = "finished" if finished else "scheduled"
            return ProposalResult(
                status=status,
                scheduled=scheduled,
                finished=finished,
                system_id=self.system_id,
                client=self.client,
                jobname=jobname,
                jobcount=jobcount,
                run_date=payload.run_date,
                identification=payload.identification,
                company_code=payload.company_code,
                vendor=payload.vendor,
                document_number=payload.document_number,
                fiscal_year=payload.fiscal_year,
                posting_date=payload.posting_date,
                docs_entered_up_to=payload.docs_entered_up_to,
                payment_method=payload.payment_method,
                proposal_only=payload.proposal_only,
                payment_document_number=payment_document_number,
                reasons=reasons,
                evidences=evidences,
                job_status=job_status,
                joblog=joblog,
            )
        except Exception as exc:
            reasons.append(str(exc))
            return ProposalResult(
                status="blocked",
                scheduled=scheduled,
                finished=finished,
                system_id=self.system_id,
                client=self.client,
                jobname=jobname,
                jobcount=jobcount,
                run_date=payload.run_date,
                identification=payload.identification,
                company_code=payload.company_code,
                vendor=payload.vendor,
                document_number=payload.document_number,
                fiscal_year=payload.fiscal_year,
                posting_date=payload.posting_date,
                docs_entered_up_to=payload.docs_entered_up_to,
                payment_method=payload.payment_method,
                proposal_only=payload.proposal_only,
                payment_document_number=payment_document_number,
                reasons=reasons,
                evidences=evidences,
                job_status=job_status,
                joblog=joblog,
            )
        finally:
            self.close()


# =============================================================================
# (3) CLI / FORMATO
# =============================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Executar a proposta de pagamento F110 via RFC/XBP.")
    parser.add_argument("--sap-system", default="QAD", help="Chave de sistema SAP. Padrao: QAD")
    parser.add_argument("--company-code", default="2010", help="Codigo da empresa (BUKRS)")
    parser.add_argument("--vendor", default="10000040", help="Fornecedor alvo")
    parser.add_argument("--document-number", default="6050000002", help="Numero do documento FI alvo")
    parser.add_argument("--fiscal-year", default="2026", help="Exercicio do documento")
    parser.add_argument("--run-date", default=DEFAULT_RUN_DATE, help="Data da execucao do job em YYYYMMDD")
    parser.add_argument("--identification", default=DEFAULT_IDENTIFICATION, help="Identificacao da F110")
    parser.add_argument("--posting-date", default=DEFAULT_DOC_DATE, help="Posting date em YYYYMMDD")
    parser.add_argument("--docs-entered-up-to", default=DEFAULT_DOCS_ENTERED_UP_TO, help="Data limite dos documentos em YYYYMMDD")
    parser.add_argument("--payment-method", default=DEFAULT_PAYMENT_METHOD, help="Metodo de pagamento")
    proposal_group = parser.add_mutually_exclusive_group()
    proposal_group.add_argument(
        "--proposal-only",
        dest="proposal_only",
        action="store_true",
        help="Executar apenas proposta",
    )
    proposal_group.add_argument(
        "--execute-payment",
        dest="proposal_only",
        action="store_false",
        help="Executar o lancamento de pagamento",
    )
    parser.set_defaults(proposal_only=True)
    parser.add_argument("--wait-seconds", type=int, default=120, help="Tempo maximo de espera pelo job")
    return parser


def format_result(result: ProposalResult) -> str:
    lines: list[str] = []
    lines.append("=" * 84)
    lines.append("UAT PROPOSTA PAGAMENTO F110")
    lines.append("=" * 84)
    lines.append(f"Estado         : {result.status.upper()}")
    lines.append(f"Run date       : {result.run_date}")
    lines.append(f"Identificacao  : {result.identification}")
    lines.append(f"Documento SAP utilizado : {result.document_number}")
    lines.append(f"Empresa        : {result.company_code}")
    lines.append(f"Fornecedor     : {result.vendor}")
    lines.append(f"Exercicio      : {result.fiscal_year}")
    if result.payment_document_number:
        lines.append(f"Documento de pagamento : {result.payment_document_number}")
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
        for row in evidence.rows[:5]:
            lines.append(f"  - {row}")
    if result.joblog:
        lines.append("")
        lines.append("Joblog")
        lines.append("-" * 84)
        for line in result.joblog[:20]:
            lines.append(line)
    return "\n".join(lines)


def executar(**kwargs: Any) -> ProposalResult:
    load_project_dotenv()
    payload = ProposalInput(
        system_key=str(kwargs.get("system_key") or kwargs.get("sap_system") or "QAD").strip(),
        company_code=str(kwargs.get("company_code") or "2010").strip(),
        vendor=str(kwargs.get("vendor") or "10000040").strip(),
        document_number=str(kwargs.get("document_number") or "6050000002").strip(),
        fiscal_year=str(kwargs.get("fiscal_year") or "2026").strip(),
        run_date=str(kwargs.get("run_date") or _today_yyyymmdd()).strip(),
        identification=str(kwargs.get("identification") or DEFAULT_IDENTIFICATION).strip(),
        posting_date=str(kwargs.get("posting_date") or _today_yyyymmdd()).strip(),
        docs_entered_up_to=str(kwargs.get("docs_entered_up_to") or _next_yyyymmdd()).strip(),
        payment_method=str(kwargs.get("payment_method") or DEFAULT_PAYMENT_METHOD).strip().upper(),
        proposal_only=_coerce_bool(kwargs.get("proposal_only"), default=True),
        jobclass=str(kwargs.get("jobclass") or DEFAULT_JOBCLASS).strip().upper(),
        wait_seconds=int(kwargs.get("wait_seconds") or 120),
    )
    runner = F110ProposalRunner()
    result = runner.run(payload)
    print(format_result(result))
    return result


def main(argv: list[str] | None = None) -> int:
    load_project_dotenv()
    args = build_parser().parse_args(argv)
    payload = ProposalInput(
        system_key=args.sap_system,
        company_code=args.company_code,
        vendor=args.vendor,
        document_number=args.document_number,
        fiscal_year=args.fiscal_year,
        run_date=parse_yyyymmdd(args.run_date) or args.run_date,
        identification=args.identification,
        posting_date=parse_yyyymmdd(args.posting_date) or args.posting_date,
        docs_entered_up_to=parse_yyyymmdd(args.docs_entered_up_to) or args.docs_entered_up_to,
        payment_method=args.payment_method,
        proposal_only=_coerce_bool(args.proposal_only, default=True),
        jobclass=DEFAULT_JOBCLASS,
        wait_seconds=int(args.wait_seconds),
    )
    runner = F110ProposalRunner()
    result = runner.run(payload)
    print(format_result(result))
    return 0 if result.finished or result.scheduled else 1


if __name__ == "__main__":
    raise SystemExit(main())
