# -*- coding: utf-8 -*-
"""
Cria um documento FI de fornecedor via RFC/BAPI para teste da F110.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any

from f110_uat_common import (
    load_project_dotenv,
    normalize_system_key,
    open_rfc_connection,
    read_table_with_fallbacks,
    zero_pad_if_numeric,
)


# =============================================================================
# (1) CONSTANTES E MODELOS
# =============================================================================

ALLOWED_SYSTEM_IDS = {"QAD", "S4Q"}
ALLOWED_SYSTEM_KEYS = {"QAD", "S4Q", "S4QCLNT100"}
DEFAULT_DOC_TYPE = "KR"
DEFAULT_BUS_ACT = "RFBU"
DEFAULT_REFERENCE = "UAT-F110-TEST"
XBLNR_MAX_LENGTH = 16
XBLNR_SEQUENCE_DIGITS = 2
DEFAULT_HEADER_TEXT = "UAT TESTE F110"
DEFAULT_ITEM_TEXT = "UAT F110 TESTE"
DEFAULT_CURRENCY = "EUR"


def _today_yyyymmdd() -> str:
    return date.today().strftime("%Y%m%d")


def _current_year() -> str:
    return date.today().strftime("%Y")


def _normalize_amount(value: str | Decimal | float | int) -> Decimal:
    if isinstance(value, Decimal):
        amount = value
    else:
        text = str(value).strip().replace(",", ".")
        try:
            amount = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError(f"Valor invalido para montante: {value!r}") from exc
    return amount.quantize(Decimal("0.01"))


def _decimal_to_bapi_text(value: Decimal) -> str:
    return format(value, ".2f")


def _line_no(index: int) -> str:
    return str(index).zfill(10)


def _bapi_messages(result: dict[str, Any]) -> list[dict[str, str]]:
    raw = result.get("RETURN", [])
    if isinstance(raw, dict):
        raw = [raw]
    messages: list[dict[str, str]] = []
    for row in raw or []:
        messages.append({key: str(value or "") for key, value in dict(row).items()})
    return messages


def _has_fatal_message(messages: list[dict[str, str]]) -> bool:
    return any(message.get("TYPE", "").upper() in {"A", "E", "X"} for message in messages)


@dataclass(frozen=True)
class DocumentInput:
    system_key: str
    company_code: str
    vendor: str
    gl_account: str
    amount: Decimal
    currency: str = DEFAULT_CURRENCY
    document_date: str = ""
    posting_date: str = ""
    baseline_date: str = ""
    payment_terms: str = ""
    payment_method: str = "S"
    doc_type: str = DEFAULT_DOC_TYPE
    reference: str = DEFAULT_REFERENCE
    header_text: str = DEFAULT_HEADER_TEXT
    item_text: str = DEFAULT_ITEM_TEXT
    check_only: bool = False


@dataclass
class DocumentEvidence:
    name: str
    status: str
    details: str
    rows: list[dict[str, str]] = field(default_factory=list)


@dataclass
class DocumentResult:
    status: str
    checked: bool
    posted: bool
    committed: bool
    system_id: str
    client: str
    company_code: str
    vendor: str
    gl_account: str
    amount: str
    currency: str
    document_date: str
    posting_date: str
    baseline_date: str
    payment_terms: str
    payment_method: str
    doc_type: str
    reference: str
    header_text: str
    item_text: str
    fiscal_year: str
    posted_belnr: str = ""
    posted_bukrs: str = ""
    posted_gjahr: str = ""
    obj_key: str = ""
    obj_type: str = ""
    obj_sys: str = ""
    reasons: list[str] = field(default_factory=list)
    evidences: list[DocumentEvidence] = field(default_factory=list)
    check_messages: list[dict[str, str]] = field(default_factory=list)
    post_messages: list[dict[str, str]] = field(default_factory=list)


# =============================================================================
# (2) HELPERS RFC / BAPI
# =============================================================================

class FiDocumentPoster:
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

    def _add_evidence(self, evidences: list[DocumentEvidence], name: str, status: str, details: str, rows: list[dict[str, str]] | None = None) -> None:
        evidences.append(DocumentEvidence(name=name, status=status, details=details, rows=list(rows or [])))

    def _read_company_currency(self, company_code: str) -> str:
        try:
            rows, _ = read_table_with_fallbacks(
                self.conn,
                "T001",
                [["BUKRS", "WAERS"], ["BUKRS"]],
                options=[f"BUKRS = '{company_code}'"],
                rowcount=1,
            )
        except Exception:
            return DEFAULT_CURRENCY
        if rows:
            return str(rows[0].get("WAERS") or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY
        return DEFAULT_CURRENCY

    def _read_vendor_defaults(self, company_code: str, vendor: str) -> dict[str, str]:
        try:
            rows, _ = read_table_with_fallbacks(
                self.conn,
                "LFB1",
                [["BUKRS", "LIFNR", "ZTERM", "ZWELS"], ["BUKRS", "LIFNR", "ZTERM"], ["BUKRS", "LIFNR"]],
                options=[
                    f"BUKRS = '{company_code}'",
                    f"AND LIFNR = '{vendor}'",
                ],
                rowcount=1,
            )
        except Exception:
            return {}
        return dict(rows[0]) if rows else {}

    def _next_reference(self, company_code: str, reference_base: str) -> str:
        base = str(reference_base or DEFAULT_REFERENCE).strip().upper() or DEFAULT_REFERENCE
        prefix = base[: XBLNR_MAX_LENGTH - XBLNR_SEQUENCE_DIGITS]
        rows, _ = read_table_with_fallbacks(
            self.conn,
            "BKPF",
            [["XBLNR"], ["BUKRS", "XBLNR"]],
            options=[
                f"BUKRS = '{company_code}'",
                f"AND XBLNR LIKE '{prefix}%'",
            ],
            rowcount=500,
        )
        highest_sequence = 0
        for row in rows:
            xblnr = str(row.get("XBLNR") or "").strip().upper()
            if not xblnr.startswith(prefix):
                continue
            suffix = xblnr[len(prefix) :]
            if len(suffix) != XBLNR_SEQUENCE_DIGITS or not suffix.isdigit():
                continue
            highest_sequence = max(highest_sequence, int(suffix))

        next_sequence = highest_sequence + 1
        if next_sequence > 10**XBLNR_SEQUENCE_DIGITS - 1:
            raise RuntimeError(
                f"Nao foi possivel gerar XBLNR sequencial para o prefixo {prefix!r}: limite de {XBLNR_SEQUENCE_DIGITS} digitos atingido."
            )
        return f"{prefix}{next_sequence:0{XBLNR_SEQUENCE_DIGITS}d}"

    @staticmethod
    def _build_document_header(payload: DocumentInput, username: str) -> dict[str, str]:
        return {
            "BUS_ACT": DEFAULT_BUS_ACT,
            "USERNAME": username,
            "HEADER_TXT": payload.header_text,
            "COMP_CODE": payload.company_code,
            "DOC_DATE": payload.document_date,
            "PSTNG_DATE": payload.posting_date,
            "FISC_YEAR": payload.posting_date[:4],
            "DOC_TYPE": payload.doc_type,
            "REF_DOC_NO": payload.reference,
        }

    @staticmethod
    def _build_tables(payload: DocumentInput, payment_terms: str, company_currency: str) -> dict[str, list[dict[str, str]]]:
        vendor_no = zero_pad_if_numeric(payload.vendor)
        gl_account = zero_pad_if_numeric(payload.gl_account)
        currency = (payload.currency or company_currency or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY
        amount_text = _decimal_to_bapi_text(payload.amount)
        credit_amount = f"-{amount_text}" if not amount_text.startswith("-") else amount_text
        debit_amount = amount_text.lstrip("-")
        baseline_date = str(payload.baseline_date or "").strip()
        payment_method = str(payload.payment_method or "").strip().upper()
        payment_terms = str(payment_terms or payload.payment_terms or "").strip().upper()

        account_payable = [
            {
                "ITEMNO_ACC": _line_no(1),
                "VENDOR_NO": vendor_no,
                "COMP_CODE": payload.company_code,
                "PMNTTRMS": payment_terms,
                "BLINE_DATE": baseline_date,
                "PYMT_METH": payment_method,
                "ALLOC_NMBR": payload.reference,
                "ITEM_TEXT": payload.item_text,
            }
        ]
        account_gl = [
            {
                "ITEMNO_ACC": _line_no(2),
                "GL_ACCOUNT": gl_account,
                "COMP_CODE": payload.company_code,
                "ITEM_TEXT": payload.item_text,
                "ALLOC_NMBR": payload.reference,
                "DE_CRE_IND": "S",
            }
        ]
        currency_amount = [
            {
                "ITEMNO_ACC": _line_no(1),
                "CURR_TYPE": "00",
                "CURRENCY": currency,
                "AMT_DOCCUR": credit_amount,
            },
            {
                "ITEMNO_ACC": _line_no(2),
                "CURR_TYPE": "00",
                "CURRENCY": currency,
                "AMT_DOCCUR": debit_amount,
            },
        ]
        return {
            "ACCOUNTGL": account_gl,
            "ACCOUNTPAYABLE": account_payable,
            "CURRENCYAMOUNT": currency_amount,
        }

    def _call_bapi(self, function_name: str, payload: DocumentInput, username: str, payment_terms: str, company_currency: str) -> dict[str, Any]:
        tables = self._build_tables(payload, payment_terms, company_currency)
        return self.conn.call(
            function_name,
            DOCUMENTHEADER=self._build_document_header(payload, username),
            ACCOUNTGL=tables["ACCOUNTGL"],
            ACCOUNTPAYABLE=tables["ACCOUNTPAYABLE"],
            CURRENCYAMOUNT=tables["CURRENCYAMOUNT"],
        )

    @staticmethod
    def _parse_obj_key(obj_key: str) -> tuple[str, str, str]:
        key = str(obj_key or "").strip()
        if len(key) >= 18:
            return key[:10], key[10:14], key[14:18]
        return "", "", ""

    @staticmethod
    def _format_messages(messages: list[dict[str, str]]) -> list[str]:
        lines: list[str] = []
        for message in messages:
            msg_type = message.get("TYPE", "").upper() or "?"
            msg_id = message.get("ID", "").strip()
            msg_no = message.get("NUMBER", "").strip()
            text = message.get("MESSAGE", "").strip()
            rows = [part for part in [msg_type, msg_id, msg_no, text] if part]
            lines.append(" | ".join(rows))
        return lines

    def run(self, payload: DocumentInput) -> DocumentResult:
        evidences: list[DocumentEvidence] = []
        reasons: list[str] = []
        check_messages: list[dict[str, str]] = []
        post_messages: list[dict[str, str]] = []
        posted = False
        committed = False
        obj_key = ""
        obj_type = ""
        obj_sys = ""
        posted_belnr = ""
        posted_bukrs = ""
        posted_gjahr = ""
        status = "blocked"
        normalized_system_key = normalize_system_key(payload.system_key)

        try:
            if normalized_system_key not in ALLOWED_SYSTEM_KEYS:
                reasons.append(
                    f"Chave de sistema '{payload.system_key}' fora do escopo permitido para UAT F110."
                )
                return DocumentResult(
                    status="blocked",
                    checked=False,
                    posted=False,
                    committed=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    gl_account=payload.gl_account,
                    amount=_decimal_to_bapi_text(payload.amount),
                    currency=payload.currency,
                    document_date=payload.document_date,
                    posting_date=payload.posting_date,
                    baseline_date=payload.baseline_date,
                    payment_terms=payload.payment_terms,
                    payment_method=payload.payment_method,
                    doc_type=payload.doc_type,
                    reference=payload.reference,
                    header_text=payload.header_text,
                    item_text=payload.item_text,
                    fiscal_year=payload.posting_date[:4],
                    reasons=reasons,
                    evidences=evidences,
                )

            self.connect(payload.system_key)

            if self.system_id not in ALLOWED_SYSTEM_IDS:
                reasons.append(
                    f"Execucao bloqueada: sistema ligado '{self.system_id or '?'}' fora do escopo permitido para UAT F110."
                )
                return DocumentResult(
                    status="blocked",
                    checked=False,
                    posted=False,
                    committed=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=payload.company_code,
                    vendor=payload.vendor,
                    gl_account=payload.gl_account,
                    amount=_decimal_to_bapi_text(payload.amount),
                    currency=payload.currency,
                    document_date=payload.document_date,
                    posting_date=payload.posting_date,
                    baseline_date=payload.baseline_date,
                    payment_terms=payload.payment_terms,
                    payment_method=payload.payment_method,
                    doc_type=payload.doc_type,
                    reference=payload.reference,
                    header_text=payload.header_text,
                    item_text=payload.item_text,
                    fiscal_year=payload.posting_date[:4],
                    reasons=reasons,
                    evidences=evidences,
                )

            company_code = payload.company_code
            vendor_no = zero_pad_if_numeric(payload.vendor)
            gl_account = zero_pad_if_numeric(payload.gl_account)
            company_currency = self._read_company_currency(company_code)
            vendor_defaults = self._read_vendor_defaults(company_code, vendor_no)
            resolved_payment_terms = str(payload.payment_terms or vendor_defaults.get("ZTERM") or "").strip().upper()
            resolved_currency = str(payload.currency or company_currency or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY
            resolved_reference = self._next_reference(company_code, payload.reference)
            resolved_payload = DocumentInput(
                system_key=payload.system_key,
                company_code=company_code,
                vendor=vendor_no,
                gl_account=gl_account,
                amount=payload.amount,
                currency=resolved_currency,
                document_date=payload.document_date,
                posting_date=payload.posting_date,
                baseline_date=payload.baseline_date,
                payment_terms=resolved_payment_terms,
                payment_method=payload.payment_method,
                doc_type=payload.doc_type,
                reference=resolved_reference,
                header_text=payload.header_text,
                item_text=payload.item_text,
                check_only=payload.check_only,
            )

            self._add_evidence(
                evidences,
                "MASTER_DATA",
                "ok",
                "Dados base resolvidos para a criacao do documento.",
                [
                    {
                        "COMP_CODE": company_code,
                        "VENDOR": vendor_no,
                        "GL_ACCOUNT": gl_account,
                        "CURRENCY": resolved_currency,
                        "PAYMENT_TERMS": resolved_payment_terms,
                        "PAYMENT_METHOD": resolved_payload.payment_method,
                    }
                ],
            )

            # --------------------------------------------------------------
            # 1) CHECK
            # --------------------------------------------------------------
            check_result = self._call_bapi("BAPI_ACC_DOCUMENT_CHECK", resolved_payload, self.user, resolved_payment_terms, resolved_currency)
            check_messages = _bapi_messages(check_result)
            if check_messages:
                self._add_evidence(
                    evidences,
                    "BAPI_ACC_DOCUMENT_CHECK",
                    "ok" if not _has_fatal_message(check_messages) else "erro",
                    "Validacao BAPI executada.",
                    check_messages,
                )
            else:
                self._add_evidence(evidences, "BAPI_ACC_DOCUMENT_CHECK", "aviso", "A BAPI nao devolveu mensagens.")

            if _has_fatal_message(check_messages):
                reasons.extend(self._format_messages(check_messages))
                status = "blocked"
                return DocumentResult(
                    status=status,
                    checked=True,
                    posted=False,
                    committed=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=resolved_payload.company_code,
                    vendor=resolved_payload.vendor,
                    gl_account=resolved_payload.gl_account,
                    amount=_decimal_to_bapi_text(resolved_payload.amount),
                    currency=resolved_payload.currency,
                    document_date=resolved_payload.document_date,
                    posting_date=resolved_payload.posting_date,
                    baseline_date=resolved_payload.baseline_date,
                    payment_terms=resolved_payload.payment_terms,
                    payment_method=resolved_payload.payment_method,
                    doc_type=resolved_payload.doc_type,
                    reference=resolved_payload.reference,
                    header_text=resolved_payload.header_text,
                    item_text=resolved_payload.item_text,
                    fiscal_year=resolved_payload.posting_date[:4],
                    reasons=reasons,
                    evidences=evidences,
                    check_messages=check_messages,
                    post_messages=post_messages,
                )

            if resolved_payload.check_only:
                status = "checked"
                return DocumentResult(
                    status=status,
                    checked=True,
                    posted=False,
                    committed=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=resolved_payload.company_code,
                    vendor=resolved_payload.vendor,
                    gl_account=resolved_payload.gl_account,
                    amount=_decimal_to_bapi_text(resolved_payload.amount),
                    currency=resolved_payload.currency,
                    document_date=resolved_payload.document_date,
                    posting_date=resolved_payload.posting_date,
                    baseline_date=resolved_payload.baseline_date,
                    payment_terms=resolved_payload.payment_terms,
                    payment_method=resolved_payload.payment_method,
                    doc_type=resolved_payload.doc_type,
                    reference=resolved_payload.reference,
                    header_text=resolved_payload.header_text,
                    item_text=resolved_payload.item_text,
                    fiscal_year=resolved_payload.posting_date[:4],
                    reasons=reasons,
                    evidences=evidences,
                    check_messages=check_messages,
                    post_messages=post_messages,
                )

            # --------------------------------------------------------------
            # 2) POST
            # --------------------------------------------------------------
            post_result = self._call_bapi("BAPI_ACC_DOCUMENT_POST", resolved_payload, self.user, resolved_payment_terms, resolved_currency)
            post_messages = _bapi_messages(post_result)
            obj_key = str(post_result.get("OBJ_KEY") or "").strip()
            obj_type = str(post_result.get("OBJ_TYPE") or "").strip()
            obj_sys = str(post_result.get("OBJ_SYS") or "").strip()

            if post_messages:
                self._add_evidence(
                    evidences,
                    "BAPI_ACC_DOCUMENT_POST",
                    "ok" if not _has_fatal_message(post_messages) else "erro",
                    "Lancamento executado.",
                    post_messages,
                )
            else:
                self._add_evidence(evidences, "BAPI_ACC_DOCUMENT_POST", "aviso", "A BAPI nao devolveu mensagens.")

            if _has_fatal_message(post_messages):
                reasons.extend(self._format_messages(post_messages))
                try:
                    self.conn.call("BAPI_TRANSACTION_ROLLBACK")
                except Exception:
                    pass
                status = "blocked"
                return DocumentResult(
                    status=status,
                    checked=True,
                    posted=False,
                    committed=False,
                    system_id=self.system_id,
                    client=self.client,
                    company_code=resolved_payload.company_code,
                    vendor=resolved_payload.vendor,
                    gl_account=resolved_payload.gl_account,
                    amount=_decimal_to_bapi_text(resolved_payload.amount),
                    currency=resolved_payload.currency,
                    document_date=resolved_payload.document_date,
                    posting_date=resolved_payload.posting_date,
                    baseline_date=resolved_payload.baseline_date,
                    payment_terms=resolved_payload.payment_terms,
                    payment_method=resolved_payload.payment_method,
                    doc_type=resolved_payload.doc_type,
                    reference=resolved_payload.reference,
                    header_text=resolved_payload.header_text,
                    item_text=resolved_payload.item_text,
                    fiscal_year=resolved_payload.posting_date[:4],
                    obj_key=obj_key,
                    obj_type=obj_type,
                    obj_sys=obj_sys,
                    reasons=reasons,
                    evidences=evidences,
                    check_messages=check_messages,
                    post_messages=post_messages,
                )

            # --------------------------------------------------------------
            # 3) COMMIT
            # --------------------------------------------------------------
            self.conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
            committed = True
            posted = True
            posted_belnr, posted_bukrs, posted_gjahr = self._parse_obj_key(obj_key)
            if posted_belnr and posted_bukrs and posted_gjahr:
                try:
                    bkpf_rows, _ = read_table_with_fallbacks(
                        self.conn,
                        "BKPF",
                        [["BUKRS", "BELNR", "GJAHR", "BLART", "BUDAT", "USNAM", "XBLNR"], ["BUKRS", "BELNR", "GJAHR"]],
                        options=[
                            f"BUKRS = '{posted_bukrs}'",
                            f"AND BELNR = '{posted_belnr}'",
                            f"AND GJAHR = '{posted_gjahr}'",
                        ],
                        rowcount=1,
                    )
                    if bkpf_rows:
                        self._add_evidence(evidences, "BKPF", "ok", "Documento confirmado apos o commit.", bkpf_rows)
                    else:
                        self._add_evidence(evidences, "BKPF", "aviso", "Commit executado, mas BKPF ainda nao devolveu o documento.")
                except Exception as exc:
                    self._add_evidence(evidences, "BKPF", "aviso", f"Commit executado, mas falhou a validacao em BKPF: {exc}")

            status = "posted"
            return DocumentResult(
                status=status,
                checked=True,
                posted=posted,
                committed=committed,
                system_id=self.system_id,
                client=self.client,
                company_code=resolved_payload.company_code,
                vendor=resolved_payload.vendor,
                gl_account=resolved_payload.gl_account,
                amount=_decimal_to_bapi_text(resolved_payload.amount),
                currency=resolved_payload.currency,
                document_date=resolved_payload.document_date,
                posting_date=resolved_payload.posting_date,
                baseline_date=resolved_payload.baseline_date,
                payment_terms=resolved_payload.payment_terms,
                payment_method=resolved_payload.payment_method,
                doc_type=resolved_payload.doc_type,
                reference=resolved_payload.reference,
                header_text=resolved_payload.header_text,
                item_text=resolved_payload.item_text,
                fiscal_year=resolved_payload.posting_date[:4],
                posted_belnr=posted_belnr,
                posted_bukrs=posted_bukrs,
                posted_gjahr=posted_gjahr,
                obj_key=obj_key,
                obj_type=obj_type,
                obj_sys=obj_sys,
                reasons=reasons,
                evidences=evidences,
                check_messages=check_messages,
                post_messages=post_messages,
            )
        except Exception as exc:
            reasons.append(str(exc))
            try:
                self.conn.call("BAPI_TRANSACTION_ROLLBACK")
            except Exception:
                pass
            return DocumentResult(
                status="blocked",
                checked=bool(check_messages),
                posted=posted,
                committed=committed,
                system_id=self.system_id,
                client=self.client,
                company_code=payload.company_code,
                vendor=payload.vendor,
                gl_account=payload.gl_account,
                amount=_decimal_to_bapi_text(payload.amount),
                currency=payload.currency,
                document_date=payload.document_date,
                posting_date=payload.posting_date,
                baseline_date=payload.baseline_date,
                payment_terms=payload.payment_terms,
                payment_method=payload.payment_method,
                doc_type=payload.doc_type,
                reference=payload.reference,
                header_text=payload.header_text,
                item_text=payload.item_text,
                fiscal_year=payload.posting_date[:4],
                obj_key=obj_key,
                obj_type=obj_type,
                obj_sys=obj_sys,
                reasons=reasons,
                evidences=evidences,
                check_messages=check_messages,
                post_messages=post_messages,
            )
        finally:
            self.close()


# =============================================================================
# (3) RESULTADO / CLI
# =============================================================================

def _default_payload_from_args(args: argparse.Namespace) -> DocumentInput:
    return DocumentInput(
        system_key=str(args.sap_system).strip(),
        company_code=str(args.company_code).strip(),
        vendor=str(args.vendor).strip(),
        gl_account=str(args.gl_account).strip(),
        amount=_normalize_amount(args.amount),
        currency=str(args.currency or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY,
        document_date=str(args.document_date or _today_yyyymmdd()).strip(),
        posting_date=str(args.posting_date or _today_yyyymmdd()).strip(),
        baseline_date=str(args.baseline_date or "").strip(),
        payment_terms=str(args.payment_terms or "").strip().upper(),
        payment_method=str(args.payment_method or "S").strip().upper(),
        doc_type=str(args.doc_type or DEFAULT_DOC_TYPE).strip().upper(),
        reference=str(args.reference or DEFAULT_REFERENCE).strip(),
        header_text=str(args.header_text or DEFAULT_HEADER_TEXT).strip(),
        item_text=str(args.item_text or DEFAULT_ITEM_TEXT).strip(),
        check_only=bool(args.check_only),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Criar documento FI de fornecedor via RFC/BAPI para teste da F110.")
    parser.add_argument("--sap-system", default="QAD", help="Chave de sistema SAP. Padrao: QAD")
    parser.add_argument("--company-code", default="2010", help="Codigo da empresa (BUKRS)")
    parser.add_argument("--vendor", default="10000040", help="Numero do fornecedor")
    parser.add_argument("--gl-account", default="12010741", help="Conta contabil de contrapartida")
    parser.add_argument("--amount", default="88,88", help="Valor do documento")
    parser.add_argument("--currency", default=DEFAULT_CURRENCY, help="Moeda do documento")
    parser.add_argument("--document-date", default=_today_yyyymmdd(), help="Data do documento em YYYYMMDD")
    parser.add_argument("--posting-date", default=_today_yyyymmdd(), help="Data de lancamento em YYYYMMDD")
    parser.add_argument("--baseline-date", default="", help="Data base de pagamento em YYYYMMDD")
    parser.add_argument("--payment-terms", default="", help="Condicao de pagamento do fornecedor")
    parser.add_argument("--payment-method", default="S", help="Metodo de pagamento")
    parser.add_argument("--doc-type", default=DEFAULT_DOC_TYPE, help="Tipo de documento SAP")
    parser.add_argument("--reference", default=DEFAULT_REFERENCE, help="Referencia externa do documento")
    parser.add_argument("--header-text", default=DEFAULT_HEADER_TEXT, help="Texto do cabecalho")
    parser.add_argument("--item-text", default=DEFAULT_ITEM_TEXT, help="Texto das linhas")
    parser.add_argument("--check-only", action="store_true", help="Executa apenas a BAPI de validacao")
    return parser


def format_result(result: DocumentResult) -> str:
    lines: list[str] = []
    lines.append("=" * 84)
    lines.append("UAT CRIACAO DOCUMENTO FI F110")
    lines.append("=" * 84)
    lines.append(f"Sistema        : {result.system_id or '?'}")
    lines.append(f"Cliente        : {result.client or '?'}")
    lines.append(f"Empresa        : {result.company_code}")
    lines.append(f"Fornecedor     : {result.vendor}")
    lines.append(f"Conta GL       : {result.gl_account}")
    lines.append(f"Valor          : {result.amount} {result.currency}")
    lines.append(f"Data doc.      : {result.document_date}")
    lines.append(f"Data lanc.     : {result.posting_date}")
    lines.append(f"Data base      : {result.baseline_date or '?'}")
    lines.append(f"Pagamento      : {result.payment_terms or '?'}")
    lines.append(f"Metodo esp.    : {result.payment_method}")
    lines.append(f"Tipo doc.      : {result.doc_type}")
    lines.append(f"Referencia     : {result.reference}")
    lines.append(f"Estado         : {result.status.upper()}")
    lines.append(f"Check OK       : {'SIM' if result.checked else 'NAO'}")
    lines.append(f"Post OK        : {'SIM' if result.posted else 'NAO'}")
    lines.append(f"Commit OK      : {'SIM' if result.committed else 'NAO'}")
    if result.posted_belnr:
        lines.append(f"BELNR          : {result.posted_belnr}")
        lines.append(f"BUKRS          : {result.posted_bukrs}")
        lines.append(f"GJAHR          : {result.posted_gjahr}")
    if result.obj_key:
        lines.append(f"OBJ_KEY        : {result.obj_key}")
    lines.append("")
    lines.append("Motivos")
    lines.append("-" * 84)
    if result.reasons:
        for index, reason in enumerate(result.reasons, start=1):
            lines.append(f"{index}. {reason}")
    else:
        lines.append("Sem observacoes adicionais.")
    lines.append("")
    lines.append("Mensagens CHECK")
    lines.append("-" * 84)
    if result.check_messages:
        lines.extend(_format_messages_for_output(result.check_messages))
    else:
        lines.append("Sem mensagens de check.")
    lines.append("")
    lines.append("Mensagens POST")
    lines.append("-" * 84)
    if result.post_messages:
        lines.extend(_format_messages_for_output(result.post_messages))
    else:
        lines.append("Sem mensagens de post.")
    lines.append("")
    lines.append("Evidencias")
    lines.append("-" * 84)
    for evidence in result.evidences:
        lines.append(f"[{evidence.status.upper()}] {evidence.name}: {evidence.details}")
        for row in evidence.rows[:5]:
            lines.append(f"  - {row}")
    return "\n".join(lines)


def _format_messages_for_output(messages: list[dict[str, str]]) -> list[str]:
    output: list[str] = []
    for index, line in enumerate(messages, start=1):
        msg_type = line.get("TYPE", "?")
        msg_id = line.get("ID", "")
        msg_no = line.get("NUMBER", "")
        text = line.get("MESSAGE", "")
        output.append(f"{index}. [{msg_type}] {msg_id} {msg_no} {text}".rstrip())
    return output


def executar(ambiente_cockpit: str | None = None, **kwargs: Any) -> DocumentResult:
    load_project_dotenv()
    payload = DocumentInput(
        system_key=str(kwargs.get("system_key") or ambiente_cockpit or kwargs.get("sap_system") or "QAD").strip(),
        company_code=str(kwargs.get("company_code") or "2010").strip(),
        vendor=str(kwargs.get("vendor") or "10000040").strip(),
        gl_account=str(kwargs.get("gl_account") or "12010741").strip(),
        amount=_normalize_amount(kwargs.get("amount") or "88,88"),
        currency=str(kwargs.get("currency") or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY,
        document_date=str(kwargs.get("document_date") or _today_yyyymmdd()).strip(),
        posting_date=str(kwargs.get("posting_date") or _today_yyyymmdd()).strip(),
        baseline_date=str(kwargs.get("baseline_date") or "").strip(),
        payment_terms=str(kwargs.get("payment_terms") or "").strip().upper(),
        payment_method=str(kwargs.get("payment_method") or "S").strip().upper(),
        doc_type=str(kwargs.get("doc_type") or DEFAULT_DOC_TYPE).strip().upper(),
        reference=str(kwargs.get("reference") or DEFAULT_REFERENCE).strip(),
        header_text=str(kwargs.get("header_text") or DEFAULT_HEADER_TEXT).strip(),
        item_text=str(kwargs.get("item_text") or DEFAULT_ITEM_TEXT).strip(),
        check_only=bool(kwargs.get("check_only", False)),
    )
    poster = FiDocumentPoster()
    result = poster.run(payload)
    print(format_result(result))
    return result


def main(argv: list[str] | None = None) -> int:
    load_project_dotenv()
    args = build_parser().parse_args(argv)
    payload = _default_payload_from_args(args)
    poster = FiDocumentPoster()
    result = poster.run(payload)
    print(format_result(result))
    return 0 if result.posted or result.checked else 1


if __name__ == "__main__":
    raise SystemExit(main())
