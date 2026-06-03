from __future__ import annotations

import re

from .models import SapErrorSignal, TicketContext


_TRANSACTION_RE = re.compile(r"\b(?:T[- ]?CODE|TRANSA[CÇ][AÃ]O|TRANSACTION)[:\s-]+([A-Z0-9_/]{2,20})", re.IGNORECASE)
_MESSAGE_RE = re.compile(r"\b(?:MSG|MESSAGE|MENSAGEM)[:\s-]+([A-Z0-9]{1,20})[-/\s]+(\d{1,3})", re.IGNORECASE)
_BUKRS_RE = re.compile(r"\b(?:BUKRS|EMPRESA|COMPANY CODE)[:\s-]+([A-Z0-9]{4})", re.IGNORECASE)
_BELNR_RE = re.compile(r"\b(?:BELNR|DOCUMENTO|DOC(?:UMENT)? NUMBER|PEDIDO)[:\s-]+(\d{4,10})", re.IGNORECASE)
_GJAHR_RE = re.compile(r"\b(?:GJAHR|EXERC[IÍ]CIO|FISCAL YEAR|ANO)[:\s-]+(\d{4})", re.IGNORECASE)
_JOB_RE = re.compile(r"\b(?:JOB|BACKGROUND JOB)[:\s-]+([A-Z0-9_/-]{3,32})", re.IGNORECASE)
_USER_RE_EXACT = re.compile(r"\b(?:USER|USU[AÁ]RIO|UTILIZADOR)[:\s-]+([A-Z0-9_\-.]{3,40})", re.IGNORECASE)
_USER_RE_FUZZY = re.compile(r"\b(?:COLABORADORA?|EMPLEAD[OA]|FUNCION[AÁ]RIA?)[:\s]+(?:[a-zA-Z]+\s+){0,4}(\d{5,10})\b", re.IGNORECASE)
_IBAN_RE = re.compile(r"\b([A-Z]{2}\d{2}[A-Z0-9]{10,30})\b", re.IGNORECASE)
_KEYNO_RE = re.compile(r"\b(?:KEYNO|N[ºO]\s*CHAVE|CHAVE)[:\s-]+([A-Z0-9]{3,20})", re.IGNORECASE)
_PROGRAM_RE = re.compile(r"\b(?:PROGRAMA|PROGRAM|REPORT|CLASSE|CLASS|INCLUDE)[:\s-]+([A-Z0-9_=/]{3,80})", re.IGNORECASE)


def _first(regex: re.Pattern[str], text: str, group: int = 1) -> str | None:
    match = regex.search(text)
    return match.group(group).upper() if match else None


def _extract_documents(text: str) -> str | None:
    match = _BELNR_RE.search(text)
    if match:
        return match.group(1).upper()
        
    # Fallback to any 10 digit number
    ten_digit_numbers = list(set(re.findall(r"\b([1-9]\d{9})\b", text)))
    if ten_digit_numbers:
        return ", ".join(sorted(ten_digit_numbers))
    
    return None


def extract_signal(ticket: TicketContext) -> SapErrorSignal:
    text = ticket.full_text
    message_match = _MESSAGE_RE.search(text)
    module = _infer_module(ticket)
    return SapErrorSignal(
        transaction=_first(_TRANSACTION_RE, text),
        program=_first(_PROGRAM_RE, text),
        message_id=message_match.group(1).upper() if message_match else None,
        message_number=message_match.group(2).zfill(3) if message_match else None,
        company_code=_first(_BUKRS_RE, text),
        document_number=_extract_documents(text),
        fiscal_year=_first(_GJAHR_RE, text),
        job_name=_first(_JOB_RE, text),
        user=_first(_USER_RE_EXACT, text) or _first(_USER_RE_FUZZY, text),
        iban=_first(_IBAN_RE, text),
        keyno=_first(_KEYNO_RE, text),
        module=module,
    )


def _infer_module(ticket: TicketContext) -> str | None:
    text = " ".join([ticket.summary, ticket.description, " ".join(ticket.labels), " ".join(ticket.components)]).upper()
    for module in ("FI", "MM", "SD", "WM", "EWM", "HCM", "CO", "BW", "BASIS", "ABAP"):
        if re.search(rf"\b{module}\b", text):
            return module
    if any(token in text for token in ("FBL", "BKPF", "BSEG", "PAYRQ", "F110", "F111", "FIBL")):
        return "FI"
    if any(token in text for token in ("ME21N", "ME22N", "MIGO", "MIRO", "EKPO", "EKKO")):
        return "MM"
    if any(token in text for token in ("VA01", "VA02", "VL01N", "VF01", "VBAK", "VBAP")):
        return "SD"
    return None
