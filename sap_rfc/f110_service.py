from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    make_option_eq,
    make_write_guard,
    read_table,
)

# Este serviço só vai até a PROPOSTA do F110 (equivalente a "Vorlauf"/PAR_XVL='X').
# Nunca chama JOB_SUBMIT com PAR_XVL vazio (isso seria o pagamento/cobrança real).
OPERATION_PAGAMENTO = "pagamento"
OPERATION_COBRANCA = "cobranca"

WRITE_ALLOWED_FUNCTIONS = (
    "RFC_PING",
    "RFC_READ_TABLE",
    "JOB_OPEN",
    "JOB_SUBMIT",
    "JOB_CLOSE",
)
WRITE_ALLOWED_TABLES = ("TBTCO", "REGUV", "REGUH", "REGUP", "BSIK", "BSID")

_JOB_STATUS_LABELS = {
    "S": "Agendado",
    "R": "Em execução",
    "F": "Concluído",
    "A": "Cancelado (abortado)",
    "Y": "Pronto para iniciar",
    "Z": "Erro ao iniciar",
}


@dataclass
class F110ProposalResult:
    ok: bool
    status: str
    message: str
    operation_type: str
    environment: str
    company_code: str
    account_number: str
    run_date: str
    run_id: str
    job_name: str = ""
    job_count: str = ""
    job_status: str = ""
    document_number: str = ""
    document_found: bool | None = None
    document_included_in_proposal: bool | None = None
    proposal_items: list[dict[str, Any]] = field(default_factory=list)
    payload: dict[str, Any] = field(default_factory=dict)


def _require_pyrfc() -> None:
    if Connection is None:
        raise RuntimeError(f"PyRFC indisponível: {_PYRFC_IMPORT_ERROR}")


def _normalize_operation_type(operation_type: str) -> str:
    value = str(operation_type or "").strip().lower()
    aliases = {
        "pagamento": OPERATION_PAGAMENTO,
        "pagar": OPERATION_PAGAMENTO,
        "fornecedor": OPERATION_PAGAMENTO,
        "cobranca": OPERATION_COBRANCA,
        "cobrança": OPERATION_COBRANCA,
        "cliente": OPERATION_COBRANCA,
    }
    normalized = aliases.get(value)
    if not normalized:
        raise ValueError(f"Tipo de operação inválido: {operation_type!r} (use 'pagamento' ou 'cobranca').")
    return normalized


def _to_sap_date(value: Any) -> str:
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    raw = str(value or "").strip()
    if not raw:
        raise ValueError("Data obrigatória em falta.")
    return date.fromisoformat(raw).strftime("%Y%m%d")


def _generate_run_id(operation_type: str) -> str:
    """Gera um PAR_LFID (5 caracteres, obrigatório e único por LAUFD) para o cockpit."""
    prefix = "P" if operation_type == OPERATION_PAGAMENTO else "C"
    suffix = datetime.now().strftime("%H%M%S")[-4:]
    return f"{prefix}{suffix}"


def _rsparam(selname: str, kind: str, low: str, high: str = "", sign: str = "I", option: str = "EQ") -> dict[str, str]:
    return {
        "SELNAME": selname,
        "KIND": kind,
        "SIGN": sign,
        "OPTION": option,
        "LOW": low,
        "HIGH": high,
    }


def _check_document_exists(
    connection: Any,
    guard: Any,
    *,
    operation_type: str,
    company_code: str,
    account_number: str,
    document_number: str,
) -> bool:
    table_name = "BSIK" if operation_type == OPERATION_PAGAMENTO else "BSID"
    account_field = "LIFNR" if operation_type == OPERATION_PAGAMENTO else "KUNNR"
    options = (
        make_option_eq("BUKRS", company_code)
        + [{"TEXT": f"AND {account_field} = '{account_number}'"}]
        + [{"TEXT": f"AND BELNR = '{document_number}'"}]
    )
    rows = read_table(
        connection,
        guard,
        table_name=table_name,
        fields=["BELNR"],
        options=options,
        rowcount=1,
    )
    return bool(rows)


def _read_proposal_items(
    connection: Any,
    guard: Any,
    *,
    run_date: str,
    run_id: str,
) -> list[dict[str, Any]]:
    rows = read_table(
        connection,
        guard,
        table_name="REGUP",
        fields=["LAUFD", "LAUFI", "LIFNR", "KUNNR", "BELNR", "GJAHR", "WRBTR", "WAERS"],
        options=make_option_eq("LAUFD", run_date) + [{"TEXT": f"AND LAUFI = '{run_id}'"}],
        rowcount=0,
    )
    items = []
    for lifnr, kunnr, belnr, gjahr, wrbtr, waers in [(r[2], r[3], r[4], r[5], r[6], r[7]) for r in rows]:
        items.append(
            {
                "vendor": lifnr.strip(),
                "customer": kunnr.strip(),
                "document_number": belnr.strip(),
                "fiscal_year": gjahr.strip(),
                "amount": wrbtr.strip(),
                "currency": waers.strip(),
            }
        )
    return items


def _poll_job_status(
    connection: Any,
    guard: Any,
    *,
    job_name: str,
    job_count: str,
    timeout_seconds: int = 45,
    interval_seconds: float = 1.5,
) -> str:
    deadline = time.monotonic() + timeout_seconds
    last_status = ""
    while time.monotonic() < deadline:
        rows = read_table(
            connection,
            guard,
            table_name="TBTCO",
            fields=["STATUS"],
            options=make_option_eq("JOBNAME", job_name) + [{"TEXT": f"AND JOBCOUNT = '{job_count}'"}],
            rowcount=1,
        )
        if rows:
            last_status = rows[0][0].strip()
            if last_status in {"F", "A"}:
                return last_status
        time.sleep(interval_seconds)
    return last_status or "?"


def run_f110_proposal(
    environment: str,
    operation_type: str,
    *,
    company_code: str,
    payment_method: str,
    account_number: str,
    posting_date: str,
    next_due_date: str,
    document_number: str = "",
    run_date: str = "",
) -> F110ProposalResult:
    """Executa só a etapa de PROPOSTA do F110 (PAR_XVL='X'), via JOB_SUBMIT do RFF110S.

    Nunca executa o pagamento/cobrança real. `document_number`, quando informado, é
    apenas validado antes (existe para a conta/empresa?) e conferido depois na
    proposta gerada — a seleção real de itens em aberto continua sendo feita pelo
    próprio F110 conforme a janela de datas informada.
    """
    _require_pyrfc()

    op_type = _normalize_operation_type(operation_type)
    company_code = str(company_code or "").strip().upper()
    payment_method = str(payment_method or "").strip().upper()
    account_number = str(account_number or "").strip().upper()
    document_number = str(document_number or "").strip().upper()

    if not company_code:
        raise ValueError("company_code é obrigatório.")
    if not payment_method:
        raise ValueError("payment_method é obrigatório.")
    if not account_number:
        raise ValueError("account_number (fornecedor/cliente) é obrigatório.")

    posting_date_sap = _to_sap_date(posting_date)
    next_due_date_sap = _to_sap_date(next_due_date)
    run_date_sap = _to_sap_date(run_date) if str(run_date or "").strip() else posting_date_sap
    run_id = _generate_run_id(op_type)

    connection_params = build_connection_params_for_env(environment)
    guard = make_write_guard(WRITE_ALLOWED_FUNCTIONS, WRITE_ALLOWED_TABLES)

    payload = {
        "operation_type": op_type,
        "company_code": company_code,
        "payment_method": payment_method,
        "account_number": account_number,
        "posting_date": posting_date_sap,
        "next_due_date": next_due_date_sap,
        "run_date": run_date_sap,
        "run_id": run_id,
        "document_number": document_number,
    }

    connection = Connection(**connection_params)  # type: ignore[misc]
    try:
        document_found: bool | None = None
        if document_number:
            document_found = _check_document_exists(
                connection,
                guard,
                operation_type=op_type,
                company_code=company_code,
                account_number=account_number,
                document_number=document_number,
            )
            if not document_found:
                return F110ProposalResult(
                    ok=False,
                    status="ERRO",
                    message=(
                        f"Documento {document_number} não encontrado em aberto para "
                        f"{account_number} na empresa {company_code}."
                    ),
                    operation_type=op_type,
                    environment=str(environment or "").strip().upper(),
                    company_code=company_code,
                    account_number=account_number,
                    run_date=run_date_sap,
                    run_id=run_id,
                    document_number=document_number,
                    document_found=False,
                    payload=payload,
                )

        params = [
            _rsparam("PAR_LFID", "P", run_id),
            _rsparam("PAR_XVL", "P", "X"),
            _rsparam("PAR_BUDA", "P", posting_date_sap),
            _rsparam("PAR_GRDA", "P", posting_date_sap),
            _rsparam("PAR_NEDA", "P", next_due_date_sap),
            _rsparam("PAR_ZWE", "P", payment_method),
            _rsparam("SEL_BUKR", "S", company_code),
        ]
        if op_type == OPERATION_PAGAMENTO:
            params.append(_rsparam("SEL_KRED", "S", account_number))
        else:
            params.append(_rsparam("SEL_DEBI", "S", account_number))

        job_name = f"COCKPIT_F110_{run_id}"

        guard.assert_function_allowed("JOB_OPEN")
        open_result = connection.call("JOB_OPEN", JOBNAME=job_name, JOBGROUP="F110")
        job_count = str(open_result.get("JOBCOUNT") or "").strip()
        if not job_count:
            return F110ProposalResult(
                ok=False,
                status="ERRO",
                message="JOB_OPEN não devolveu JOBCOUNT.",
                operation_type=op_type,
                environment=str(environment or "").strip().upper(),
                company_code=company_code,
                account_number=account_number,
                run_date=run_date_sap,
                run_id=run_id,
                document_number=document_number,
                document_found=document_found,
                job_name=job_name,
                payload=payload,
            )

        guard.assert_function_allowed("JOB_SUBMIT")
        connection.call(
            "JOB_SUBMIT",
            AUTHCKNAM=connection_params["user"],
            JOBCOUNT=job_count,
            JOBNAME=job_name,
            REPORT="RFF110S",
            SELECTION_TABLE=params,
        )

        guard.assert_function_allowed("JOB_CLOSE")
        connection.call(
            "JOB_CLOSE",
            JOBNAME=job_name,
            JOBCOUNT=job_count,
            STRTIMMED="X",
        )

        job_status = _poll_job_status(connection, guard, job_name=job_name, job_count=job_count)
        job_status_label = _JOB_STATUS_LABELS.get(job_status, job_status)

        proposal_items = _read_proposal_items(connection, guard, run_date=run_date_sap, run_id=run_id)
        document_included = None
        if document_number:
            document_included = any(item["document_number"] == document_number for item in proposal_items)

        ok = job_status == "F"
        message = (
            f"Proposta F110 {run_date_sap}/{run_id} — job {job_name} ({job_count}): {job_status_label}. "
            f"{len(proposal_items)} item(ns) na proposta."
        )
        if document_number:
            message += (
                f" Documento {document_number}: "
                + ("incluído na proposta." if document_included else "NÃO apareceu na proposta — confira o F110.")
            )

        return F110ProposalResult(
            ok=ok,
            status="SUCESSO" if ok else "ATENÇÃO",
            message=message,
            operation_type=op_type,
            environment=str(environment or "").strip().upper(),
            company_code=company_code,
            account_number=account_number,
            run_date=run_date_sap,
            run_id=run_id,
            job_name=job_name,
            job_count=job_count,
            job_status=job_status,
            document_number=document_number,
            document_found=document_found,
            document_included_in_proposal=document_included,
            proposal_items=proposal_items,
            payload=payload,
        )
    finally:
        try:
            connection.close()
        except Exception:
            pass
