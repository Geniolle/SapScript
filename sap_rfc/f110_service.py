from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
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
    "BAPI_XMI_LOGON",
    "BAPI_XMI_LOGOFF",
    "BAPI_XBP_JOB_OPEN",
    "BAPI_XBP_JOB_ADD_ABAP_STEP",
    "BAPI_XBP_JOB_CLOSE",
    "BAPI_XBP_JOB_START_IMMEDIATELY",
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


def _is_windows_runtime() -> bool:
    return os.name == "nt" or sys.platform.startswith("win")


def _bridge_python_executable() -> str | None:
    candidates = [
        os.getenv("SAP_FI_BRIDGE_PYTHON", "").strip(),
        os.getenv("WORKFLOW_PYTHON_EXEC", "").strip(),
    ]
    if _is_windows_runtime():
        candidates.append(str((Path(__file__).resolve().parents[1] / ".venv-rfc" / "Scripts" / "python.exe").resolve()))

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def _run_f110_proposal_via_bridge(
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
    python_exe = _bridge_python_executable()
    runtime = "Windows" if _is_windows_runtime() else "WSL/Linux"
    if not python_exe:
        raise RuntimeError(
            f"Execução F110 via bridge indisponível neste runtime ({runtime}). "
            "Configure SAP_FI_BRIDGE_PYTHON ou WORKFLOW_PYTHON_EXEC com um Python compatível com PyRFC."
        )

    payload = {
        "environment": environment,
        "operation_type": operation_type,
        "company_code": company_code,
        "payment_method": payment_method,
        "account_number": account_number,
        "posting_date": posting_date,
        "next_due_date": next_due_date,
        "document_number": document_number,
        "run_date": run_date,
    }
    bridge_code = (
        "import json, sys\n"
        "from pathlib import Path\n"
        "sys.path.insert(0, Path.cwd().as_posix())\n"
        "from sap_rfc.f110_service import _run_f110_proposal_core as _run\n"
        "payload = json.loads(sys.argv[1])\n"
        "result = _run(\n"
        "    payload['environment'],\n"
        "    payload['operation_type'],\n"
        "    company_code=payload['company_code'],\n"
        "    payment_method=payload['payment_method'],\n"
        "    account_number=payload['account_number'],\n"
        "    posting_date=payload['posting_date'],\n"
        "    next_due_date=payload['next_due_date'],\n"
        "    document_number=payload.get('document_number', ''),\n"
        "    run_date=payload.get('run_date', ''),\n"
        ")\n"
        "print(json.dumps(result.__dict__, ensure_ascii=False))\n"
    )

    proc = subprocess.run(
        [python_exe, "-c", bridge_code, json.dumps(payload, ensure_ascii=False)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        cwd=str(Path(__file__).resolve().parent.parent),
    )
    if proc.returncode not in {0, 1}:
        raise RuntimeError(
            f"Bridge F110 falhou com exit code {proc.returncode}.\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}"
        )
    if not proc.stdout.strip():
        raise RuntimeError(f"Bridge F110 devolveu saída vazia.\nSTDERR: {proc.stderr}")

    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except Exception as exc:
        raise RuntimeError(f"Bridge F110 devolveu JSON inválido.\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}") from exc

    return F110ProposalResult(**data)


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
        + [{"TEXT": f"AND BELNR = '{document_number.strip()[:10]}'"}]
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


def _run_f110_proposal_core(
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
    xmi_logged_on = False
    guard.assert_function_allowed("BAPI_XMI_LOGON")
    xmi_logon = connection.call(
        "BAPI_XMI_LOGON",
        EXTCOMPANY="SapScript",
        EXTPRODUCT="SapScriptCockpit",
        INTERFACE="XBP",
        VERSION="2.0",
    )
    if not str(xmi_logon.get("SESSIONID") or "").strip():
        return F110ProposalResult(
            ok=False,
            status="ERRO",
            message="BAPI_XMI_LOGON não retornou SESSIONID.",
            operation_type=op_type,
            environment=str(environment or "").strip().upper(),
            company_code=company_code,
            account_number=account_number,
            run_date=run_date_sap,
            run_id=run_id,
            document_number=document_number,
            document_found=None,
            payload=payload,
        )
    xmi_logged_on = True
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

        guard.assert_function_allowed("BAPI_XBP_JOB_OPEN")
        open_result = connection.call(
            "BAPI_XBP_JOB_OPEN",
            JOBNAME=job_name,
            JOBCLASS="C",
            EXTERNAL_USER_NAME=connection_params["user"],
        )
        job_count = str(open_result.get("JOBCOUNT") or "").strip()
        if not job_count:
            return F110ProposalResult(
                ok=False,
                status="ERRO",
                message="BAPI_XBP_JOB_OPEN não devolveu JOBCOUNT.",
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

        guard.assert_function_allowed("BAPI_XBP_JOB_ADD_ABAP_STEP")
        connection.call(
            "BAPI_XBP_JOB_ADD_ABAP_STEP",
            JOBCOUNT=job_count,
            JOBNAME=job_name,
            ABAP_PROGRAM_NAME="RFF110S",
            ABAP_VARIANT_NAME="",
            SELINFO=params,
        )

        guard.assert_function_allowed("BAPI_XBP_JOB_CLOSE")
        connection.call(
            "BAPI_XBP_JOB_CLOSE",
            JOBNAME=job_name,
            JOBCOUNT=job_count,
            EXTERNAL_USER_NAME=connection_params["user"],
        )

        guard.assert_function_allowed("BAPI_XBP_JOB_START_IMMEDIATELY")
        connection.call(
            "BAPI_XBP_JOB_START_IMMEDIATELY",
            JOBNAME=job_name,
            JOBCOUNT=job_count,
            EXTERNAL_USER_NAME=connection_params["user"],
        )

        job_status = _poll_job_status(connection, guard, job_name=job_name, job_count=job_count)
        job_status_label = _JOB_STATUS_LABELS.get(job_status, job_status)

        proposal_items = _read_proposal_items(connection, guard, run_date=run_date_sap, run_id=run_id)
        document_included = None
        if document_number:
            document_included = any(item["document_number"] == document_number.strip()[:10] for item in proposal_items)

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
        if xmi_logged_on:
            try:
                guard.assert_function_allowed("BAPI_XMI_LOGOFF")
                connection.call("BAPI_XMI_LOGOFF", INTERFACE="XBP")
            except Exception:
                pass
        connection.close()


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
    if Connection is None:
        return _run_f110_proposal_via_bridge(
            environment,
            operation_type,
            company_code=company_code,
            payment_method=payment_method,
            account_number=account_number,
            posting_date=posting_date,
            next_due_date=next_due_date,
            document_number=document_number,
            run_date=run_date,
        )
    return _run_f110_proposal_core(
        environment,
        operation_type,
        company_code=company_code,
        payment_method=payment_method,
        account_number=account_number,
        posting_date=posting_date,
        next_due_date=next_due_date,
        document_number=document_number,
        run_date=run_date,
    )
