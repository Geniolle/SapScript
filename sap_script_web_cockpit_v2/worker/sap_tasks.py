from __future__ import annotations

import importlib
import os
import sys
import traceback
from typing import Any

import pythoncom
import win32com.client


class SapExecutionError(Exception):
    pass


def _prepare_project_imports() -> None:
    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if project_dir and project_dir not in sys.path:
        sys.path.insert(0, project_dir)


def get_first_available_session() -> Any:
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except Exception as exc:
        raise SapExecutionError(
            "Nao foi possivel ligar ao SAP GUI. Confirma se o SAP Logon esta aberto "
            "e se o SAP GUI Scripting esta ativo."
        ) from exc

    for connection_index in range(application.Children.Count):
        connection = application.Children(connection_index)
        for session_index in range(connection.Children.Count):
            session = connection.Children(session_index)
            try:
                if not session.Busy:
                    return session
            except Exception:
                continue

    raise SapExecutionError("Nao existe nenhuma sessao SAP disponivel.")


def read_sbar_status(session: Any) -> str:
    try:
        return str(session.findById("wnd[0]/sbar").Text).strip()
    except Exception as exc:
        return f"Nao foi possivel ler STATUS em wnd[0]/sbar: {exc}"


def _open_transaction(params: dict[str, Any]) -> tuple[str, str]:
    transaction = str(params.get("transacao") or "SE10").strip().upper().lstrip("/")
    session = get_first_available_session()
    session.findById("wnd[0]/tbar[0]/okcd").Text = f"/n{transaction}"
    session.findById("wnd[0]").sendVKey(0)
    status = read_sbar_status(session)
    log = f"Transacao solicitada: {transaction}\nSTATUS: {status}"
    return status or f"Transacao {transaction} aberta; STATUS vazio em wnd[0]/sbar", log


def _run_sap_cockpit(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    module_name = os.getenv("SAP_COCKPIT_MODULE", "sap_cockpit_web_ready").strip()

    try:
        cockpit = importlib.import_module(module_name)
    except Exception as exc:
        raise SapExecutionError(
            f"Nao foi possivel importar o modulo '{module_name}'. "
            "Confirma SAP_SCRIPT_PROJECT_DIR e SAP_COCKPIT_MODULE."
        ) from exc

    if not hasattr(cockpit, "run_sap_cockpit"):
        raise SapExecutionError(
            f"O modulo '{module_name}' nao tem a funcao run_sap_cockpit(payload)."
        )

    result = cockpit.run_sap_cockpit(params)

    if isinstance(result, tuple) and len(result) == 2:
        return str(result[0] or ""), str(result[1] or "")

    if isinstance(result, dict):
        status = str(result.get("status") or result.get("STATUS") or "").strip()
        log = str(result.get("log") or result.get("log_text") or "")
        return status, log

    return str(result or ""), ""


def run_sap_task(job: dict[str, Any]) -> tuple[str, str]:
    task = job["task"]
    params = job.get("params", {}) or {}
    log_lines: list[str] = [f"Job: {job['id']}", f"Task: {task}", f"Params: {params}"]

    try:
        if task == "ping_status":
            session = get_first_available_session()
            status = read_sbar_status(session)
            log_lines.append("STATUS atual lido sem navegar no SAP.")
            return status or "STATUS vazio em wnd[0]/sbar", "\n".join(log_lines)

        if task == "open_transaction":
            status, log = _open_transaction(params)
            log_lines.append(log)
            return status, "\n".join(log_lines)

        if task == "sap_cockpit":
            status, log = _run_sap_cockpit(params)
            log_lines.append(log)
            return status or "Execucao concluida, mas STATUS veio vazio.", "\n".join(log_lines)

        raise SapExecutionError(f"Rotina desconhecida: {task}")

    except Exception:
        log_lines.append(traceback.format_exc())
        raise
