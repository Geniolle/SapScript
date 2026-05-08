from __future__ import annotations

import importlib
import json
import os
import sys
import traceback
from typing import Any

import pythoncom
import win32com.client


class SapExecutionError(Exception):
    pass


WORKER_DIR = os.path.dirname(os.path.abspath(__file__))
WORKER_STATE_PATH = os.path.join(WORKER_DIR, ".sap_script_web_worker_state.json")


def _prepare_project_imports() -> None:
    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if project_dir and project_dir not in sys.path:
        sys.path.insert(0, project_dir)


def _load_worker_state() -> dict[str, Any]:
    if not os.path.exists(WORKER_STATE_PATH):
        return {}

    try:
        with open(WORKER_STATE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass

    return {}


def _save_worker_state(state: dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(WORKER_STATE_PATH), exist_ok=True)

    with open(WORKER_STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def _get_last_excel_dir() -> str:
    state = _load_worker_state()
    last_dir = str(state.get("last_excel_dir") or "").strip()

    if last_dir and os.path.isdir(last_dir):
        return last_dir

    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if project_dir and os.path.isdir(project_dir):
        return project_dir

    return os.path.expanduser("~")


def _set_last_excel_dir(file_path: str) -> None:
    folder = os.path.dirname(os.path.abspath(file_path))

    if not os.path.isdir(folder):
        return

    state = _load_worker_state()
    state["last_excel_dir"] = folder
    _save_worker_state(state)


def select_excel_file_on_windows(params: dict[str, Any] | None = None) -> tuple[str, str]:
    """
    Abre uma janela nativa do Windows para escolher ficheiro Excel.

    Importante:
    - Esta função roda no worker Windows, não no Docker.
    - O browser não consegue obter o caminho real do ficheiro por segurança.
    - Por isso o caminho completo vem daqui, do worker.
    - A última pasta usada fica guardada em worker/.sap_script_web_worker_state.json.
    """
    params = params or {}

    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise SapExecutionError(
            "Não foi possível abrir o seletor de ficheiros. "
            "Confirma se o Python do worker tem tkinter disponível."
        ) from exc

    initial_dir = str(params.get("initial_dir") or "").strip()

    if not initial_dir or not os.path.isdir(initial_dir):
        initial_dir = _get_last_excel_dir()

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        selected_path = filedialog.askopenfilename(
            title="Selecione o ficheiro Excel para o SAP Script",
            initialdir=initial_dir,
            filetypes=(
                ("Ficheiros Excel", "*.xlsx *.xlsm *.xls"),
                ("Todos os ficheiros", "*.*"),
            ),
        )
    finally:
        root.destroy()

    selected_path = str(selected_path or "").strip()

    if not selected_path:
        raise SapExecutionError("Seleção de ficheiro cancelada pelo utilizador.")

    if not os.path.exists(selected_path):
        raise SapExecutionError(f"Ficheiro selecionado não existe: {selected_path}")

    _set_last_excel_dir(selected_path)

    log = (
        "Ficheiro Excel selecionado no Windows.\n"
        f"Caminho: {selected_path}\n"
        f"Última pasta guardada: {os.path.dirname(os.path.abspath(selected_path))}"
    )

    return selected_path, log


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
        if task == "select_excel_file":
            status, log = select_excel_file_on_windows(params)
            log_lines.append(log)
            return status, "\n".join(log_lines)

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
            os.environ["SAP_JOB_ID"] = str(job["id"])
            os.environ["SAP_API_BASE_URL"] = os.getenv("API_BASE_URL", "http://localhost:8000").rstrip("/")
            os.environ["SAP_WORKER_TOKEN"] = os.getenv("WORKER_TOKEN", "change-me")
            status, log = _run_sap_cockpit(params)
            log_lines.append(log)
            return status or "Execucao concluida, mas STATUS veio vazio.", "\n".join(log_lines)

        raise SapExecutionError(f"Rotina desconhecida: {task}")

    except Exception:
        log_lines.append(traceback.format_exc())
        raise
