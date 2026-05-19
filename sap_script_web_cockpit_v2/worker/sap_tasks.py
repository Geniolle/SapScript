from __future__ import annotations

import importlib
import json
import os
import sys
import time
import traceback
from typing import Any

import pythoncom
import win32com.client
import queue
import threading
import requests
import ctypes


class SapExecutionError(Exception):
    pass


class JobCancelledException(BaseException):
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


def get_any_session() -> Any:
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except Exception as exc:
        raise SapExecutionError("Nao foi possivel ligar ao SAP GUI.") from exc

    for connection_index in range(application.Children.Count):
        connection = application.Children(connection_index)
        for session_index in range(connection.Children.Count):
            try:
                session = connection.Children(session_index)
                return session
            except Exception:
                continue

    raise SapExecutionError("Nao existe nenhuma sessao SAP.")


def _force_terminate_worker() -> None:
    try:
        import subprocess
        # Procura e encerra o processo PowerShell supervisor para este workspace
        cmd = "powershell.exe -Command \"Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*sap_script_web_cockpit_v2*start_worker_auto.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }\""
        subprocess.run(cmd, shell=True)
    except Exception:
        pass
    os._exit(1)


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


def _run_sap_search_requests(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    caminho = os.path.join(project_dir, "Processos", "pesquisar_request.py")
    if not os.path.exists(caminho):
        raise SapExecutionError(f"Nao encontrei o ficheiro pesquisar_request.py no caminho: {caminho}")
        
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("pesquisar_request", caminho)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    except Exception as exc:
        raise SapExecutionError(f"Falha ao carregar modulo pesquisar_request.py: {exc}")
        
    ambiente = str(params.get("ambiente") or "DEV").upper()
    mapa_sistema = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}
    sistema_desejado = mapa_sistema.get(ambiente, "S4D")
    
    try:
        lista = mod.listar_requests(
            system_name=sistema_desejado,
            max_rows="5000",
            include_requests=False,
            use_new_mode=True,
            minimize=True,
            close_after=True,
        )
    except Exception as exc:
        raise SapExecutionError(f"Erro ao pesquisar requests no SAP: {exc}")
        
    if not lista:
        return "[]", f"Pesquisa concluida. Nenhuma request encontrada para o sistema {sistema_desejado}."
        
    itens = [{"trkorr": item[0], "as4text": item[1]} for item in lista]
    status_json = json.dumps(itens)
    
    log = f"Pesquisa concluida com sucesso. Encontradas {len(lista)} requests."
    return status_json, log


def run_sap_task(job: dict[str, Any]) -> tuple[str, str]:
    task = job["task"]
    params = job.get("params", {}) or {}
    log_lines: list[str] = [f"Job: {job['id']}", f"Task: {task}", f"Params: {params}"]

    try:
        if task == "sap_search_requests":
            status, log = _run_sap_search_requests(params)
            log_lines.append(log)
            return status, "\n".join(log_lines)

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
            
            main_thread_id = threading.get_ident()
            
            class APILogStream:
                def __init__(self, job_id, original_stream, main_thread_id):
                    self.job_id = job_id
                    self.original = original_stream
                    self.main_thread_id = main_thread_id
                    self.queue = queue.Queue()
                    self.running = True
                    self.buffer = ""
                    self.api_url = os.environ["SAP_API_BASE_URL"]
                    self.token = os.environ["SAP_WORKER_TOKEN"]
                    self.thread = threading.Thread(target=self._sender_loop, daemon=True)
                    self.thread.start()

                def _is_progress_line(self, line: str) -> bool:
                    import re
                    clean = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', line)
                    if '\r' in clean:
                        return True
                    if any(c in clean for c in ['━', '█', '░', '▒', '▓', '▕', '▏']):
                        return True
                    if re.search(r'\d+%\s*\(\d+/\d+\)', clean):
                        return True
                    return False

                def write(self, data):
                    self.original.write(data)
                    self.buffer += data
                    if '\n' in self.buffer:
                        lines = self.buffer.split('\n')
                        self.buffer = lines.pop()
                        for line in lines:
                            cleaned_line = line.strip()
                            if cleaned_line and not self._is_progress_line(cleaned_line):
                                self.queue.put(cleaned_line)

                def flush(self):
                    self.original.flush()
                    cleaned_line = self.buffer.strip()
                    if cleaned_line and not self._is_progress_line(cleaned_line):
                        self.queue.put(cleaned_line)
                        self.buffer = ""

                def _sender_loop(self):
                    while self.running or not self.queue.empty():
                        try:
                            # Aguarda no máximo 0.5s por uma nova linha
                            first_line = self.queue.get(timeout=0.5)
                            lines = [first_line]
                            
                            # Esvazia a fila para agrupar o máximo de linhas possível (até 50 linhas)
                            while len(lines) < 50:
                                try:
                                    lines.append(self.queue.get_nowait())
                                except queue.Empty:
                                    break
                            
                            batch_data = "\n".join(lines)
                            try:
                                r = requests.post(
                                    f"{self.api_url}/api/jobs/{self.job_id}/log",
                                    headers={"X-Worker-Token": self.token},
                                    json={"log_line": batch_data},
                                    timeout=5
                                )
                                if r.status_code == 409:
                                    ctypes.pythonapi.PyThreadState_SetAsyncExc(
                                        ctypes.c_long(self.main_thread_id),
                                        ctypes.py_object(JobCancelledException)
                                    )
                                    time.sleep(1.5)
                                    self.original.write("\n⚠️ Log stream detectou cancelamento. A fechar PowerShell e a terminar o worker...\n")
                                    self.original.flush()
                                    try:
                                        pythoncom.CoInitialize()
                                        session = get_any_session()
                                        if session:
                                            conn = session.Parent
                                            conn.CloseSession(session.Id)
                                    except Exception:
                                        pass
                                    _force_terminate_worker()
                            except Exception as le:
                                self.original.write(f"\n[DEBUG LOG STREAM] Erro: {le}\n")
                                self.original.flush()
                        except queue.Empty:
                            pass

                def close(self):
                    self.flush()
                    self.running = False
                    self.thread.join(timeout=2.0)

            cancel_event = threading.Event()

            def poll_status():
                api_url = os.environ["SAP_API_BASE_URL"]
                token = os.environ["SAP_WORKER_TOKEN"]
                while not cancel_event.is_set():
                    try:
                        r = requests.get(
                            f"{api_url}/api/jobs/{job['id']}",
                            headers={"X-Worker-Token": token},
                            timeout=5
                        )
                        if r.status_code == 200:
                            job_data = r.json()
                            if job_data.get("state") == "failed" and "cancel" in str(job_data.get("status", "")).lower():
                                ctypes.pythonapi.PyThreadState_SetAsyncExc(
                                    ctypes.c_long(main_thread_id),
                                    ctypes.py_object(JobCancelledException)
                                )
                                for _ in range(15):
                                    if cancel_event.is_set():
                                        return
                                    time.sleep(0.1)
                                print("\n⚠️ Poller detectou cancelamento e processo principal bloqueado. A fechar PowerShell e a terminar o worker...")
                                sys.stdout.flush()
                                try:
                                    pythoncom.CoInitialize()
                                    session = get_any_session()
                                    if session:
                                        conn = session.Parent
                                        conn.CloseSession(session.Id)
                                except Exception:
                                    pass
                                _force_terminate_worker()
                                break
                    except Exception as pe:
                        print(f"\n[DEBUG POLLER] Erro ao consultar estado do job: {pe}")
                        sys.stdout.flush()
                    cancel_event.wait(2.0)

            poller_thread = threading.Thread(target=poll_status, daemon=True)
            poller_thread.start()

            orig_stdout = sys.stdout
            streamer = APILogStream(job["id"], orig_stdout, main_thread_id)
            sys.stdout = streamer

            try:
                status, log = _run_sap_cockpit(params)
            except JobCancelledException:
                print("\n❌ Execução cancelada pelo utilizador. A abortar transações SAP...")
                try:
                    session = get_any_session()
                    if session:
                        while len(session.Children) > 1:
                            try:
                                top_wnd = session.Children(len(session.Children) - 1)
                                top_wnd.close()
                            except Exception:
                                break
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                        session.findById("wnd[0]").sendVKey(0)
                        # Fechar a própria conexão da sessão para fechar a janela SAP correspondente
                        conn = session.Parent
                        conn.CloseSession(session.Id)
                except Exception:
                    pass
                status = "Cancelado"
                log = "Execução cancelada pelo utilizador."
                _force_terminate_worker()
            finally:
                cancel_event.set()
                sys.stdout = orig_stdout
                streamer.close()

            log_lines.append(log)
            return status or "Execucao concluida, mas STATUS veio vazio.", "\n".join(log_lines)

        raise SapExecutionError(f"Rotina desconhecida: {task}")

    except Exception:
        log_lines.append(traceback.format_exc())
        raise
