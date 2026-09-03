from __future__ import annotations

import importlib
import dataclasses
import json
import os
import subprocess
import sys
import time
import traceback
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client
from pathlib import Path

_project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
if _project_dir and _project_dir not in sys.path:
    sys.path.insert(0, _project_dir)

from pfcg.pfcg_create_excel_analyzer import analyze_pfcg_create_excel
import queue
import threading
import requests
import ctypes
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class SapExecutionError(Exception):
    pass


class JobCancelledException(BaseException):
    pass


WORKER_DIR = os.path.dirname(os.path.abspath(__file__))
WORKER_STATE_PATH = os.path.join(WORKER_DIR, ".sap_script_web_worker_state.json")
RFC_VENV_RELATIVE_PYTHON = Path(".venv-rfc") / "Scripts" / "python.exe"
RFC_SDK_HOME = r"C:\nwrfcsdk"
RFC_ALLOWED_EXIT_CODES = {0, 1}
API_CONNECT_TIMEOUT = float(os.getenv("WORKER_API_CONNECT_TIMEOUT", "3"))
API_READ_TIMEOUT = float(os.getenv("WORKER_API_READ_TIMEOUT", "15"))
API_RETRY_SESSION = requests.Session()
API_RETRY_SESSION.headers.update({"Connection": "keep-alive"})
_API_RETRY = Retry(
    total=2,
    connect=2,
    read=2,
    status=2,
    backoff_factor=0.25,
    allowed_methods=frozenset({"GET", "POST"}),
)
API_RETRY_SESSION.mount("http://", HTTPAdapter(max_retries=_API_RETRY))
API_RETRY_SESSION.mount("https://", HTTPAdapter(max_retries=_API_RETRY))


def _prepare_project_imports() -> None:
    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if project_dir and project_dir not in sys.path:
        sys.path.insert(0, project_dir)
    cockpit_dir = Path(project_dir) / "sap_script_web_cockpit_v2" if project_dir else None
    if cockpit_dir and cockpit_dir.exists():
        cockpit_dir_str = str(cockpit_dir)
        if cockpit_dir_str not in sys.path:
            sys.path.insert(0, cockpit_dir_str)


def _parse_env_line(raw: str) -> tuple[str | None, str | None]:
    line = str(raw or "").strip()
    if not line or line.startswith("#") or "=" not in line:
        return None, None

    key, value = line.split("=", 1)
    key = key.strip()
    value = value.strip()
    if not key:
        return None, None

    if len(value) >= 2 and (
        (value.startswith('"') and value.endswith('"'))
        or (value.startswith("'") and value.endswith("'"))
    ):
        value = value[1:-1]

    return key, value


def _load_project_env_manual(project_dir: str) -> None:
    env_path = Path(project_dir) / ".env"
    if not env_path.exists():
        return

    with env_path.open("r", encoding="utf-8-sig") as file_obj:
        for raw in file_obj:
            key, value = _parse_env_line(raw)
            if key and key not in os.environ:
                os.environ[key] = value or ""


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


def _resolve_api_base_url() -> str:
    configured = os.getenv("API_BASE_URL", "").strip().rstrip("/")
    if configured:
        return configured
    return "http://localhost:8010"


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


def _run_pfcg_create_excel_analysis(params: dict[str, Any]) -> tuple[str, str]:
    excel_path = str(params.get("excel_path") or "").strip()
    role_name = str(params.get("role_name") or "PFCG_CREATE").strip() or "PFCG_CREATE"
    result = analyze_pfcg_create_excel(excel_path=excel_path, expected_role_name=role_name)
    if result.get("ok") is True:
        log = (
            "Analise read-only do Excel concluida com sucesso.\n"
            f"Role: {result.get('role')}\n"
            f"Ficheiro: {Path(excel_path).name if excel_path else ''}"
        )
    else:
        log = (
            "Analise read-only do Excel terminou com problemas.\n"
            f"Role: {result.get('role')}\n"
            f"Erro: {result.get('message') or ', '.join(result.get('errors') or [])}"
        )
    return json.dumps(result, ensure_ascii=False), log


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
        importlib.reload(cockpit)
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


def _run_sap_agent_analysis(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    ticket_key = str(params.get("ticket_key") or "").strip()
    if not ticket_key:
        raise SapExecutionError("Chave de ticket vazia.")
    
    try:
        from sap_agent.runner import build_engine
        from sap_agent.jira_client import JiraClient
        from sap_agent.config import JiraConfig as SapJiraConfig
        import dataclasses
        
        project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
        config_path = os.path.join(project_dir, "config", "sap_agent.yaml")
        
        engine, agent_config, jira_config = build_engine(config_path)
        
        temp_jira_config = SapJiraConfig(
            base_url=jira_config.base_url,
            email=jira_config.email,
            api_token=jira_config.api_token,
            jql=f"key = {ticket_key}",
            max_results=1,
            update_jira=False
        )
        
        jira = JiraClient(temp_jira_config)
        tickets = jira.search_tickets()
        if not tickets:
            raise SapExecutionError(f"Ticket {ticket_key} não encontrado no JIRA.")
            
        ticket = tickets[0]
        diagnosis = engine.diagnose(ticket)
        
        def default_serializer(o):
            if dataclasses.is_dataclass(o):
                return dataclasses.asdict(o)
            return str(o)
            
        result_json = json.dumps(diagnosis, default=default_serializer, ensure_ascii=False)
        return result_json, f"Análise do ticket {ticket_key} concluída com sucesso."
    except Exception as exc:
        raise SapExecutionError(f"Erro ao executar análise do Agente SAP: {exc}")


def _get_project_dir() -> Path:
    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if not project_dir:
        raise SapExecutionError("SAP_SCRIPT_PROJECT_DIR não definido.")
    path = Path(project_dir).resolve()
    if not path.exists():
        raise SapExecutionError(f"SAP_SCRIPT_PROJECT_DIR não existe: {path}")
    return path


def _build_rfc_bridge_env(project_dir: Path) -> dict[str, str]:
    env = os.environ.copy()
    sdk_home = (
        str(env.get("SAP_NWRFC_HOME") or "").strip()
        or str(env.get("SAPNWRFC_HOME") or "").strip()
        or RFC_SDK_HOME
    )
    env["SAP_SCRIPT_PROJECT_DIR"] = str(project_dir)
    env["SAPNWRFC_HOME"] = sdk_home
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    sdk_lib = str(Path(sdk_home) / "lib")
    current_path = env.get("PATH", "")
    if sdk_lib.lower() not in current_path.lower():
        path_entries = [sdk_lib]
        if current_path:
            path_entries.append(current_path)
        env["PATH"] = os.pathsep.join(path_entries)
    return env


def _run_pfcg_role_analysis(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()

    try:
        from sap_rfc.pfcg_role_service import validate_role_name
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar a validação PFCG: {exc}") from exc

    project_dir = _get_project_dir()
    raw_role_name = str(params.get("role_name") or "")
    try:
        role_name = validate_role_name(raw_role_name)
    except ValueError as exc:
        payload = {
            "ok": False,
            "status": "ERRO",
            "role": raw_role_name.strip().upper(),
            "error_type": "INVALID_INPUT",
            "message": str(exc),
            "system": "PRD",
            "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
        }
        log = (
            "Análise PFCG rejeitada antes da bridge RFC.\n"
            f"Role: {raw_role_name}\n"
            f"Motivo: {exc}"
        )
        return json.dumps(payload, ensure_ascii=False), log

    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise SapExecutionError(f"Python RFC não encontrado: {rfc_python}")

    env = _build_rfc_bridge_env(project_dir)
    env["PFCG_TARGET_ENV"] = str(params.get("system") or "PRD").strip().upper() or "PRD"
    command = [
        str(rfc_python),
        "-m",
        "sap_rfc.pfcg_role_cli",
        "--role-name",
        role_name,
    ]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise SapExecutionError(
            f"Timeout ao executar análise RFC PFCG para a função {role_name}."
        ) from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise SapExecutionError(
            f"Bridge RFC PFCG falhou com exit code {run.returncode}."
        )
    if not stdout:
        raise SapExecutionError("Bridge RFC PFCG não devolveu JSON em stdout.")

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise SapExecutionError("Bridge RFC PFCG devolveu JSON inválido.") from exc

    if not isinstance(payload, dict):
        raise SapExecutionError("Bridge RFC PFCG devolveu payload inválido.")

    log_lines = [
        "Análise PFCG executada via subprocesso RFC controlado.",
        f"Role: {role_name}",
        f"Python RFC: {rfc_python}",
        f"Exit code: {run.returncode}",
        f"Status: {payload.get('status', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")
    if payload.get("warning"):
        log_lines.append(f"Aviso: {payload['warning']}")
    if stderr:
        log_lines.append(f"stderr: {stderr}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_sub_analysis(
    params: dict[str, Any],
    *,
    cli_module: str,
    label: str,
) -> tuple[str, str]:
    """Executa uma sub-análise PFCG (transações/utilizadores) via bridge RFC isolada.

    Aceita EXCLUSIVAMENTE `role_name` de `params` — qualquer outra chave (task,
    script, module, executable, command, table, ...) vinda do frontend/job é
    ignorada; a tabela SAP consultada é decidida apenas pelo serviço Python
    invocado (sap_rfc.*), nunca pelo payload do pedido.
    """
    _prepare_project_imports()

    try:
        from sap_rfc.pfcg_role_service import validate_role_name
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar a validação PFCG: {exc}") from exc

    project_dir = _get_project_dir()
    raw_role_name = str(params.get("role_name") or "")
    try:
        role_name = validate_role_name(raw_role_name)
    except ValueError as exc:
        payload = {
            "ok": False,
            "status": "ERRO",
            "role": raw_role_name.strip().upper(),
            "error_type": "INVALID_INPUT",
            "message": str(exc),
            "system": "PRD",
            "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
        }
        log = (
            f"Análise PFCG ({label}) rejeitada antes da bridge RFC.\n"
            f"Role: {raw_role_name}\n"
            f"Motivo: {exc}"
        )
        return json.dumps(payload, ensure_ascii=False), log

    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise SapExecutionError(f"Python RFC não encontrado: {rfc_python}")

    env = _build_rfc_bridge_env(project_dir)
    env["PFCG_TARGET_ENV"] = str(params.get("system") or "PRD").strip().upper() or "PRD"
    command = [
        str(rfc_python),
        "-m",
        cli_module,
        "--role-name",
        role_name,
    ]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise SapExecutionError(
            f"Timeout ao executar análise PFCG ({label}) para a função {role_name}."
        ) from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise SapExecutionError(
            f"Bridge RFC PFCG ({label}) falhou com exit code {run.returncode}."
        )
    if not stdout:
        raise SapExecutionError(f"Bridge RFC PFCG ({label}) não devolveu JSON em stdout.")

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise SapExecutionError(f"Bridge RFC PFCG ({label}) devolveu JSON inválido.") from exc

    if not isinstance(payload, dict):
        raise SapExecutionError(f"Bridge RFC PFCG ({label}) devolveu payload inválido.")

    log_lines = [
        f"Análise PFCG ({label}) executada via subprocesso RFC controlado.",
        f"Role: {role_name}",
        f"Python RFC: {rfc_python}",
        f"Exit code: {run.returncode}",
        f"Status: {payload.get('status', '-')}",
        f"Count: {payload.get('count', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")
    if payload.get("warning"):
        log_lines.append(f"Aviso: {payload['warning']}")
    if stderr:
        log_lines.append(f"stderr: {stderr}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_transactions_analysis(params: dict[str, Any]) -> tuple[str, str]:
    return _run_pfcg_role_sub_analysis(
        params,
        cli_module="sap_rfc.pfcg_role_transactions_cli",
        label="transações",
    )


def _run_pfcg_role_users_analysis(params: dict[str, Any]) -> tuple[str, str]:
    return _run_pfcg_role_sub_analysis(
        params,
        cli_module="sap_rfc.pfcg_role_users_cli",
        label="utilizadores",
    )


def _run_pfcg_transaction_roles(params: dict[str, Any]) -> tuple[str, str]:
    """Read-only via bridge RFC isolada: funcoes PFCG a que uma transacao esta
    atribuida em PRD. Aceita EXCLUSIVAMENTE `tcode` de params."""
    _prepare_project_imports()
    project_dir = _get_project_dir()
    raw_tcode = str(params.get("tcode") or "").strip().upper()

    try:
        from sap_rfc.pfcg_transaction_roles_service import validate_tcode
        tcode = validate_tcode(raw_tcode)
    except ValueError as exc:
        payload = {
            "ok": False,
            "status": "ERRO",
            "tcode": raw_tcode,
            "error_type": "INVALID_INPUT",
            "message": str(exc),
            "system": "PRD",
            "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
        }
        return json.dumps(payload, ensure_ascii=False), (
            f"Analise 'transacao -> funcoes' rejeitada antes da bridge RFC.\nTcode: {raw_tcode}\nMotivo: {exc}"
        )
    except Exception as exc:
        raise SapExecutionError(f"Nao foi possivel importar a validacao de transacao: {exc}") from exc

    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise SapExecutionError(f"Python RFC nao encontrado: {rfc_python}")

    env = _build_rfc_bridge_env(project_dir)
    env["PFCG_TARGET_ENV"] = str(params.get("system") or "PRD").strip().upper() or "PRD"
    command = [str(rfc_python), "-m", "sap_rfc.pfcg_transaction_roles_cli", "--tcode", tcode]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise SapExecutionError(f"Timeout ao consultar funcoes da transacao {tcode}.") from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise SapExecutionError(f"Bridge RFC (transacao -> funcoes) falhou com exit code {run.returncode}.")
    if not stdout:
        raise SapExecutionError("Bridge RFC (transacao -> funcoes) nao devolveu JSON em stdout.")

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise SapExecutionError("Bridge RFC (transacao -> funcoes) devolveu JSON invalido.") from exc
    if not isinstance(payload, dict):
        raise SapExecutionError("Bridge RFC (transacao -> funcoes) devolveu payload invalido.")

    log_lines = [
        "Analise 'transacao -> funcoes' via subprocesso RFC controlado.",
        f"Tcode: {tcode}",
        f"Python RFC: {rfc_python}",
        f"Exit code: {run.returncode}",
        f"Status: {payload.get('status', '-')}",
        f"Count: {payload.get('count', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")
    if payload.get("warning"):
        log_lines.append(f"Aviso: {payload['warning']}")
    if stderr:
        log_lines.append(f"stderr: {stderr}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_object_roles(params: dict[str, Any]) -> tuple[str, str]:
    """Read-only via bridge RFC isolada: funcoes PFCG que contem um objeto de
    autorizacao em PRD. Aceita EXCLUSIVAMENTE `auth_object` de params."""
    _prepare_project_imports()
    project_dir = _get_project_dir()
    raw_obj = str(params.get("auth_object") or "").strip().upper()

    try:
        from sap_rfc.pfcg_object_roles_service import validate_auth_object
        auth_object = validate_auth_object(raw_obj)
    except ValueError as exc:
        payload = {
            "ok": False,
            "status": "ERRO",
            "auth_object": raw_obj,
            "error_type": "INVALID_INPUT",
            "message": str(exc),
            "system": "PRD",
            "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
        }
        return json.dumps(payload, ensure_ascii=False), (
            f"Analise 'objeto -> funcoes' rejeitada antes da bridge RFC.\nObjeto: {raw_obj}\nMotivo: {exc}"
        )
    except Exception as exc:
        raise SapExecutionError(f"Nao foi possivel importar a validacao de objeto: {exc}") from exc

    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise SapExecutionError(f"Python RFC nao encontrado: {rfc_python}")

    env = _build_rfc_bridge_env(project_dir)
    env["PFCG_TARGET_ENV"] = str(params.get("system") or "PRD").strip().upper() or "PRD"
    command = [str(rfc_python), "-m", "sap_rfc.pfcg_object_roles_cli", "--object", auth_object]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise SapExecutionError(f"Timeout ao consultar funcoes do objeto {auth_object}.") from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise SapExecutionError(f"Bridge RFC (objeto -> funcoes) falhou com exit code {run.returncode}.")
    if not stdout:
        raise SapExecutionError("Bridge RFC (objeto -> funcoes) nao devolveu JSON em stdout.")

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise SapExecutionError("Bridge RFC (objeto -> funcoes) devolveu JSON invalido.") from exc
    if not isinstance(payload, dict):
        raise SapExecutionError("Bridge RFC (objeto -> funcoes) devolveu payload invalido.")

    log_lines = [
        "Analise 'objeto -> funcoes' via subprocesso RFC controlado.",
        f"Objeto: {auth_object}",
        f"Python RFC: {rfc_python}",
        f"Exit code: {run.returncode}",
        f"Status: {payload.get('status', '-')}",
        f"Count: {payload.get('count', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")
    if payload.get("warning"):
        log_lines.append(f"Aviso: {payload['warning']}")
    if stderr:
        log_lines.append(f"stderr: {stderr}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_user_roles(params: dict[str, Any]) -> tuple[str, str]:
    """Read-only via bridge RFC isolada: funcoes PFCG atribuidas a um utilizador
    em PRD. Aceita EXCLUSIVAMENTE `username` de params."""
    _prepare_project_imports()
    project_dir = _get_project_dir()
    raw_user = str(params.get("username") or "").strip().upper()

    try:
        from sap_rfc.pfcg_user_roles_service import validate_username
        username = validate_username(raw_user)
    except ValueError as exc:
        payload = {
            "ok": False,
            "status": "ERRO",
            "username": raw_user,
            "error_type": "INVALID_INPUT",
            "message": str(exc),
            "system": "PRD",
            "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
        }
        return json.dumps(payload, ensure_ascii=False), (
            f"Analise 'utilizador -> funcoes' rejeitada antes da bridge RFC.\nUtilizador: {raw_user}\nMotivo: {exc}"
        )
    except Exception as exc:
        raise SapExecutionError(f"Nao foi possivel importar a validacao de utilizador: {exc}") from exc

    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise SapExecutionError(f"Python RFC nao encontrado: {rfc_python}")

    env = _build_rfc_bridge_env(project_dir)
    env["PFCG_TARGET_ENV"] = str(params.get("system") or "PRD").strip().upper() or "PRD"
    command = [str(rfc_python), "-m", "sap_rfc.pfcg_user_roles_cli", "--user", username]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise SapExecutionError(f"Timeout ao consultar funcoes do utilizador {username}.") from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise SapExecutionError(f"Bridge RFC (utilizador -> funcoes) falhou com exit code {run.returncode}.")
    if not stdout:
        raise SapExecutionError("Bridge RFC (utilizador -> funcoes) nao devolveu JSON em stdout.")

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise SapExecutionError("Bridge RFC (utilizador -> funcoes) devolveu JSON invalido.") from exc
    if not isinstance(payload, dict):
        raise SapExecutionError("Bridge RFC (utilizador -> funcoes) devolveu payload invalido.")

    log_lines = [
        "Analise 'utilizador -> funcoes' via subprocesso RFC controlado.",
        f"Utilizador: {username}",
        f"Python RFC: {rfc_python}",
        f"Exit code: {run.returncode}",
        f"Status: {payload.get('status', '-')}",
        f"Count: {payload.get('count', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")
    if payload.get("warning"):
        log_lines.append(f"Aviso: {payload['warning']}")
    if stderr:
        log_lines.append(f"stderr: {stderr}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_create_preview(params: dict[str, Any]) -> tuple[str, str]:
    """Pré-visualização (read-only) da criação individual de função PFCG via RFC.

    Aceita EXCLUSIVAMENTE environment/role_name/description/tcodes/transport_mode/
    request_number/request_description vindos de `params`. Nenhuma outra chave do
    payload é usada; a função SAP a chamar é decidida apenas dentro de
    sap_rfc.pfcg_role_create_service, nunca pelo pedido do frontend.
    """
    _prepare_project_imports()
    project_dir = _get_project_dir()

    environment = str(params.get("environment") or "").strip().upper()
    role_name = str(params.get("role_name") or "").strip()
    description = str(params.get("description") or "").strip()
    raw_tcodes = params.get("tcodes") or []
    tcodes = [str(t).strip() for t in raw_tcodes if str(t).strip()] if isinstance(raw_tcodes, list) else []
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip()
    request_description = str(params.get("request_description") or "").strip()

    try:
        from pfcg.pfcg_create_rfc_service import preview_pfcg_role_create_rfc
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço de pré-visualização PFCG: {exc}") from exc

    try:
        payload = preview_pfcg_role_create_rfc(
            environment, role_name, description, tcodes, transport_mode, request_number, request_description
        )
    except Exception as exc:
        raise SapExecutionError(f"Bridge de pré-visualização PFCG (RFC) falhou: {exc}") from exc

    log_lines = [
        "Pré-visualização de criação individual PFCG (RFC) executada via subprocesso isolado.",
        f"Ambiente: {environment}",
        f"Role: {role_name}",
        f"Python RFC: {project_dir / RFC_VENV_RELATIVE_PYTHON}",
        f"Status: {payload.get('status', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_create_rfc(params: dict[str, Any]) -> tuple[str, str]:
    """Criação individual REAL de função PFCG via RFC. Só é permitida em DEV.

    Aceita EXCLUSIVAMENTE environment/role_name/description/tcodes/transport_mode/
    request_number/request_description vindos de `params` — o frontend nunca pode
    enviar function_module/command/script/executable/module/table/shell/python_path;
    esta função ignora qualquer chave além dessas sete.
    """
    _prepare_project_imports()
    project_dir = _get_project_dir()

    environment = str(params.get("environment") or "").strip().upper()
    role_name = str(params.get("role_name") or "").strip()
    description = str(params.get("description") or "").strip()
    raw_tcodes = params.get("tcodes") or []
    tcodes = [str(t).strip() for t in raw_tcodes if str(t).strip()] if isinstance(raw_tcodes, list) else []
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip()
    request_description = str(params.get("request_description") or "").strip()

    try:
        from pfcg.pfcg_create_rfc_service import create_pfcg_role_rfc
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço de criação PFCG: {exc}") from exc

    try:
        payload = create_pfcg_role_rfc(
            environment, role_name, description, tcodes, transport_mode, request_number, request_description
        )
    except Exception as exc:
        raise SapExecutionError(f"Bridge de criação PFCG (RFC) falhou: {exc}") from exc

    log_lines = [
        "Criação individual PFCG (RFC) executada via subprocesso isolado.",
        f"Ambiente: {environment}",
        f"Role: {role_name}",
        f"Python RFC: {project_dir / RFC_VENV_RELATIVE_PYTHON}",
        f"Status: {payload.get('status', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_delete_preview(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    project_dir = _get_project_dir()

    environment = str(params.get("environment") or "").strip().upper()
    role_name = str(params.get("role_name") or "").strip()
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip()
    request_description = str(params.get("request_description") or "").strip()

    try:
        from pfcg.pfcg_delete_rfc_service import preview_pfcg_role_delete_rfc
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço de pré-visualização de eliminação PFCG: {exc}") from exc

    try:
        payload = preview_pfcg_role_delete_rfc(
            environment, role_name, transport_mode, request_number, request_description
        )
    except Exception as exc:
        raise SapExecutionError(f"Bridge de pré-visualização de eliminação PFCG (RFC) falhou: {exc}") from exc

    log_lines = [
        "Pré-visualização de eliminação PFCG (RFC) executada via subprocesso isolado.",
        f"Ambiente: {environment}",
        f"Role: {role_name}",
        f"Python RFC: {project_dir / RFC_VENV_RELATIVE_PYTHON}",
        f"Status: {payload.get('status', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_role_delete_rfc(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    project_dir = _get_project_dir()

    environment = str(params.get("environment") or "").strip().upper()
    role_name = str(params.get("role_name") or "").strip()
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip()
    request_description = str(params.get("request_description") or "").strip()

    try:
        from pfcg.pfcg_delete_rfc_service import delete_pfcg_role_rfc
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço de eliminação PFCG: {exc}") from exc

    try:
        payload = delete_pfcg_role_rfc(
            environment, role_name, transport_mode, request_number, request_description
        )
    except Exception as exc:
        raise SapExecutionError(f"Bridge de eliminação PFCG (RFC) falhou: {exc}") from exc

    log_lines = [
        "Eliminação individual PFCG (RFC) executada via subprocesso isolado.",
        f"Ambiente: {environment}",
        f"Role: {role_name}",
        f"Python RFC: {project_dir / RFC_VENV_RELATIVE_PYTHON}",
        f"Status: {payload.get('status', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_transport_search(params: dict[str, Any]) -> tuple[str, str]:
    """Pesquisa (read-only via RFC) das Requests de transporte abertas do utilizador RFC.

    Aceita EXCLUSIVAMENTE environment vindo de `params`. Substitui, para o fluxo
    "Criar Individualmente", a pesquisa GUI/SE16H de Processos/pesquisar_request.py por
    uma chamada RFC real (CTS_WBO_SELECT_REQUESTS) com o mesmo critério funcional
    (Requests abertas pertencentes ao utilizador).
    """
    _prepare_project_imports()

    environment = str(params.get("environment") or "").strip().upper()

    try:
        from pfcg.pfcg_create_rfc_service import search_transport_requests_rfc
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço de pesquisa de Requests: {exc}") from exc

    try:
        payload = search_transport_requests_rfc(environment)
    except Exception as exc:
        raise SapExecutionError(f"Bridge de pesquisa de Requests (RFC) falhou: {exc}") from exc

    log_lines = [
        "Pesquisa de Requests de transporte abertas (RFC) executada via subprocesso isolado.",
        f"Ambiente: {environment}",
        f"Status: {payload.get('status', '-')}",
        f"Requests encontradas: {payload.get('requests_count', '-')}",
    ]
    if payload.get("message"):
        log_lines.append(f"Mensagem: {payload['message']}")

    return json.dumps(payload, ensure_ascii=False), "\n".join(log_lines)


def _run_pfcg_composta_create_preview(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    environment = str(params.get("environment") or "DEV").strip().upper()
    role_name = str(params.get("role_name") or "").strip().upper()
    description = str(params.get("description") or "").strip()
    raw_child_roles = params.get("child_roles") or []
    child_roles = [str(r).strip().upper() for r in raw_child_roles if str(r).strip()] if isinstance(raw_child_roles, list) else []
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip().upper()
    request_description = str(params.get("request_description") or "").strip()

    if not role_name:
        payload = {"ok": False, "status": "ERRO", "error_type": "INVALID_INPUT", "message": "Informe o nome da Função Composta."}
        return json.dumps(payload, ensure_ascii=False), "Nome de role vazio."

    if not description:
        payload = {"ok": False, "status": "ERRO", "error_type": "INVALID_INPUT", "message": "Informe a descrição da Função Composta."}
        return json.dumps(payload, ensure_ascii=False), "Descrição vazia."

    if not child_roles:
        payload = {"ok": False, "status": "ERRO", "error_type": "INVALID_INPUT", "message": "Informe pelo menos uma função componente (role filha)."}
        return json.dumps(payload, ensure_ascii=False), "Lista de roles filhas vazia."

    try:
        from sap_rfc.pfcg_transport_service import validate_transport_inputs
        transport = validate_transport_inputs(transport_mode, request_number, request_description)
    except Exception as exc:
        payload = {"ok": False, "status": "ERRO", "error_type": "INVALID_TRANSPORT_INPUT", "message": str(exc)}
        return json.dumps(payload, ensure_ascii=False), f"Validação de transporte falhou: {exc}"

    payload = {
        "ok": True,
        "status": "PREVIEW_READY",
        "environment": environment,
        "role": role_name,
        "description": description,
        "child_roles": child_roles,
        "tipo": "Função Composta",
        "system": "DEV",
        "client": "100",
        "transport": {
            "transport_mode": transport["transport_mode"],
            "request_number": transport["request_number"],
            "request_description": transport["request_description"],
        },
    }
    log = f"Pré-visualização da Função Composta {role_name} pronta com {len(child_roles)} roles filhas."
    return json.dumps(payload, ensure_ascii=False), log


def _run_pfcg_composta_create(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    import openpyxl
    import tempfile

    environment = str(params.get("environment") or "DEV").strip().upper()
    role_name = str(params.get("role_name") or "").strip().upper()
    description = str(params.get("description") or "").strip()
    raw_child_roles = params.get("child_roles") or []
    child_roles = [str(r).strip().upper() for r in raw_child_roles if str(r).strip()] if isinstance(raw_child_roles, list) else []
    transport_mode = str(params.get("transport_mode") or "LOCAL").strip().upper()
    request_number = str(params.get("request_number") or "").strip().upper()
    request_description = str(params.get("request_description") or "").strip()

    if transport_mode == "CREATE_REQUEST":
        try:
            from sap_rfc.pfcg_transport_service import create_transport_request
            tr_result = create_transport_request(environment="DEV", description=request_description or f"Criar Funcao Composta {role_name}")
            request_number = tr_result.get("request_number", "")
        except Exception as exc:
            raise SapExecutionError(f"Falha ao criar Ordem de Transporte: {exc}") from exc

    temp_dir = tempfile.gettempdir()
    temp_excel_path = os.path.join(temp_dir, f"pfcg_composta_{role_name}_{int(time.time())}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["TEXT", "STATUS", "MSG", "TIMESTEMP", "AGR_NAME_COMPOSTA", "AGR_NAME"])
    ws.append([description, "", "", "", role_name, ", ".join(child_roles)])
    wb.save(temp_excel_path)
    wb.close()

    try:
        cockpit_params = {
            "ambiente": environment,
            "processo": "Funções PFCG",
            "subprocesso": "D. PFCG_COMPOSTA.py",
            "caminho_ficheiro": temp_excel_path,
            "request_transporte": request_number if transport_mode != "LOCAL" else "",
            "modo_nao_interativo": True,
            "pedir_confirmacao": False,
        }
        status_res, log_res = _run_sap_cockpit(cockpit_params)
        payload = {
            "ok": True,
            "status": "SUCESSO",
            "role": role_name,
            "description": description,
            "child_roles": child_roles,
            "request_number": request_number if transport_mode != "LOCAL" else None,
            "message": f"Função Composta {role_name} criada com sucesso em DEV.",
        }
        return json.dumps(payload, ensure_ascii=False), log_res or "Execução concluída com sucesso."
    finally:
        try:
            if os.path.exists(temp_excel_path):
                os.remove(temp_excel_path)
        except Exception:
            pass


def _run_sap_gui_chat_action(params: dict[str, Any]) -> tuple[str, str]:
    """Executa uma ação SAP GUI solicitada pelo chat (Gemini function calling).

    params deve conter:
      - action: "se16n_query" | "open_transaction" | "read_sbar"
      - Para se16n_query: table, filters (list), fields (list), max_rows
      - Para open_transaction: transaction
      - description: texto descritivo opcional
    """
    _prepare_project_imports()
    try:
        from sap_agent.sap_gui_actions import execute_sap_gui_action
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar sap_gui_actions: {exc}") from exc

    # Usar ensure_sap_access para garantir que o SAP está aberto e com sessão ativa
    try:
        from sap_session import ensure_sap_access_from_env
        # Chave de sistema configurada no .env (S4PCLNT100 = PRODUÇÃO por padrão)
        sap_key = str(params.get("sap_key") or "S4PCLNT100").strip().upper()
        ensure_sap_access_from_env(key=sap_key)
    except Exception as exc:
        raise SapExecutionError(
            f"Não foi possível abrir/validar sessão SAP ({sap_key}): {exc}"
        ) from exc

    result = execute_sap_gui_action(params)

    # Serializar resultado para JSON (status) + log textual
    import json as _json
    import dataclasses as _dc

    status_payload = {
        "action": result.action,
        "description": result.description,
        "result_text": result.result_text,
        "rows": result.rows,
        "error": result.error,
        "success": result.success,
    }
    status_json = _json.dumps(status_payload, ensure_ascii=False)
    log = (
        f"SAP GUI Action: {result.action}\n"
        f"Descrição: {result.description}\n"
        f"Sucesso: {result.success}\n"
        f"Linhas retornadas: {len(result.rows)}\n"
        + (f"Erro: {result.error}" if result.error else "")
    )
    return status_json, log


def _run_fi_default_document(job: dict[str, Any], params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()

    try:
        from sap_rfc.fi_document_service import post_fi_document
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço FI: {exc}") from exc
    try:
        from sap_script_web_cockpit_v2.worker.fi_default_document_job import (
            FiDefaultDocumentJobError,
            run_fi_default_document_job,
        )
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o executor FI isolado: {exc}") from exc

    try:
        return run_fi_default_document_job(
            job_id=job["id"],
            params=params,
            post_fi_document=post_fi_document,
        )
    except FiDefaultDocumentJobError as exc:
        raise SapExecutionError(str(exc)) from exc


def _run_f110_proposal(job: dict[str, Any], params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()

    try:
        from sap_rfc.f110_service import run_f110_proposal
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço F110: {exc}") from exc

    try:
        from sap_script_web_cockpit_v2.worker.f110_proposal_job import (
            F110ProposalJobError,
            run_f110_proposal_job,
        )
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o executor F110 isolado: {exc}") from exc

    try:
        return run_f110_proposal_job(
            job_id=job["id"],
            params=params,
            run_f110_proposal=run_f110_proposal,
        )
    except F110ProposalJobError as exc:
        raise SapExecutionError(str(exc)) from exc


def _run_f110_payment(job: dict[str, Any], params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()

    try:
        from sap_rfc.f110_service import run_f110_payment
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o serviço F110 de pagamento: {exc}") from exc

    try:
        from sap_script_web_cockpit_v2.worker.f110_payment_job import (
            run_f110_payment_job,
        )
    except ImportError as exc:
        raise SapExecutionError(f"Não foi possível importar o executor F110 de pagamento: {exc}") from exc

    try:
        return run_f110_payment_job(
            job_id=job["id"],
            params=params,
            run_f110_payment=run_f110_payment,
        )
    except Exception as exc:
        raise SapExecutionError(str(exc)) from exc


def _handle_ping_status(job: dict[str, Any], params: dict[str, Any]) -> tuple[str, str]:
    session = get_first_available_session()
    status = read_sbar_status(session)
    return status or "STATUS vazio em wnd[0]/sbar", "STATUS atual lido sem navegar no SAP."


# Dispatch das tasks simples: task -> callable(job, params) -> (status, log).
# `sap_cockpit` fica FORA (streaming/threads/documentacao proprios em run_sap_task).
TASK_HANDLERS: dict[str, "Any"] = {
    "sap_agent_analysis": lambda job, params: _run_sap_agent_analysis(params),
    "sap_cockpit_auto_trigger": lambda job, params: _run_sap_cockpit(params),
    "pfcg_role_analysis": lambda job, params: _run_pfcg_role_analysis(params),
    "pfcg_role_transactions_analysis": lambda job, params: _run_pfcg_role_transactions_analysis(params),
    "pfcg_role_users_analysis": lambda job, params: _run_pfcg_role_users_analysis(params),
    "pfcg_transaction_roles": lambda job, params: _run_pfcg_transaction_roles(params),
    "pfcg_object_roles": lambda job, params: _run_pfcg_object_roles(params),
    "pfcg_user_roles": lambda job, params: _run_pfcg_user_roles(params),
    "pfcg_create_excel_analysis": lambda job, params: _run_pfcg_create_excel_analysis(params),
    "pfcg_role_create_preview": lambda job, params: _run_pfcg_role_create_preview(params),
    "pfcg_role_create_rfc": lambda job, params: _run_pfcg_role_create_rfc(params),
    "pfcg_composta_create_preview": lambda job, params: _run_pfcg_composta_create_preview(params),
    "pfcg_composta_create": lambda job, params: _run_pfcg_composta_create(params),
    "pfcg_role_delete_preview": lambda job, params: _run_pfcg_role_delete_preview(params),
    "pfcg_role_delete_rfc": lambda job, params: _run_pfcg_role_delete_rfc(params),
    "pfcg_transport_search": lambda job, params: _run_pfcg_transport_search(params),
    "sap_search_requests": lambda job, params: _run_sap_search_requests(params),
    "select_excel_file": lambda job, params: select_excel_file_on_windows(params),
    "ping_status": _handle_ping_status,
    "open_transaction": lambda job, params: _open_transaction(params),
    "sap_gui_chat_action": lambda job, params: _run_sap_gui_chat_action(params),
    "fi_default_document": _run_fi_default_document,
    "f110_proposal": _run_f110_proposal,
    "f110_payment": _run_f110_payment,
}



def run_sap_task(job: dict[str, Any]) -> tuple[str, str]:
    _project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if _project_dir:
        _load_project_env_manual(_project_dir)

    task = job["task"]
    params = job.get("params", {}) or {}
    log_lines: list[str] = [f"Job: {job['id']}", f"Task: {task}", f"Params: {params}"]

    try:
        handler = TASK_HANDLERS.get(task)
        if handler is not None:
            status, log = handler(job, params)
            log_lines.append(log)
            return status, "\n".join(log_lines)


        if task == "sap_cockpit":
            os.environ["SAP_JOB_ID"] = str(job["id"])
            os.environ["SAP_API_BASE_URL"] = _resolve_api_base_url()
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
                    # Janela de acumulação: aguarda até 300ms colhendo linhas antes de enviar.
                    # Isso reduz o número de requisições HTTP quando o script é muito verboso.
                    BATCH_WINDOW_S = 0.30
                    while self.running or not self.queue.empty():
                        lines = []
                        deadline = time.monotonic() + BATCH_WINDOW_S

                        # Colhe linhas durante a janela de tempo
                        while time.monotonic() < deadline:
                            remaining = max(0.01, deadline - time.monotonic())
                            try:
                                lines.append(self.queue.get(timeout=remaining))
                            except queue.Empty:
                                break

                        # Esvazia o restante da fila sem esperar (até 200 linhas no total)
                        while len(lines) < 200:
                            try:
                                lines.append(self.queue.get_nowait())
                            except queue.Empty:
                                break

                        if not lines:
                            continue

                        batch_data = "\n".join(lines)
                        try:
                            r = API_RETRY_SESSION.post(
                                f"{self.api_url}/api/jobs/{self.job_id}/log",
                                headers={"X-Worker-Token": self.token},
                                json={"log_line": batch_data},
                                timeout=(API_CONNECT_TIMEOUT, API_READ_TIMEOUT)
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
                        r = API_RETRY_SESSION.get(
                            f"{api_url}/api/jobs/{job['id']}",
                            headers={"X-Worker-Token": token},
                            timeout=(API_CONNECT_TIMEOUT, API_READ_TIMEOUT),
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
                    cancel_event.wait(float(os.getenv("WORKER_POLL_INTERVAL_SECONDS", "3")))

            poller_thread = threading.Thread(target=poll_status, daemon=True)
            poller_thread.start()

            orig_stdout = sys.stdout
            streamer = APILogStream(job["id"], orig_stdout, main_thread_id)
            sys.stdout = streamer

            # ── Inicializar documentação de evidências ─────────────────────────────
            documentation = None
            doc_row_context: dict[str, str] = {}
            try:
                import importlib.util as _ilu
                from pathlib import Path as _Path
                _project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
                if _project_dir and _project_dir not in sys.path:
                    sys.path.insert(0, _project_dir)
                from workflow.documentation import WorkflowDocumentation  # type: ignore
                _ticket_key = (
                    str(params.get("jira_key") or "").strip().upper()
                    or str(job.get("id", ""))[:8].upper()
                )
                _processo = str(params.get("processo") or "").strip()
                _subprocesso = str(params.get("subprocesso") or "").strip()
                workflow_parts = [part for part in (_processo, _subprocesso) if part]
                _workflow_name = " | ".join(workflow_parts) if workflow_parts else "sap_cockpit"
                doc_row_context = {
                    "ticket_key": _ticket_key,
                    "categoria_sap": _processo,
                    "request_number": str(params.get("request_number") or "").strip().upper(),
                    "xlsx_path": str(params.get("caminho_ficheiro") or "").strip(),
                    "ambiente": str(params.get("ambiente") or "").strip(),
                }
                documentation = WorkflowDocumentation.from_env(
                    base_dir=_Path(_project_dir) if _project_dir else _Path("."),
                    row_context=doc_row_context,
                    workflow_name=_workflow_name,
                )
            except Exception as _doc_init_exc:
                print(f"[DOC] Aviso: não foi possível inicializar documentação: {_doc_init_exc}")
            # ──────────────────────────────────────────────────────────────────────

            _cockpit_ok = True
            _cockpit_error = ""
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
                _cockpit_ok = False
                _cockpit_error = "Cancelado pelo utilizador."
                _force_terminate_worker()
            except Exception as _cockpit_exc:
                _cockpit_ok = False
                _cockpit_error = str(_cockpit_exc)
                raise
            finally:
                cancel_event.set()
                # ── Gerar documento de evidências ──────────────────────────────────────
                if documentation:
                    try:
                        _step_name = (
                            str(params.get("subprocesso") or params.get("processo") or "Execução SAP")
                        )
                        documentation.capture_step(
                            step_name=_step_name,
                            row_context=doc_row_context,
                            note="" if _cockpit_ok else f"Erro: {_cockpit_error}",
                            allow_live_capture=_cockpit_ok,
                        )
                        _doc_path = documentation.finalize(
                            row_context=doc_row_context,
                            success=_cockpit_ok,
                            error=_cockpit_error,
                        )
                        if _doc_path:
                            print(f"[DOC] Documento de evidências gerado: {_doc_path}")
                            log_lines.append(f"[DOC] Evidências: {_doc_path}")
                    except Exception as _doc_fin_exc:
                        print(f"[DOC] Aviso: falha ao gerar documento: {_doc_fin_exc}")
                # ──────────────────────────────────────────────────────────────────────
                sys.stdout = orig_stdout
                streamer.close()

            log_lines.append(log)
            return status or "Execucao concluida, mas STATUS veio vazio.", "\n".join(log_lines)


        raise SapExecutionError(f"Rotina desconhecida: {task}")

    except Exception:
        log_lines.append(traceback.format_exc())
        raise
