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
        return str(result[0] or ""), str(result[1] or ""), True

    if isinstance(result, dict):
        status = str(result.get("status") or result.get("STATUS") or "").strip()
        log = str(result.get("log") or result.get("log_text") or "")
        success = result.get("success", True)
        return status, log, success

    return str(result or ""), "", True


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
    
    session = None
    try:
        session = get_first_available_session()
    except Exception:
        pass

    try:
        lista = mod.listar_requests(
            system_name=sistema_desejado,
            max_rows="5000",
            include_requests=False,
            use_new_mode=True,
            minimize=True,
            close_after=True,
            session=session
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


# Whitelist de subprocessos com suporte a documentação funcional automática.
# Para adicionar suporte a um novo subprocesso basta adicionar o nome aqui.
SUPPORTED_FUNCTIONAL_DOC_PROCESSES: frozenset[str] = frozenset({
    "A. PFCG_CREATE.py",
})


def _run_authorization_open_cua(params: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()

    target_user = str(params.get("target_user") or "").strip().upper()
    target_system_key = str(params.get("target_system_key") or "").strip().upper()
    analysis_type = str(params.get("analysis_type") or "").strip().lower()
    cua_sap_key = str(params.get("cua_sap_key") or os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()

    if not target_user:
        raise SapExecutionError("Utilizador a analisar não foi informado.")

    if not target_system_key:
        raise SapExecutionError("Sistema alvo da análise não foi informado.")

    if analysis_type not in {"master_data", "authorizations"}:
        raise SapExecutionError("Tipo de análise inválido.")

    try:
        from sap_session import ensure_sap_access_from_env, session_info

        session = ensure_sap_access_from_env(
            key=cua_sap_key,
            timeout_s=40,
            load_env=True,
        )
        info = session_info(session)
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível abrir a sessão SAP CUA: {exc}") from exc

    expected_system = "SPA"
    expected_client = "001"

    if str(info.get("system_name") or "").strip().upper() != expected_system:
        raise SapExecutionError("A sessão aberta não corresponde ao sistema CUA esperado.")

    if str(info.get("client") or "").strip() != expected_client:
        raise SapExecutionError("A sessão aberta não corresponde ao cliente CUA esperado.")

    result = {
        "success": True,
        "execution_environment": "CUA",
        "system_name": info.get("system_name", ""),
        "client": info.get("client", ""),
        "target_user": target_user,
        "target_system_key": target_system_key,
        "analysis_type": analysis_type,
    }

    status = json.dumps(result, ensure_ascii=False)
    log = (
        "Sessão SAP CUA validada com sucesso.\n"
        f"Utilizador analisado: {target_user}\n"
        f"Sistema alvo: {target_system_key}\n"
        f"Tipo de análise: {analysis_type}\n"
        f"Sistema técnico: {info.get('system_name', '')}\n"
        f"Cliente técnico: {info.get('client', '')}"
    )

    return status, log


def validate_completed_analysis(result: dict[str, Any]) -> None:
    if not result:
        raise SapExecutionError("Resultado da análise está vazio.")

    if not result.get("success"):
        raise SapExecutionError(f"A análise falhou: {result.get('message')}")

    code = result.get("code")
    if code not in {"analysis_complete", "user_not_assigned_to_system", "user_not_found"}:
        raise SapExecutionError(f"Código de conclusão de análise inválido: {code}")

    if result.get("data_source_verified") is not True:
        raise SapExecutionError("A veracidade das fontes da análise não pôde ser confirmada.")

    if result.get("worker_feature_version") != "authorization-tables-v1":
        raise SapExecutionError("A sessão CUA foi aberta, mas a consulta das tabelas não foi executada pelo Worker (versão antiga detetada).")

    queries = result.get("queries") or []
    if not queries:
        raise SapExecutionError("A sessão CUA foi aberta, mas a consulta das tabelas não foi executada.")

    tables_executed = {q.get("table") for q in queries if q.get("executed") and q.get("filters_applied")}
    source = str(result.get("source") or "").strip().upper()
    dev_flow = source == "RFC_AGR_USERS" or "AGR_USERS" in tables_executed
    master_flow = source in {"RFC_USER_MASTER", "CUA_USER_MASTER"} or "USR02" in tables_executed

    if code == "analysis_complete":
        if master_flow:
            required = {"USR02", "USR21", "USR04", "AGR_USERS"}
        else:
            required = {"AGR_USERS", "AGR_TCODES"} if dev_flow else {"USZBVSYS", "USLA04", "USL04"}
        missing = required - tables_executed
        if missing:
            raise SapExecutionError(f"A consulta das tabelas obrigatórias está incompleta. Tabelas em falta: {', '.join(missing)}")

    elif code in {"user_not_assigned_to_system", "user_not_found"}:
        if master_flow:
            required_table = "USR02"
        else:
            required_table = "AGR_USERS" if dev_flow else "USZBVSYS"
        if required_table not in tables_executed:
            raise SapExecutionError(f"A tabela {required_table} não foi consultada para confirmar a associação ao sistema.")


def _authorization_uses_rfc(target_system_key: str) -> bool:
    system_key = str(target_system_key or "").strip().upper()
    if os.getenv("AUTHORIZATION_FORCE_RFC", "").strip().lower() in {"1", "true", "yes", "sim"}:
        return True
    return system_key.startswith("S4")


def _make_job_progress_logger(job_id: str):
    api_url = os.environ.get("SAP_API_BASE_URL", "").strip()
    token = os.environ.get("SAP_WORKER_TOKEN", "").strip()

    def _log(line: str) -> None:
        message = str(line or "").strip()
        if not message or not api_url or not token:
            return
        try:
            requests.post(
                f"{api_url}/api/jobs/{job_id}/log",
                headers={"X-Worker-Token": token},
                json={"log_line": message},
                timeout=5,
            )
        except Exception:
            pass

    return _log


def _run_authorization_analyze_user_rfc(params: dict[str, Any], progress_logger: Any | None = None) -> tuple[str, str]:
    _prepare_project_imports()

    target_user = str(params.get("target_user") or "").strip().upper()
    target_system_key = str(params.get("target_system_key") or "").strip().upper()
    analysis_type = str(params.get("analysis_type") or "").strip().lower()

    if not target_user:
        raise SapExecutionError("Utilizador a analisar não foi informado.")

    if not target_system_key:
        raise SapExecutionError("Sistema alvo da análise não foi informado.")

    if analysis_type not in {"master_data", "authorizations"}:
        raise SapExecutionError("Tipo de análise inválido.")

    try:
        if analysis_type == "master_data":
            from .user_master_data_analysis import analyze_user_master_data_rfc
        elif target_system_key.startswith(("S4D", "S4P")):
            from .authorization_rfc_dev_analysis import analyze_user_authorizations_rfc_dev
        else:
            from .authorization_rfc_analysis import analyze_user_authorizations_rfc
    except ImportError:
        if analysis_type == "master_data":
            from user_master_data_analysis import analyze_user_master_data_rfc
        elif target_system_key.startswith(("S4D", "S4P")):
            from authorization_rfc_dev_analysis import analyze_user_authorizations_rfc_dev
        else:
            from authorization_rfc_analysis import analyze_user_authorizations_rfc

    if analysis_type == "master_data":
        result = analyze_user_master_data_rfc(
            target_user=target_user,
            target_system_key=target_system_key,
            progress_logger=progress_logger,
            connection_params=params.get("rfc_connection") if isinstance(params.get("rfc_connection"), dict) else None,
        )
    elif target_system_key.startswith(("S4D", "S4P")):
        result = analyze_user_authorizations_rfc_dev(
            target_user=target_user,
            target_system_key=target_system_key,
            progress_logger=progress_logger,
            connection_params=params.get("rfc_connection") if isinstance(params.get("rfc_connection"), dict) else None,
        )
    else:
        result = analyze_user_authorizations_rfc(
            target_user=target_user,
            target_system_key=target_system_key,
            progress_logger=progress_logger,
            connection_params=params.get("rfc_connection") if isinstance(params.get("rfc_connection"), dict) else None,
        )

    validate_completed_analysis(result)

    log_lines = [
        "[AUTH RFC] Ligação RFC validada.",
    ]

    queries = result.get("queries", [])

    if analysis_type == "master_data":
        q_usr02 = next((q for q in queries if q["table"] == "USR02"), None)
        if q_usr02:
            log_lines.append("[AUTH RFC] Tabela USR02 informada.")
            log_lines.append("[AUTH RFC] Filtros USR02 aplicados.")
            log_lines.append("[AUTH RFC] Dados de validade e bloqueio lidos.")

        q_usr21 = next((q for q in queries if q["table"] == "USR21"), None)
        if q_usr21:
            log_lines.append("[AUTH RFC] Tabela USR21 informada.")
            log_lines.append("[AUTH RFC] Ligação do utilizador ao endereço lida.")

        q_usr04 = next((q for q in queries if q["table"] == "USR04"), None)
        if q_usr04:
            log_lines.append("[AUTH RFC] Tabela USR04 informada.")
            log_lines.append("[AUTH RFC] Perfis do utilizador lidos.")

        q_roles = next((q for q in queries if q["table"] == "AGR_USERS"), None)
        if q_roles:
            log_lines.append("[AUTH RFC] Tabela AGR_USERS informada.")
            log_lines.append("[AUTH RFC] Roles do utilizador lidas.")
    else:
        q_dev_users = next((q for q in queries if q["table"] == "AGR_USERS"), None)
        if q_dev_users:
            log_lines.append("[AUTH RFC] Tabela AGR_USERS informada.")
            log_lines.append("[AUTH RFC] Filtros AGR_USERS aplicados.")
            log_lines.append("[AUTH RFC] Roles lidas via AGR_USERS.")

        q_sys = next((q for q in queries if q["table"] == "USZBVSYS"), None)
        if q_sys:
            log_lines.append("[AUTH RFC] Tabela USZBVSYS informada.")
            log_lines.append("[AUTH RFC] Filtros USZBVSYS aplicados.")
            log_lines.append("[AUTH RFC] Resultado USZBVSYS lido.")

        if result.get("code") == "analysis_complete":
            q_dev_tcodes = next((q for q in queries if q["table"] == "AGR_TCODES"), None)
            if q_dev_tcodes:
                log_lines.append("[AUTH RFC] Tabela AGR_TCODES informada.")
                log_lines.append("[AUTH RFC] Filtros AGR_TCODES aplicados.")
                log_lines.append("[AUTH RFC] Funções AGR_TCODES lidas.")

            q_roles = next((q for q in queries if q["table"] == "USLA04"), None)
            if q_roles:
                log_lines.append("[AUTH RFC] Tabela USLA04 informada.")
                log_lines.append("[AUTH RFC] Filtros USLA04 aplicados.")
                log_lines.append("[AUTH RFC] Lista clássica USLA04 lida.")

            q_profs = next((q for q in queries if q["table"] == "USL04"), None)
            if q_profs:
                log_lines.append("[AUTH RFC] Tabela USL04 informada.")
                log_lines.append("[AUTH RFC] Filtros USL04 aplicados.")
                log_lines.append("[AUTH RFC] Resultado USL04 lido.")

    log_lines.append("[AUTH RFC] Análise concluída com sucesso.")

    safe_log = "\n".join(log_lines)
    return json.dumps(result, ensure_ascii=False), safe_log


def _run_authorization_analyze_user(params: dict[str, Any], progress_logger: Any | None = None) -> tuple[str, str]:
    _prepare_project_imports()

    target_user = str(params.get("target_user") or "").strip().upper()
    target_system_key = str(params.get("target_system_key") or "").strip().upper()
    analysis_type = str(params.get("analysis_type") or "").strip().lower()
    cua_sap_key = str(params.get("cua_sap_key") or os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()

    if not target_user:
        raise SapExecutionError("Utilizador a analisar não foi informado.")

    if not target_system_key:
        raise SapExecutionError("Sistema alvo da análise não foi informado.")

    if analysis_type not in {"master_data", "authorizations"}:
        raise SapExecutionError("Tipo de análise inválido.")

    try:
        from sap_session import ensure_sap_access_from_env, session_info

        cua_session = ensure_sap_access_from_env(
            key=cua_sap_key,
            timeout_s=40,
            load_env=True,
        )
        info = session_info(cua_session)
    except Exception as exc:
        raise SapExecutionError(f"Não foi possível abrir a sessão SAP CUA: {exc}") from exc

    expected_system = "SPA"
    expected_client = "001"

    if str(info.get("system_name") or "").strip().upper() != expected_system:
        raise SapExecutionError("A sessão aberta não corresponde ao sistema CUA esperado (deve ser SPA).")

    if str(info.get("client") or "").strip() != expected_client:
        raise SapExecutionError("A sessão aberta não corresponde ao cliente CUA esperado (deve ser 001).")

    import sys
    import os
    WORKER_DIR = os.path.dirname(os.path.abspath(__file__))
    if WORKER_DIR not in sys.path:
        sys.path.insert(0, WORKER_DIR)

    if analysis_type == "master_data":
        try:
            from .user_master_data_analysis import analyze_user_master_data
        except ImportError:
            from user_master_data_analysis import analyze_user_master_data

        result = analyze_user_master_data(
            session=cua_session,
            target_user=target_user,
            target_system_key=target_system_key,
            progress_logger=progress_logger,
        )
    else:
        try:
            from .authorization_table_analysis import analyze_user_authorizations
        except ImportError:
            from authorization_table_analysis import analyze_user_authorizations

        result = analyze_user_authorizations(
            session=cua_session,
            target_user=target_user,
            target_system_key=target_system_key,
            progress_logger=progress_logger,
        )

    validate_completed_analysis(result)

    log_lines = [
        "[AUTH] Sessão CUA validada: SPA/001."
    ]

    queries = result.get("queries", [])

    if analysis_type == "master_data":
        q_usr02 = next((q for q in queries if q["table"] == "USR02"), None)
        if q_usr02:
            log_lines.append("[AUTH] Tabela USR02 informada.")
            log_lines.append("[AUTH] Filtros USR02 aplicados.")
            log_lines.append("[AUTH] Dados de validade e bloqueio lidos.")

        q_usr21 = next((q for q in queries if q["table"] == "USR21"), None)
        if q_usr21:
            log_lines.append("[AUTH] Tabela USR21 informada.")
            log_lines.append("[AUTH] Ligação do utilizador ao endereço lida.")

        q_usr04 = next((q for q in queries if q["table"] == "USR04"), None)
        if q_usr04:
            log_lines.append("[AUTH] Tabela USR04 informada.")
            log_lines.append("[AUTH] Perfis do utilizador lidos.")

        q_roles = next((q for q in queries if q["table"] == "AGR_USERS"), None)
        if q_roles:
            log_lines.append("[AUTH] Tabela AGR_USERS informada.")
            log_lines.append("[AUTH] Roles do utilizador lidas.")
    else:
        q_sys = next((q for q in queries if q["table"] == "USZBVSYS"), None)
        if q_sys:
            log_lines.append("[AUTH] A abrir SE16.")
            log_lines.append("[AUTH] Tabela USZBVSYS informada.")
            log_lines.append("[AUTH] Filtros USZBVSYS aplicados.")
            log_lines.append("[AUTH] Resultado USZBVSYS lido.")

        if result.get("code") == "analysis_complete":
            q_roles = next((q for q in queries if q["table"] == "USLA04"), None)
            if q_roles:
                log_lines.append("[AUTH] Tabela USLA04 informada.")
                log_lines.append("[AUTH] Filtros USLA04 aplicados.")
                log_lines.append("[AUTH] Lista clássica USLA04 lida.")

            q_profs = next((q for q in queries if q["table"] == "USL04"), None)
            if q_profs:
                log_lines.append("[AUTH] Tabela USL04 informada.")
                log_lines.append("[AUTH] Filtros USL04 aplicados.")
                log_lines.append("[AUTH] Resultado USL04 lido.")

    log_lines.append("[AUTH] Análise concluída com sucesso.")

    safe_log = "\n".join(log_lines)

    return json.dumps(result, ensure_ascii=False), safe_log


def run_sap_task(job: dict[str, Any]) -> tuple[str, str]:
    _prepare_project_imports()
    module_name = os.getenv("SAP_COCKPIT_MODULE", "sap_script_web_cockpit_v2.sap_cockpit_web_ready").strip()
    try:
        cockpit = importlib.import_module(module_name)
        if hasattr(cockpit, "reload_project_env"):
            cockpit.reload_project_env(force_override=True)
        else:
            from dotenv import load_dotenv
            load_dotenv(override=True)
    except Exception:
        try:
            from sap_script_web_cockpit_v2.sap_cockpit_web_ready import reload_project_env
            reload_project_env(force_override=True)
        except Exception:
            from dotenv import load_dotenv
            load_dotenv(override=True)

    task = job["task"]
    params = job.get("params", {}) or {}
    log_lines: list[str] = [f"Job: {job['id']}", f"Task: {task}", f"Params: {params}"]
    try:
        if task == "authorization_analyze_user":
            os.environ["SAP_JOB_ID"] = str(job["id"])
            os.environ["SAP_API_BASE_URL"] = os.getenv("API_BASE_URL", "http://localhost:8000").rstrip("/")
            os.environ["SAP_WORKER_TOKEN"] = os.getenv("WORKER_TOKEN", "change-me")
            progress_logger = _make_job_progress_logger(str(job["id"]))
            target_user = str(params.get("target_user") or "").strip().upper()
            target_system_key = str(params.get("target_system_key") or "").strip().upper()
            analysis_type = str(params.get("analysis_type") or "").strip().lower()
            requested_execution_mode = str(params.get("execution_mode") or "").strip().upper()
            execution_mode = "RFC" if _authorization_uses_rfc(target_system_key) else "CUA"
            cua_sap_key = str(params.get("cua_sap_key") or os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()
            modo_label = f"modo={execution_mode}"
            if requested_execution_mode and requested_execution_mode != execution_mode:
                modo_label = f"modo_pedido={requested_execution_mode}, modo_execucao={execution_mode}"
            progress_logger(
                "[AUTH] Pedido recebido: "
                f"utilizador={target_user or '<vazio>'}, "
                f"sistema={target_system_key or '<vazio>'}, "
                f"tipo={analysis_type or '<vazio>'}, "
                f"{modo_label}, "
                f"cua_sap_key={cua_sap_key}."
            )
            progress_logger(
                f"[AUTH] A iniciar a análise no sistema {target_system_key}."
                if not _authorization_uses_rfc(target_system_key)
                else f"[AUTH RFC] A iniciar a ligação RFC ao sistema {target_system_key}."
            )
            if _authorization_uses_rfc(target_system_key):
                status, log = _run_authorization_analyze_user_rfc(params, progress_logger=progress_logger)
            else:
                status, log = _run_authorization_analyze_user(params, progress_logger=progress_logger)
            log_lines.append(log)
            return status, "\n".join(log_lines)

        if task == "authorization_open_cua":
            status, log = _run_authorization_open_cua(params)
            log_lines.append(log)
            return status, "\n".join(log_lines)

        if task == "sap_agent_analysis":
            status, log = _run_sap_agent_analysis(params)
            log_lines.append(log)
            return status, "\n".join(log_lines)

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

        if task == "sap_gui_chat_action":
            status, log = _run_sap_gui_chat_action(params)
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
                            timeout=15
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
                                    # Fechar a própria conexão da sessão para fechar a janela SAP correspondente
                                except Exception:
                                    pass
                                _force_terminate_worker()
                                break
                    except Exception as pe:
                        # Increment poller timeout counters
                        os.environ["CURRENT_ROLE_POLLER_TIMEOUT"] = str(int(os.environ.get("CURRENT_ROLE_POLLER_TIMEOUT", 0)) + 1)
                        os.environ["TOTAL_POLLER_TIMEOUT"] = str(int(os.environ.get("TOTAL_POLLER_TIMEOUT", 0)) + 1)
                        print(f"\n[TECHNICAL WARN] [DEBUG POLLER] Erro ao consultar estado do job: {pe}")
                        sys.stdout.flush()
                    cancel_event.wait(5.0)

            poller_thread = threading.Thread(target=poll_status, daemon=True)
            poller_thread.start()

            orig_stdout = sys.stdout
            streamer = APILogStream(job["id"], orig_stdout, main_thread_id)
            sys.stdout = streamer

            # ── Guardar nome_pasta no ambiente para processos que suportam DOC ──────
            # A documentação funcional é gerida internamente pelo processo (ex.: A. PFCG_CREATE.py).
            # Aqui apenas verificamos se o subprocesso está na whitelist e, se não estiver,
            # emitimos um log discreto para informar que a documentação foi ignorada.
            _nome_pasta = str(params.get("nome_pasta") or "").strip()
            _subprocesso_solicitado = str(params.get("subprocesso") or "").strip()
            if _nome_pasta and _subprocesso_solicitado and _subprocesso_solicitado not in SUPPORTED_FUNCTIONAL_DOC_PROCESSES:
                print(f"[DOC] Documentação funcional ignorada: subprocesso ainda não suportado ({_subprocesso_solicitado})")
            # ──────────────────────────────────────────────────────────────────────

            success_val = True
            try:
                status, log, success_val = _run_sap_cockpit(params)
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
                sys.stdout = orig_stdout
                streamer.close()

            log_lines.append(log)
            return status or "Execucao concluida, mas STATUS veio vazio.", "\n".join(log_lines), success_val


        raise SapExecutionError(f"Rotina desconhecida: {task}")

    except Exception:
        log_lines.append(traceback.format_exc())
        raise
