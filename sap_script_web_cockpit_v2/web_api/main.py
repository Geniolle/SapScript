import ast
import dataclasses
import importlib
import json
import os
from pathlib import Path
import sys
import requests

from uuid import uuid4
from typing import Any, Optional
from datetime import datetime
import time

import threading

WORKER_OFFLINE_AFTER_SECONDS = float(os.getenv("WORKER_OFFLINE_AFTER_SECONDS", "60.0"))
worker_last_seen: dict[str, float] = {}
worker_last_seen_lock = threading.Lock()
WORKER_HEARTBEAT_DEBUG = os.getenv("WORKER_HEARTBEAT_DEBUG", "false").strip().lower() in ("1", "true", "yes", "sim")
delayed_heartbeats_count = 0
delayed_heartbeats_lock = threading.Lock()

def update_worker_ping(worker_name: str) -> None:
    if not worker_name:
        return
    with worker_last_seen_lock:
        worker_last_seen[worker_name] = time.time()

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field

from web_api.store import append_job_log, cancel_job, claim_next_job, complete_job, create_job, get_job, init_db, list_jobs, archive_job, unarchive_job, delete_job, delete_all_jobs, update_job_params, save_jira_tickets_to_db, list_jira_tickets, update_jira_ticket_assignee, update_jira_ticket_type_db, update_jira_ticket_status_db, update_jira_ticket_supplier_db, log_auto_trigger_entry, list_auto_trigger_log, has_active_job_for_ticket, clear_auto_trigger_log, delete_auto_trigger_log_entry, get_latest_sap_agent_analysis, save_jira_ticket_batch_only, create_agent_rule, list_agent_rules, update_agent_rule, delete_agent_rule, get_agent_rules_for_ticket, get_transacao_by_processo, get_job_state
from web_api.jira_client import fetch_jira_tickets_from_api, assign_jira_ticket, update_jira_ticket_type, get_jira_issue_transitions, transition_jira_issue, update_jira_ticket_supplier, fetch_auto_trigger_tickets, download_ticket_attachments_to_dir, fetch_ticket_details, add_jira_comment, clean_excel_leading_spaces
from web_api.authorization_agent import get_analysis_types, get_execution_mode_for_system_key
import asyncio
import difflib
import unicodedata
from openpyxl import Workbook

def update_worker_ping_for_job(job_id: str) -> None:
    try:
        job = get_job(job_id)
        if job and job.get("worker_name"):
            update_worker_ping(job["worker_name"])
    except Exception:
        pass

WORKER_REQUIRED_TASKS = {
    "authorization_open_cua",
    "authorization_analyze_user",
    "authorization_hr_search",
    "sap_cockpit",
    "sap_agent_analysis",
    "sap_gui_chat_action",
    "open_transaction",
    "select_excel_file",
    "ping_status",
    "sap_search_requests",
}

def get_worker_runtime_status(worker_name: str | None = None) -> dict[str, Any]:
    import time
    now = time.time()
    is_online = False
    last_seen_seconds = None
    with worker_last_seen_lock:
        if worker_name:
            last_seen = worker_last_seen.get(worker_name)
            if last_seen is not None:
                last_seen_seconds = now - last_seen
                is_online = last_seen_seconds < WORKER_OFFLINE_AFTER_SECONDS
        else:
            if worker_last_seen:
                most_recent = max(worker_last_seen.values())
                last_seen_seconds = now - most_recent
                is_online = last_seen_seconds < WORKER_OFFLINE_AFTER_SECONDS

    if is_online:
        return {
            "online": True,
            "state": "online",
            "message": "Worker online.",
            "last_seen_seconds": round(last_seen_seconds, 1) if last_seen_seconds is not None else None
        }
    else:
        return {
            "online": False,
            "state": "offline",
            "message": "O Worker Windows está desligado.",
            "last_seen_seconds": round(last_seen_seconds, 1) if last_seen_seconds is not None else None
        }

def is_worker_online() -> bool:
    return get_worker_runtime_status()["online"]

def require_worker_online(task: str | None = None, message: str | None = None):
    if task and task not in WORKER_REQUIRED_TASKS:
        return
    status = get_worker_runtime_status()
    if not status["online"]:
        msg = message or "O Worker Windows está desligado. Ligue o worker antes de iniciar a ação."
        raise HTTPException(
            status_code=503,
            detail={
                "code": "worker_offline",
                "message": msg,
            },
        )


WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
SAP_SCRIPT_PROJECT_DIR = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
UPLOADS_DIR = Path(os.getenv("UPLOADS_DIR", "/uploads"))
UPLOADS_WINDOWS_DIR = os.getenv("UPLOADS_WINDOWS_DIR", "").strip()

# Auto-trigger configuration
AUTO_TRIGGER_INTERVAL_SECONDS = int(os.getenv("AUTO_TRIGGER_INTERVAL_SECONDS", "300"))
AUTO_TRIGGER_AMBIENTE = os.getenv("AUTO_TRIGGER_AMBIENTE", "PRD").strip().upper()
AUTO_TRIGGER_ENABLED = os.getenv("AUTO_TRIGGER_ENABLED", "true").strip().lower() in ("1", "true", "yes", "sim")

# Diretório de download de anexos JIRA
# No container Docker: /data/jira  (montado a partir de C:\Jira no host Windows)
JIRA_DOWNLOAD_DIR_CONTAINER = os.getenv("JIRA_DOWNLOAD_DIR_CONTAINER", "/data/jira").strip()
JIRA_DOWNLOAD_DIR_WINDOWS = os.getenv("JIRA_DOWNLOAD_DIR_WINDOWS", r"C:\Jira").strip()

# Mapeamento Categoria JIRA → parâmetros SAP
# Formato JSON: {"Categoria": {"processo": "...", "subprocesso": "...", "request_option": "N"}}
_DEFAULT_CATEGORY_MAP = json.dumps({
    "FI Extracto Cadeias de Pesquisa": {
        "processo": "Cadeias de Pesquisa",
        "subprocesso": "Criar Atribuir Cadeias.py",
        "request_option": "1",
        "ambiente": "DEV",
    }
})

def _load_category_map() -> dict[str, dict]:
    raw = os.getenv("AUTO_TRIGGER_CATEGORY_MAP", _DEFAULT_CATEGORY_MAP).strip()
    try:
        return json.loads(raw)
    except Exception as exc:
        print(f"[AUTO-TRIGGER] Erro ao carregar AUTO_TRIGGER_CATEGORY_MAP: {exc}")
        return json.loads(_DEFAULT_CATEGORY_MAP)

app = FastAPI(title="SAP Script Web")
app.mount("/static", StaticFiles(directory="web_api/static"), name="static")
templates = Jinja2Templates(directory="web_api/templates")


class CompleteJobRequest(BaseModel):
    state: str
    status: str
    log: str = ""


def _prepare_project_imports() -> None:
    if SAP_SCRIPT_PROJECT_DIR and SAP_SCRIPT_PROJECT_DIR not in sys.path:
        sys.path.insert(0, SAP_SCRIPT_PROJECT_DIR)


def _load_project_config():
    _prepare_project_imports()
    return importlib.import_module("app.config")


def get_available_environments() -> list[dict[str, str]]:
    try:
        config = _load_project_config()

        ambientes = getattr(config, "AMBIENTES", {})
        mapa_sistema = getattr(config, "MAPA_SISTEMA", {})
        clientes = getattr(config, "CLIENTES_POR_AMBIENTE", {})

    except Exception:
        ambientes = {
            "1": ("DEV", "DESENVOLVIMENTO (S4H)"),
            "2": ("QAD", "QUALIDADE (S4H)"),
            "3": ("PRD", "PRODUÇÃO (S4H)"),
            "4": ("CUA", "CUA (PRD)"),
        }

        mapa_sistema = {
            "DEV": "S4D",
            "QAD": "S4Q",
            "PRD": "S4P",
            "CUA": "SPA",
        }

        clientes = {
            "DEV": "100",
            "QAD": "100",
            "PRD": "100",
            "CUA": "001",
        }

    def sort_key(item):
        numero, valores = item
        try:
            return int(numero), valores[0]
        except Exception:
            return 9999, str(numero)

    result = []

    for _numero, valores in sorted(ambientes.items(), key=sort_key):
        codigo = str(valores[0]).strip().upper()
        nome = str(valores[1]).strip()

        sistema = str(mapa_sistema.get(codigo, "")).strip().upper()
        cliente = str(clientes.get(codigo, "")).strip()

        label = f"{codigo} - {nome}"

        result.append({
            "codigo": codigo,
            "nome": nome,
            "label": label,
        })

    return result


def _candidate_process_dirs() -> list[str]:
    candidatos: list[str] = []

    try:
        config = _load_project_config()
        processos_dir_config = str(getattr(config, "PROCESSOS_DIR", "") or "").strip()
        if processos_dir_config:
            candidatos.append(processos_dir_config)
    except Exception:
        pass

    if SAP_SCRIPT_PROJECT_DIR:
        candidatos.append(os.path.join(SAP_SCRIPT_PROJECT_DIR, "Processos"))
        candidatos.append(os.path.join(SAP_SCRIPT_PROJECT_DIR, "processos"))

    candidatos.append(os.path.abspath(os.path.join(os.getcwd(), "..", "Processos")))
    candidatos.append(os.path.abspath(os.path.join(os.getcwd(), "..", "processos")))

    result: list[str] = []
    vistos: set[str] = set()

    for caminho in candidatos:
        if not caminho:
            continue

        caminho_abs = os.path.abspath(caminho)

        if caminho_abs in vistos:
            continue

        vistos.add(caminho_abs)
        result.append(caminho_abs)

    return result


def _resolve_processes_dir() -> str | None:
    for caminho in _candidate_process_dirs():
        if os.path.isdir(caminho):
            return caminho
    return None


def _resolve_process_path(processo: str) -> str | None:
    processo = str(processo or "").strip()

    if not processo:
        return None

    if os.path.isabs(processo):
        return None

    processo_normalizado = os.path.normpath(processo)

    if processo_normalizado.startswith(".."):
        return None

    if processo_normalizado in (".", ""):
        return None

    processos_dir = _resolve_processes_dir()

    if not processos_dir:
        return None

    processos_dir_abs = os.path.abspath(processos_dir)
    caminho = os.path.abspath(os.path.join(processos_dir_abs, processo_normalizado))

    def _normalizar_chave(valor: str) -> str:
        texto = str(valor or "").strip()
        if not texto:
            return ""

        candidatos = [texto]
        for origem, destino in (("utf-8", "latin1"), ("utf-8", "cp1252")):
            try:
                candidatos.append(texto.encode(origem).decode(destino))
            except Exception:
                pass

        melhor = ""
        for candidato in candidatos:
            base = unicodedata.normalize("NFKD", candidato)
            base = "".join(ch for ch in base if not unicodedata.combining(ch))
            base = re.sub(r"[^0-9a-z]+", "", base.casefold())
            if len(base) > len(melhor):
                melhor = base
        return melhor

    if caminho != processos_dir_abs and not caminho.startswith(processos_dir_abs + os.sep):
        return None

    if not os.path.isdir(caminho):
        processo_key = _normalizar_chave(processo_normalizado)
        melhor_caminho = ""
        melhor_score = 0.0

        for nome in os.listdir(processos_dir_abs):
            if nome == "__pycache__":
                continue

            candidato_path = os.path.join(processos_dir_abs, nome)
            if not os.path.isdir(candidato_path):
                continue

            score = difflib.SequenceMatcher(None, processo_key, _normalizar_chave(nome)).ratio()
            if score > melhor_score:
                melhor_score = score
                melhor_caminho = candidato_path

        if melhor_caminho and melhor_score >= 0.75:
            return melhor_caminho

        return None

    return caminho


def get_available_processes() -> list[dict[str, str]]:
    processos_dir = _resolve_processes_dir()

    if not processos_dir:
        return []

    processos: list[dict[str, str]] = []
    # ###################################################################################
    # (1) PRIORIDADE VISUAL DOS PROCESSOS NO MENU
    # ###################################################################################
    prioridade_menu = {
        "UAT Simulação": 0,
    }

    def sort_key(nome: str) -> tuple[int, str]:
        return prioridade_menu.get(nome, 1000), nome.casefold()

    for nome in sorted(os.listdir(processos_dir), key=sort_key):
        if nome.startswith("~$"):
            continue

        if nome == "__pycache__":
            continue

        caminho = os.path.join(processos_dir, nome)

        if not os.path.isdir(caminho):
            continue

        processos.append({
            "nome": nome,
            "label": nome,
            "path": caminho,
        })

    return processos


def get_available_subprocesses(processo: str) -> list[dict[str, str]]:
    caminho_processo = _resolve_process_path(processo)

    if not caminho_processo:
        return []

    subprocessos: list[dict[str, str]] = []

    for nome in sorted(os.listdir(caminho_processo), key=str.casefold):
        if nome.startswith("~$"):
            continue

        if not nome.lower().endswith(".py"):
            continue

        caminho = os.path.join(caminho_processo, nome)

        if not os.path.isfile(caminho):
            continue

        if _extract_ast_var(caminho, "WEB_HIDDEN") is True:
            continue

        subprocessos.append({
            "nome": nome,
            "label": nome,
            "path": caminho,
        })

    return subprocessos


def _extract_ast_var(script_path: str, var_name: str):
    """
    Extrai o valor de uma variável de módulo de um ficheiro .py via AST,
    sem executar o código (evita side-effects como logging, SAP, etc.).
    Suporta apenas literais Python (listas, dicts, strings, bools, None).
    """
    try:
        with open(script_path, "r", encoding="utf-8") as f:
            source = f.read()
        tree = ast.parse(source, filename=script_path)
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id == var_name:
                        return ast.literal_eval(node.value)
    except Exception:
        pass
    return None


@app.get("/api/subprocess-web-params")
def api_subprocess_web_params(processo: str = "", subprocesso: str = "") -> dict[str, Any]:
    """
    Retorna WEB_PARAMS e WEB_CONFIG definidos num subprocess .py via análise AST.
    Usado pelo frontend para construir o popup dinamicamente por processo.
    """
    process_path = _resolve_process_path(processo)
    if not process_path:
        return {"params": None, "config": None}

    nome = str(subprocesso).strip()
    if not nome.lower().endswith(".py"):
        nome = f"{nome}.py"

    script_path = os.path.join(process_path, nome)
    if not os.path.isfile(script_path):
        return {"params": None, "config": None}

    return {
        "params": _extract_ast_var(script_path, "WEB_PARAMS"),
        "config": _extract_ast_var(script_path, "WEB_CONFIG"),
    }


def _safe_upload_filename(filename: str) -> str:
    """
    Gera um nome seguro para guardar ficheiros enviados pelo browser.
    Mantém apenas caracteres simples e prefixa com um ID único.
    """
    raw_name = Path(filename or "ficheiro").name.strip() or "ficheiro"
    safe_chars = []

    for char in raw_name:
        if char.isalnum() or char in {".", "-", "_", " "}:
            safe_chars.append(char)
        else:
            safe_chars.append("_")

    safe_name = "".join(safe_chars).strip(" .") or "ficheiro"
    return f"{uuid4().hex}_{safe_name}"


def _windows_upload_path(saved_name: str) -> str:
    """
    Converte o nome guardado no container para um caminho Windows acessível ao worker.
    """
    if UPLOADS_WINDOWS_DIR:
        return str(Path(UPLOADS_WINDOWS_DIR) / saved_name)

    return str(UPLOADS_DIR / saved_name)

def _fetch_all_sync_tickets() -> list[dict]:
    # 1. Fetch standard tickets (open tickets based on JIRA_SYNC_JQL)
    open_tickets = fetch_jira_tickets_from_api()
    
    # 2. Fetch resolved tickets for current year (2026 onwards)
    jql = os.getenv("JIRA_SYNC_JQL", "assignee = currentUser() AND statusCategory != Done")
    if "statusCategory != Done" in jql:
        resolved_jql = jql.replace("statusCategory != Done", "statusCategory = Done")
    else:
        resolved_jql = "(project = 'IT - Salsa Jeans' OR project = 'SAP - Desenvolvimento') AND statusCategory = Done"
    
    # Restrict to current year resolves for performance
    resolved_jql += ' AND resolved >= "2026-01-01"'
    
    try:
        resolved_tickets = fetch_jira_tickets_from_api(jql=resolved_jql)
    except Exception as exc:
        print(f"[JIRA SYNC] Erro ao buscar tickets resolvidos: {exc}")
        resolved_tickets = []
        
    combined = {}
    for t in open_tickets:
        if t.get("key"):
            combined[t["key"]] = t
    for t in resolved_tickets:
        if t.get("key"):
            combined[t["key"]] = t
            
    return list(combined.values())


async def sync_jira_tickets_loop() -> None:
    """
    Loop em segundo plano que roda a cada 60 segundos buscando os tickets JIRA.
    """
    while True:
        try:
            # Executa a busca HTTP em thread pool para evitar travar o event loop do FastAPI
            tickets = await asyncio.to_thread(_fetch_all_sync_tickets)
            # Guarda na BD local
            await asyncio.to_thread(save_jira_tickets_to_db, tickets)
        except Exception as exc:
            print(f"[JIRA SYNC LOOP ERROR]: {exc}")
        await asyncio.sleep(60)


async def historical_jira_sync() -> None:
    """
    Sincronização histórica executada em segundo plano.
    Busca tickets resolvidos de anos anteriores (antes de 2026) e salva em lotes na BD local.
    """
    print("[JIRA HISTORICAL SYNC] A verificar necessidade de sincronização histórica...")
    try:
        from web_api.store import get_connection, save_jira_ticket_batch_only
        with get_connection() as conn:
            row = conn.execute(
                "SELECT count(*) FROM jira_tickets WHERE resolved_at IS NOT NULL AND resolved_at != '' AND resolved_at < '2026-01-01'"
            ).fetchone()
            count = row[0] if row else 0

        if count >= 100:
            print(f"[JIRA HISTORICAL SYNC] Encontrados {count} tickets históricos na BD. Sincronização histórica ignorada.")
            return

        print("[JIRA HISTORICAL SYNC] A iniciar sincronização histórica (tickets resolvidos antes de 2026)...")
        historical_jql = '(project = "IT - Salsa Jeans" OR project = "SAP - Desenvolvimento") AND statusCategory = Done AND resolved < "2026-01-01"'

        def save_batch(batch):
            save_jira_ticket_batch_only(batch)
            print(f"[JIRA HISTORICAL SYNC] Gravados {len(batch)} tickets históricos na BD.")

        await asyncio.to_thread(fetch_jira_tickets_from_api, historical_jql, save_batch)
        print("[JIRA HISTORICAL SYNC] Sincronização histórica concluída com sucesso!")
    except Exception as exc:
        print(f"[JIRA HISTORICAL SYNC ERROR]: {exc}")


async def run_auto_trigger() -> dict[str, Any]:
    """
    Lógica central do auto-trigger.

    Para cada ticket elegível:
      1. Verifica se a categoria tem mapeamento SAP configurado
      2. Verifica anti-duplicação (ticket_key + updated_at)
      3. Descarrega o anexo XLSX do ticket JIRA para /data/jira/{key}/
      4. Cria job SAP com o processo/subprocesso/caminho_ficheiro corretos
    """
    result: dict[str, Any] = {
        "tickets_found": 0,
        "triggered": 0,
        "skipped": 0,
        "errors": 0,
        "entries": [],
    }

    # Check if worker is online
    if not is_worker_online():
        print("[AUTO-TRIGGER] Worker Windows offline. Skipping auto-trigger execution.")
        return result

    try:
        tickets = await asyncio.to_thread(fetch_auto_trigger_tickets)
    except Exception as exc:
        print(f"[AUTO-TRIGGER] Erro ao consultar JIRA: {exc}")
        return result

    result["tickets_found"] = len(tickets)
    category_map = _load_category_map()

    for ticket in tickets:
        key = ticket.get("key", "")
        summary = ticket.get("summary", "")
        updated_at = ticket.get("updated_at", "")
        categoria = ticket.get("process", "")  # IT SALSA - Categoria SAP

        entry: dict[str, Any] = {
            "key": key,
            "summary": summary,
            "categoria": categoria,
            "status": "",
            "job_id": None,
            "reason": "",
            "caminho_ficheiro": "",
        }

        try:
            # ----------------------------------------------------------------
            # 1. Verificar mapeamento de categoria
            # ----------------------------------------------------------------
            sap_config = category_map.get(categoria)
            if not sap_config:
                entry["status"] = "skipped"
                entry["reason"] = f"Sem mapeamento para categoria: '{categoria}'"
                result["skipped"] += 1
                await asyncio.to_thread(
                    log_auto_trigger_entry, key, summary, None, "skipped",
                    f"sem_mapeamento:{categoria}"
                )
                print(f"[AUTO-TRIGGER] {key}: sem mapeamento para '{categoria}'")
                result["entries"].append(entry)
                continue

            processo = sap_config.get("processo", "")
            subprocesso = sap_config.get("subprocesso", "")
            request_option = sap_config.get("request_option", "1")
            ambiente = sap_config.get("ambiente", AUTO_TRIGGER_AMBIENTE)

            # ----------------------------------------------------------------
            # 2. Anti-duplicação
            # ----------------------------------------------------------------
            already_active = await asyncio.to_thread(
                has_active_job_for_ticket, key, updated_at
            )
            if already_active:
                entry["status"] = "skipped"
                entry["reason"] = "Já existe job ativo para esta versão do ticket"
                result["skipped"] += 1
                await asyncio.to_thread(
                    log_auto_trigger_entry, key, summary, None, "skipped", updated_at
                )
                result["entries"].append(entry)
                continue

            # ----------------------------------------------------------------
            # 3. Download do anexo XLSX do ticket JIRA
            # ----------------------------------------------------------------
            xlsx_files = await asyncio.to_thread(
                download_ticket_attachments_to_dir,
                key,
                JIRA_DOWNLOAD_DIR_CONTAINER,
                JIRA_DOWNLOAD_DIR_WINDOWS,
                True,   # only_xlsx
                False,  # overwrite: False → se já existe, usa o existente
            )

            if not xlsx_files:
                entry["status"] = "skipped"
                entry["reason"] = "Sem ficheiro XLSX anexado ao ticket"
                result["skipped"] += 1
                await asyncio.to_thread(
                    log_auto_trigger_entry, key, summary, None, "skipped",
                    "sem_anexo_xlsx"
                )
                print(f"[AUTO-TRIGGER] {key}: sem anexo XLSX - job não criado.")
                result["entries"].append(entry)
                continue

            # Usa o ficheiro XLSX mais recente (primeiro da lista, já ordenada)
            caminho_ficheiro = xlsx_files[0]
            entry["caminho_ficheiro"] = caminho_ficheiro
            print(f"[AUTO-TRIGGER] {key}: ficheiro -> {caminho_ficheiro}")

            # ----------------------------------------------------------------
            # 4. Criar job SAP
            # ----------------------------------------------------------------
            job_params: dict[str, Any] = {
                "jira_key": key,
                "jira_summary": summary,
                "jira_updated_at": updated_at,
                "jira_categoria": categoria,
                "ambiente": ambiente,
                "processo": processo,
                "subprocesso": subprocesso,
                "request_option": request_option,
                "request_number": "",
                "request_desc": f"{key} | {summary}",
                "request_type": "1",
                "caminho_ficheiro": caminho_ficheiro,
                "transacao": "",
                "auto_triggered": True,
            }
            job = await asyncio.to_thread(create_job, "sap_cockpit", job_params)
            job_id = job["id"]

            entry["status"] = "triggered"
            entry["job_id"] = job_id
            entry["reason"] = updated_at
            result["triggered"] += 1

            await asyncio.to_thread(
                log_auto_trigger_entry, key, summary, job_id, "triggered", updated_at
            )
            print(
                f"[AUTO-TRIGGER] Job criado para {key} | processo={processo} | "
                f"subprocesso={subprocesso} | ficheiro={caminho_ficheiro} | job_id={job_id}"
            )

        except Exception as exc:
            entry["status"] = "error"
            entry["reason"] = str(exc)
            result["errors"] += 1
            await asyncio.to_thread(
                log_auto_trigger_entry, key, summary, None, "error", str(exc)
            )
            print(f"[AUTO-TRIGGER] Erro ao processar {key}: {exc}")

        result["entries"].append(entry)

    return result



async def auto_trigger_loop() -> None:
    """
    Loop em segundo plano que corre o auto-trigger a cada AUTO_TRIGGER_INTERVAL_SECONDS.
    """
    # Aguarda 30s no arranque para deixar o servidor estabilizar
    await asyncio.sleep(30)
    while True:
        try:
            result = await run_auto_trigger()
            print(
                f"[AUTO-TRIGGER LOOP] found={result['tickets_found']} "
                f"triggered={result['triggered']} skipped={result['skipped']} "
                f"errors={result['errors']}"
            )
        except Exception as exc:
            print(f"[AUTO-TRIGGER LOOP ERROR]: {exc}")
        await asyncio.sleep(AUTO_TRIGGER_INTERVAL_SECONDS)



@app.on_event("startup")
def startup() -> None:
    init_db()
    asyncio.create_task(sync_jira_tickets_loop())
    asyncio.create_task(historical_jira_sync())
    if AUTO_TRIGGER_ENABLED:
        asyncio.create_task(auto_trigger_loop())


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    context = {
        "request": request,
        "ambientes": get_available_environments(),
        "processos": get_available_processes(),
        "jira_base": os.getenv("JIRA_DADOS_COMP_HASH", "https://salsajeans.atlassian.net").strip(),
    }
    response = templates.TemplateResponse(
        request=request,
        name="index.html",
        context=context,
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/api/jira/tickets")
def api_list_jira_tickets(limit: int = 50, exclude_closed: bool = True) -> dict[str, Any]:
    try:
        return {"tickets": list_jira_tickets(limit=limit, exclude_closed=exclude_closed)}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/jira/sync")
async def api_force_jira_sync() -> dict[str, Any]:
    try:
        tickets = await asyncio.to_thread(_fetch_all_sync_tickets)
        await asyncio.to_thread(save_jira_tickets_to_db, tickets)
        # Dispara sincronização histórica se necessário
        asyncio.create_task(historical_jira_sync())
        return {"status": "success", "synced_count": len(tickets)}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao sincronizar com JIRA: {str(exc)}")


class AssigneeRequest(BaseModel):
    assignee: str


@app.post("/api/jira/tickets/{ticket_key}/assign")
async def api_assign_jira_ticket(ticket_key: str, payload: AssigneeRequest) -> dict[str, Any]:
    try:
        # Update locally in SQLite first
        await asyncio.to_thread(update_jira_ticket_assignee, ticket_key, payload.assignee)
        
        # Try to sync with Jira API
        success = await asyncio.to_thread(assign_jira_ticket, ticket_key, payload.assignee)
        
        return {"status": "success", "jira_updated": success}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class TicketTypeRequest(BaseModel):
    ticket_type: str


@app.post("/api/jira/tickets/{ticket_key}/type")
async def api_update_jira_ticket_type(ticket_key: str, payload: TicketTypeRequest) -> dict[str, Any]:
    try:
        # Update locally in SQLite first
        await asyncio.to_thread(update_jira_ticket_type_db, ticket_key, payload.ticket_type)
        
        # Try to sync with Jira API
        success = await asyncio.to_thread(update_jira_ticket_type, ticket_key, payload.ticket_type)
        
        return {"status": "success", "jira_updated": success}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/api/jira/tickets/{ticket_key}/transitions")
async def api_get_jira_transitions(ticket_key: str) -> dict[str, Any]:
    try:
        transitions = await asyncio.to_thread(get_jira_issue_transitions, ticket_key)
        return {"transitions": transitions}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class TransitionRequest(BaseModel):
    transition_id: str
    status_name: str


@app.post("/api/jira/tickets/{ticket_key}/transition")
async def api_transition_jira_ticket(ticket_key: str, payload: TransitionRequest) -> dict[str, Any]:
    try:
        # Try to transition with Jira API
        success = await asyncio.to_thread(transition_jira_issue, ticket_key, payload.transition_id)
        
        # If success, update locally in SQLite
        if success:
            await asyncio.to_thread(update_jira_ticket_status_db, ticket_key, payload.status_name)
        
        return {"status": "success", "jira_updated": success}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class SupplierRequest(BaseModel):
    supplier: str


@app.post("/api/jira/tickets/{ticket_key}/supplier")
async def api_update_jira_ticket_supplier(ticket_key: str, payload: SupplierRequest) -> dict[str, Any]:
    try:
        # Update locally in SQLite first
        await asyncio.to_thread(update_jira_ticket_supplier_db, ticket_key, payload.supplier)
        
        # Try to sync with Jira API
        success = await asyncio.to_thread(update_jira_ticket_supplier, ticket_key, payload.supplier)
        
        return {"status": "success", "jira_updated": success}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class CommentRequest(BaseModel):
    comment: str


@app.post("/api/jira/tickets/{ticket_key}/comment")
async def api_add_jira_comment(ticket_key: str, payload: CommentRequest) -> dict[str, Any]:
    """Adiciona um comentário 'Reply to customer' ao ticket JIRA."""
    try:
        if not payload.comment or not payload.comment.strip():
            raise HTTPException(status_code=400, detail="O comentário não pode estar vazio.")
        success = await asyncio.to_thread(add_jira_comment, ticket_key, payload.comment.strip())
        return {"status": "success", "jira_updated": success}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/api/jira/tickets/{ticket_key}/details")
async def api_get_ticket_details(ticket_key: str) -> dict[str, Any]:
    """Retorna o sumário, descrição e comentários de um ticket JIRA."""
    try:
        details = await asyncio.to_thread(fetch_ticket_details, ticket_key)
        signal_preview: dict[str, Any] = {}
        context_matches: list[dict[str, Any]] = []

        try:
            from sap_agent.extractor import extract_signal
            from sap_agent.models import TicketContext

            preview_ticket = TicketContext(
                key=str(ticket_key or "").strip().upper(),
                summary=str(details.get("summary") or ""),
                description=str(details.get("description") or ""),
                comments=list(details.get("comments") or []),
            )
            signal = extract_signal(preview_ticket)
            context_matches = get_agent_rules_for_ticket(
                processo=str(details.get("categoria_sap") or ""),
            )

            # ------------------------------------------------------------------
            # 1ª prioridade: IT SALSA - Categoria SAP → coluna Processo nas
            # Definições → Transação SAP (preenchimento automático principal).
            # Quando o utilizador clica em "Analisar", a primeira pesquisa é:
            # - Ler o valor do campo "IT SALSA - Categoria SAP" do ticket
            # - Encontrar a regra nas Definições onde Processo == esse valor
            # - Preencher automaticamente o campo Transação com transacao_sap
            # ------------------------------------------------------------------
            categoria_sap_val = str(details.get("categoria_sap") or "").strip()
            if categoria_sap_val and not signal.transaction:
                transacao_por_processo = get_transacao_by_processo(categoria_sap_val)
                if transacao_por_processo:
                    signal.transaction = transacao_por_processo

            # ------------------------------------------------------------------
            # 2ª prioridade: campo+valor (context_matches) — fallback se o
            # lookup por Processo não encontrou transação.
            # ------------------------------------------------------------------
            first_rule_with_transaction = next(
                (rule for rule in context_matches if rule.get("transacao_sap")),
                None,
            )
            if first_rule_with_transaction and not signal.transaction:
                signal.transaction = first_rule_with_transaction["transacao_sap"]

            signal_preview = dataclasses.asdict(signal)
        except Exception as preview_exc:
            print(
                f"[SAP AGENT PREVIEW] Falha ao extrair sinais do ticket {ticket_key}: {preview_exc}"
            )

        return {
            "status": "success",
            **details,
            "signal_preview": signal_preview,
            "context_matches": context_matches,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


# ---------------------------------------------------------------------------
# Autorizações SAP endpoints
# ---------------------------------------------------------------------------

def _find_authorization_env_file() -> Path | None:
    """Localiza o ficheiro .env na ordem de prioridade definida."""
    candidates: list[Path] = []

    explicit_path = os.getenv("SAP_AUTH_ENV_FILE", "").strip()
    if explicit_path:
        candidates.append(Path(explicit_path))

    project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if project_dir:
        candidates.append(Path(project_dir) / ".env")

    candidates.append(Path("/sap-script/.env"))
    candidates.append(Path.cwd() / ".env")

    current_file = Path(__file__).resolve()
    for parent in current_file.parents:
        candidates.append(parent / ".env")

    checked: set[str] = set()

    for candidate in candidates:
        try:
            normalized = str(candidate.resolve())
        except Exception:
            normalized = str(candidate)

        if normalized in checked:
            continue

        checked.add(normalized)

        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue

    return None

def _parse_env_file(path: Path) -> dict[str, str]:
    """Lê e parseia o .env de forma segura e em tempo de execução."""
    resultado = {}
    try:
        with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
            for line in f:
                line_str = line.strip()
                if not line_str or line_str.startswith("#"):
                    continue
                if "=" not in line_str:
                    continue
                parts = line_str.split("=", 1)
                key = parts[0].strip()
                val = parts[1].strip()
                
                # Remover aspas simples ou duplas se existirem
                if len(val) >= 2 and ((val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'"))):
                    val = val[1:-1]
                resultado[key] = val
    except Exception:
        pass
    return resultado


def _authorization_env_alias(system_key: str) -> str:
    cleaned = str(system_key or "").strip().upper()
    if cleaned.startswith("S4D"):
        return "DEV"
    if cleaned.startswith("S4Q"):
        return "QAD"
    if cleaned.startswith("S4P"):
        return "PRD"
    if cleaned.startswith("SPA"):
        return "CUA"
    return ""


def _authorization_system_short_name(system_key: str) -> str:
    cleaned = str(system_key or "").strip().upper()
    if "CLNT" in cleaned:
        return cleaned.split("CLNT", 1)[0]
    return cleaned


def _first_env_value(env_data: dict[str, str], *names: str) -> str:
    for name in names:
        if not name:
            continue
        value = str(env_data.get(name, "")).strip()
        if value:
            return value
    return ""


def _authorization_uat_create_document_defaults(env_data: dict[str, str]) -> dict[str, str]:
    today_yyyymmdd = datetime.now().strftime("%Y%m%d")

    return {
        "system_key": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_SYSTEM_KEY"),
        "company_code": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_COMPANY_CODE"),
        "vendor": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_VENDOR"),
        "customer": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CUSTOMER"),
        "customer_company_code": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CUSTOMER_COMPANY_CODE"),
        "customer_gl_account": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CUSTOMER_GL_ACCOUNT"),
        "customer_doc_type": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CUSTOMER_DOC_TYPE"),
        "vendor_payment_method": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_VENDOR_PAYMENT_METHOD"),
        "customer_payment_method": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CUSTOMER_PAYMENT_METHOD"),
        "gl_account": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_GL_ACCOUNT"),
        "amount": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_AMOUNT"),
        "currency": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_CURRENCY"),
        "document_date": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_DOCUMENT_DATE") or today_yyyymmdd,
        "posting_date": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_POSTING_DATE") or today_yyyymmdd,
        "payment_method": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_PAYMENT_METHOD"),
        "doc_type": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_DOC_TYPE"),
        "reference": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_REFERENCE"),
        "header_text": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_HEADER_TEXT"),
        "item_text": _first_env_value(env_data, "AUTHORIZATION_UAT_CREATE_DOCUMENT_ITEM_TEXT"),
    }


def _build_authorization_rfc_connection(env_data: dict[str, str], target_system_key: str) -> dict[str, str]:
    system_key = str(target_system_key or "").strip().upper()
    system_alias = _authorization_env_alias(system_key)
    system_name = system_key.split("CLNT", 1)[0]
    client = system_key.split("CLNT", 1)[1] if "CLNT" in system_key else ""

    connection = {
        "ashost": _first_env_value(
            env_data,
            f"SAP_ASHOST_{system_key}",
            f"SAP_ASHOST_{system_name}",
            f"SAP_ASHOST_{system_alias}" if system_alias else "",
            "SAP_ASHOST",
        ),
        "sysnr": _first_env_value(
            env_data,
            f"SAP_SYSNR_{system_key}",
            f"SAP_SYSNR_{system_name}",
            f"SAP_SYSNR_{system_alias}" if system_alias else "",
            "SAP_SYSNR",
        ) or "00",
        "client": client,
        "user": _first_env_value(
            env_data,
            f"SAP_USER_{system_key}",
            f"SAP_USER_{system_name}",
            f"SAP_USER_{system_alias}" if system_alias else "",
            "SAP_USER",
        ),
        "passwd": _first_env_value(
            env_data,
            f"SAP_PASSWORD_{system_key}",
            f"SAP_PASSWORD_{system_name}",
            f"SAP_PASSWORD_{system_alias}" if system_alias else "",
            "SAP_PASSWORD",
        ),
        "lang": _first_env_value(
            env_data,
            f"SAP_RFC_LANGUAGE_{system_key}",
            f"SAP_RFC_LANGUAGE_{system_name}",
            f"SAP_RFC_LANGUAGE_{system_alias}" if system_alias else "",
            f"SAP_LANGUAGE_{system_key}",
            f"SAP_LANGUAGE_{system_name}",
            f"SAP_LANGUAGE_{system_alias}" if system_alias else "",
            "SAP_RFC_LANGUAGE",
            "SAP_LANGUAGE",
        ) or "PT",
    }
    return connection

@app.get("/api/authorizations/context")
@app.get("/api/authorizations/config")
def api_get_authorizations_context() -> dict[str, Any]:
    """Retorna o contexto de autorizações atualizado relendo o .env do disco."""
    import re
    print("[AUTH CONFIG] Pedido recebido.")
    
    env_path = _find_authorization_env_file()
    
    # Logs de diagnóstico seguros
    has_explicit = "sim" if os.getenv("SAP_AUTH_ENV_FILE") else "não"
    print(f"[AUTH CONFIG] SAP_AUTH_ENV_FILE configurado: {has_explicit}")
    print(f"[AUTH CONFIG] Caminho pesquisado: {env_path or '/sap-script/.env'}")
    print(f"[AUTH CONFIG] Ficheiro encontrado: {'sim' if env_path else 'não'}")

    if not env_path:
        print("[AUTH CONFIG] Sistemas encontrados: 0.")
        print("[AUTH CONFIG] Resposta concluída.")
        return {
            "success": False,
            "user_sap": "",
            "systems": [],
            "env_found": False,
            "message": "Ficheiro .env não encontrado no contentor.",
            "create_document_defaults": _authorization_uat_create_document_defaults({}),
            "diagnostic": {
                "configured_path": os.getenv("SAP_AUTH_ENV_FILE", "/sap-script/.env"),
                "project_directory_configured": bool(os.getenv("SAP_SCRIPT_PROJECT_DIR"))
            }
        }
        
    env_data = _parse_env_file(env_path)
    user_sap = env_data.get("SAP_USER", "").strip()
    
    systems = []
    regex = re.compile(r"^(?P<system>[A-Z0-9_]+)CLNT(?P<client>[0-9]+)$")
    
    for key, val in env_data.items():
        if key.startswith("SAP_PASSWORD_"):
            sufixo = key[len("SAP_PASSWORD_"):]
            match = regex.match(sufixo)
            if match:
                sys_name = match.group("system")
                clnt_num = match.group("client")
                
                # Procurar connection_name opcional
                conn_key = f"SAP_CONNECTION_{sufixo}"
                conn_name = env_data.get(conn_key, "").strip()
                
                # Montar label premium
                if conn_name:
                    label = f"{sys_name} — Cliente {clnt_num} — {conn_name}"
                else:
                    label = f"{sys_name} — Cliente {clnt_num}"
                    
                systems.append({
                    "key": sufixo,
                    "system": sys_name,
                    "client": clnt_num,
                    "connection_name": conn_name,
                    "label": label,
                    "execution_mode": get_execution_mode_for_system_key(sufixo)
                })
                
    # Ordenar de forma previsível e remover duplicados
    systems = sorted(systems, key=lambda x: x["key"])
    
    message = ""
    if not user_sap:
        message = "A variável SAP_USER não está configurada no .env."
    elif not systems:
        message = "Nenhum sistema SAP foi encontrado no .env."
        
    print(f"[AUTH CONFIG] Sistemas encontrados: {len(systems)}.")
    print("[AUTH CONFIG] Resposta concluída.")
    return {
        "success": True,
        "user_sap": user_sap,
        "systems": systems,
        "analysis_types": get_analysis_types(),
        "env_found": True,
        "create_document_defaults": _authorization_uat_create_document_defaults(env_data),
        "message": message
    }


@app.get("/api/authorization/greeting")
def api_authorization_greeting() -> dict[str, Any]:
    """Gera uma saudação inicial dinâmica e abrangente para o Assistente SAP via Gemini (com fallback)."""
    api_key = os.getenv("GEMINI_API_KEY")
    if api_key:
        try:
            import requests
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
            prompt = (
                "És o Assistente Virtual SAP S/4HANA no Cockpit Web. "
                "Gera uma frase curta, elegante e acolhedora de saudação inicial em Português de Portugal (1 frase). "
                "Pergunta ao utilizador de forma aberta como o podes ajudar nos seus processos SAP hoje. "
                "NÃO restrinjas a saudação a autorizações ou perfis. Mantém a pergunta completamente aberta para qualquer processo SAP."
            )
            payload = {
                "contents": [{"role": "user", "parts": [{"text": prompt}]}]
            }
            res = requests.post(url, json=payload, timeout=6)
            if res.status_code == 200:
                data = res.json()
                candidates = data.get("candidates", [])
                if candidates:
                    text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "").strip()
                    if text:
                        return {"messages": [text]}
        except Exception as e:
            print(f"[AUTH GREETING] Aviso ao gerar saudação dinâmica via Gemini: {e}")

    import random
    fallbacks = [
        ["Olá! Como posso ajudar nos seus processos SAP hoje?"],
        ["Olá! Em que posso apoiar a gestão dos seus processos SAP?"],
        ["Boas-vindas! De que forma posso apoiar as suas operações e processos no SAP hoje?"]
    ]
    return {"messages": random.choice(fallbacks)}


class AuthorizationStartRequest(BaseModel):
    target_user: str
    target_system_key: str
    analysis_type: str
    subsystem_filter: Optional[str] = None


class AuthorizationRemoveRequest(BaseModel):
    target_user: str
    target_system_key: str
    roles: list[Any] = Field(default_factory=list)
    opcao_processamento: str = "sistema_user"


class HrSearchRequest(BaseModel):
    query: str
    target_system_key: Optional[str] = "S4PCLNT100"
    max_results: Optional[int] = 10


@app.post("/api/authorizations/hr-search")
async def api_search_hr_user_data(payload: HrSearchRequest) -> dict[str, Any]:
    query_text = (payload.query or "").strip()
    if not query_text:
        raise HTTPException(status_code=400, detail="Termo de pesquisa de RH não informado.")

    system_key = (payload.target_system_key or "S4PCLNT100").strip().upper()
    max_res = payload.max_results or 10

    # 1. Tentar executar localmente (se o pyrfc estiver instalado no ambiente)
    loop = asyncio.get_event_loop()

    def _try_local_search():
        try:
            try:
                from worker.hr_data_analysis import search_hr_user_data_rfc
            except ImportError:
                from hr_data_analysis import search_hr_user_data_rfc

            res = search_hr_user_data_rfc(query=query_text, target_system_key=system_key, max_results=max_res)
            return res
        except Exception as exc:
            return {"_failed_local": True, "error": str(exc)}

    local_res = await loop.run_in_executor(None, _try_local_search)
    if isinstance(local_res, dict) and not local_res.get("_failed_local") and local_res.get("success"):
        return local_res

    # 2. Se falhar localmente (ex: pyrfc indisponível dentro do contentor Linux Docker), delegar ao Worker Windows
    require_worker_online(
        "authorization_hr_search",
        message="O Worker Windows está desligado. Ligue o worker para consultar a tabela de RH via RFC."
    )

    job_params = {
        "query": query_text,
        "target_system_key": system_key,
        "max_results": max_res,
    }

    try:
        job = await asyncio.to_thread(create_job, "authorization_hr_search", job_params)
        job_id = job["id"]

        start_time = time.time()
        while time.time() - start_time < 15.0:
            await asyncio.sleep(0.5)
            j_state = await asyncio.to_thread(get_job, job_id)
            if not j_state:
                continue

            st = str(j_state.get("state") or "").lower()
            if st in ("succeeded", "succeeded_with_warnings", "failed"):
                raw_log = str(j_state.get("log") or "")
                for line in reversed(raw_log.splitlines()):
                    line_clean = line.strip()
                    if line_clean.startswith("{") and line_clean.endswith("}"):
                        try:
                            parsed = json.loads(line_clean)
                            if isinstance(parsed, dict) and "success" in parsed:
                                return parsed
                        except Exception:
                            pass

                if st == "failed":
                    return {
                        "success": False,
                        "message": "A pesquisa no RH falhou no Worker Windows.",
                        "data": [],
                        "query": query_text,
                        "total": 0,
                    }

        return {
            "success": False,
            "message": "Tempo limite excedido a aguardar resposta do Worker Windows.",
            "data": [],
            "query": query_text,
            "total": 0,
        }

    except Exception as exc:
        return {
            "success": False,
            "message": f"Erro na pesquisa de RH: {exc}",
            "data": [],
            "query": query_text,
            "total": 0,
        }


def is_any_worker_active() -> bool:
    import time
    now = time.time()
    with worker_last_seen_lock:
        for w_name, last_seen in worker_last_seen.items():
            if now - last_seen <= WORKER_OFFLINE_AFTER_SECONDS:
                return True
    return False


@app.post("/api/authorizations/start")
async def api_start_authorization_analysis(payload: AuthorizationStartRequest) -> dict[str, Any]:
    import re
    cua_sap_key = os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001").strip().upper()
    target_user = payload.target_user.strip().upper()
    requested_system = payload.target_system_key.strip().upper()
    requested_subsystem = (payload.subsystem_filter or "").strip().upper()

    execution_mode = get_execution_mode_for_system_key(requested_system)

    # SEPARAR TOTALMENTE O SISTEMA DE LOGON DO MANDANTE DE PESQUISA (SUBSYSTEM)
    if execution_mode == "CUA" or requested_system == cua_sap_key:
        target_system_key = cua_sap_key
        if requested_subsystem and requested_subsystem != cua_sap_key:
            subsystem_filter = requested_subsystem
        elif requested_system and requested_system != cua_sap_key:
            subsystem_filter = requested_system
        else:
            subsystem_filter = "S4DCLNT100"
    else:
        target_system_key = requested_system
        subsystem_filter = requested_subsystem or requested_system

    analysis_type = payload.analysis_type.strip().lower()

    if not target_user:
        raise HTTPException(status_code=400, detail="Utilizador a analisar não foi informado.")
    
    if len(target_user) > 40:
        raise HTTPException(status_code=400, detail="Utilizador a analisar excede o limite de 40 caracteres.")

    if not re.match(r"^[A-Z0-9.\-_]+$", target_user):
        raise HTTPException(status_code=400, detail="Utilizador contém caracteres inválidos.")

    if not target_system_key:
        raise HTTPException(status_code=400, detail="Sistema alvo da análise não foi informado.")

    if not re.match(r"^[A-Z0-9_]+CLNT[0-9]+$", target_system_key):
        raise HTTPException(status_code=400, detail="Formato de sistema lógico inválido.")

    # 1. Validar se o sistema existe no contexto do .env
    env_path = _find_authorization_env_file()
    if not env_path:
        raise HTTPException(status_code=400, detail="Ficheiro .env não encontrado no contentor.")
    
    env_data = _parse_env_file(env_path)
    if f"SAP_PASSWORD_{target_system_key}" not in env_data:
        raise HTTPException(status_code=400, detail=f"Sistema alvo '{target_system_key}' não está configurado no .env.")

    # 2. Validar tipo de análise
    from web_api.authorization_agent import validate_analysis_selection
    if not validate_analysis_selection(analysis_type) or analysis_type not in {"master_data", "authorizations"}:
        raise HTTPException(status_code=400, detail="Tipo de análise inválido.")

    execution_mode = get_execution_mode_for_system_key(target_system_key)

    # 3. Validar configuração da ligação SAP consoante o modo
    cua_sap_key = os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001").strip().upper()
    if execution_mode == "RFC":
        rfc_connection = _build_authorization_rfc_connection(env_data, target_system_key)
        if not rfc_connection["ashost"] or not rfc_connection["sysnr"] or not rfc_connection["user"] or not rfc_connection["passwd"]:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"Configuração RFC para {target_system_key} incompleta. "
                    "É necessário configurar o host, o número do sistema, o utilizador e a password no .env."
                )
            )
    else:
        if f"SAP_PASSWORD_{cua_sap_key}" not in env_data and f"SAP_CONNECTION_{cua_sap_key}" not in env_data:
            raise HTTPException(
                status_code=400,
                detail=f"Configuração do CUA ({cua_sap_key}) não foi encontrada no .env."
            )

    # 4. Validar se há pelo menos um worker disponível
    require_worker_online(
        "authorization_analyze_user",
        message="O Worker Windows está desligado. Ligue o worker antes de iniciar a análise."
    )

    params = {
        "target_user": target_user,
        "target_system_key": target_system_key,
        "subsystem_filter": subsystem_filter,
        "analysis_type": analysis_type,
        "cua_sap_key": cua_sap_key,
        "execution_mode": execution_mode,
        "rfc_connection": _build_authorization_rfc_connection(env_data, target_system_key) if execution_mode == "RFC" else None,
    }

    try:
        job = await asyncio.to_thread(create_job, "authorization_analyze_user", params)
        return {
            "job_id": job["id"],
            "state": job["state"],
            "message": "Análise de autorizações iniciada."
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/authorizations/remove")
async def api_remove_authorization_roles(payload: AuthorizationRemoveRequest) -> dict[str, Any]:
    import re

    target_user = payload.target_user.strip().upper()
    target_system_key = payload.target_system_key.strip().upper()
    opcao_processamento = str(payload.opcao_processamento or "sistema_user").strip().lower() or "sistema_user"
    cua_sap_key = os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001").strip().upper()

    if not target_user:
        raise HTTPException(status_code=400, detail="Utilizador a remover não foi informado.")

    if len(target_user) > 40:
        raise HTTPException(status_code=400, detail="Utilizador a remover excede o limite de 40 caracteres.")

    if not re.match(r"^[A-Z0-9.\-_]+$", target_user):
        raise HTTPException(status_code=400, detail="Utilizador contém caracteres inválidos.")

    if not target_system_key:
        raise HTTPException(status_code=400, detail="Sistema alvo não foi informado.")

    if not re.match(r"^[A-Z0-9_]+CLNT[0-9]+$", target_system_key):
        raise HTTPException(status_code=400, detail="Formato de sistema lógico inválido.")

    roles: list[str] = []
    for item in payload.roles or []:
        if isinstance(item, str):
            raw_role = item
        elif isinstance(item, dict):
            raw_role = item.get("role") or item.get("function") or item.get("agr_name") or item.get("AGR_NAME") or item.get("name") or ""
        else:
            raw_role = ""

        role = str(raw_role or "").strip().upper()
        if role:
            roles.append(role)

    roles = list(dict.fromkeys(roles))
    if not roles:
        raise HTTPException(status_code=400, detail="Nenhuma função válida foi informada para remoção.")

    if opcao_processamento not in {"sistema_user", "sistema"}:
        raise HTTPException(status_code=400, detail="Opção de processamento inválida. Valores permitidos: sistema_user, sistema.")

    env_path = _find_authorization_env_file()
    if not env_path:
        raise HTTPException(status_code=400, detail="Ficheiro .env não encontrado no contentor.")

    env_data = _parse_env_file(env_path)
    if f"SAP_PASSWORD_{cua_sap_key}" not in env_data and f"SAP_CONNECTION_{cua_sap_key}" not in env_data:
        raise HTTPException(status_code=400, detail=f"Configuração do CUA ({cua_sap_key}) não foi encontrada no .env.")

    ambiente = "CUA"

    require_worker_online(
        "sap_cockpit",
        message="O Worker Windows está desligado. Ligue o worker antes de iniciar a remoção de funções."
    )

    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

    system_short = _authorization_system_short_name(target_system_key)
    safe_name = _safe_upload_filename(f"CUA_REMOVE_{target_user}_{target_system_key}.xlsx")
    container_path = UPLOADS_DIR / safe_name

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CUA_REMOVE"
    sheet.append(["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP"])

    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    for idx, role in enumerate(roles, start=1):
        sheet.append([idx, target_user, target_system_key, role, "", "", now])

    workbook.save(container_path)

    job_params: dict[str, Any] = {
        "ambiente": ambiente,
        "processo": "Funções PFCG",
        "subprocesso": "J. CUA_REMOVE.py",
        "request_option": "4",
        "request_number": "",
        "request_desc": f"{target_user} | {target_system_key} | {len(roles)} funções a remover",
        "request_type": "1",
        "caminho_ficheiro": _windows_upload_path(safe_name),
        "transacao": "",
        "opcao_processamento": opcao_processamento,
        "target_user": target_user,
        "target_system_key": target_system_key,
        "cua_sap_key": cua_sap_key,
        "roles": roles,
        "source": "authorization_follow_up",
    }

    try:
        job = await asyncio.to_thread(create_job, "sap_cockpit", job_params)
        return {
            "success": True,
            "job_id": job["id"],
            "state": job["state"],
            "message": f"Job CUA_REMOVE criado com {len(roles)} funções.",
            "caminho_ficheiro": _windows_upload_path(safe_name),
            "ambiente": ambiente,
            "system_short": system_short,
            "roles_count": len(roles),
            "cua_sap_key": cua_sap_key,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


class AuthorizationEndDateRequest(BaseModel):
    target_user: str
    target_system_key: str
    roles: list[Any]
    valid_to: str = ""

@app.post("/api/authorizations/enddate")
async def api_authorizations_enddate(payload: AuthorizationEndDateRequest) -> dict[str, Any]:
    target_user = str(payload.target_user or "").strip().upper()
    target_system_key = str(payload.target_system_key or "").strip().upper()

    if not target_user:
        raise HTTPException(status_code=400, detail="Utilizador alvo não foi informado.")

    if not target_system_key:
        raise HTTPException(status_code=400, detail="Sistema alvo não foi informado.")

    roles: list[str] = []
    for item in payload.roles or []:
        if isinstance(item, str):
            raw_role = item
        elif isinstance(item, dict):
            raw_role = item.get("role") or item.get("function") or item.get("agr_name") or item.get("AGR_NAME") or item.get("name") or ""
        else:
            raw_role = ""

        role = str(raw_role or "").strip().upper()
        if role:
            roles.append(role)

    roles = list(dict.fromkeys(roles))
    if not roles:
        raise HTTPException(status_code=400, detail="Nenhuma função válida foi informada para alteração de validade.")

    cua_sap_key = os.getenv("CUA_SAP_KEY", "SPACLNT001").strip().upper()
    env_path = _find_authorization_env_file()
    if not env_path:
        raise HTTPException(status_code=400, detail="Ficheiro .env não encontrado no contentor.")

    env_data = _parse_env_file(env_path)
    if f"SAP_PASSWORD_{cua_sap_key}" not in env_data and f"SAP_CONNECTION_{cua_sap_key}" not in env_data:
        raise HTTPException(status_code=400, detail=f"Configuração do CUA ({cua_sap_key}) não foi encontrada no .env.")

    require_worker_online(
        "sap_cockpit",
        message="O Worker Windows está desligado. Ligue o worker antes de iniciar a alteração de validade."
    )

    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    system_short = _authorization_system_short_name(target_system_key)
    safe_name = _safe_upload_filename(f"CUA_ENDDATE_{target_user}_{target_system_key}.xlsx")
    container_path = UPLOADS_DIR / safe_name

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CUA_ENDDATE"
    
    headers = ["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP"]
    if payload.valid_to:
        headers.insert(4, "UPDATE_TO_DAT")
    sheet.append(headers)

    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    for idx, role in enumerate(roles, start=1):
        if payload.valid_to:
            sheet.append([idx, target_user, target_system_key, role, payload.valid_to, "", "", now])
        else:
            sheet.append([idx, target_user, target_system_key, role, "", "", now])

    workbook.save(container_path)

    job_params: dict[str, Any] = {
        "ambiente": "CUA",
        "processo": "Funções PFCG",
        "subprocesso": "I. CUA_ENDDATE.py",
        "request_option": "4",
        "request_number": "",
        "request_desc": f"{target_user} | {target_system_key} | {len(roles)} funções a alterar validade",
        "request_type": "1",
        "caminho_ficheiro": _windows_upload_path(safe_name),
        "transacao": "",
        "target_user": target_user,
        "target_system_key": target_system_key,
        "cua_sap_key": cua_sap_key,
        "roles": roles,
        "source": "authorization_follow_up",
    }

    try:
        job = await asyncio.to_thread(create_job, "sap_cockpit", job_params)
        return {
            "success": True,
            "job_id": job["id"],
            "state": job["state"],
            "message": f"Job CUA_ENDDATE criado com {len(roles)} funções.",
            "caminho_ficheiro": _windows_upload_path(safe_name),
            "ambiente": "CUA",
            "system_short": system_short,
            "roles_count": len(roles),
            "cua_sap_key": cua_sap_key,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


# ---------------------------------------------------------------------------
# Auto-Trigger SAP endpoints
# ---------------------------------------------------------------------------

@app.get("/api/jira/auto-trigger/config")
def api_auto_trigger_config() -> dict[str, Any]:
    """Retorna a configuração atual do auto-trigger."""
    return {
        "enabled": AUTO_TRIGGER_ENABLED,
        "interval_seconds": AUTO_TRIGGER_INTERVAL_SECONDS,
        "ambiente": AUTO_TRIGGER_AMBIENTE,
        "assignee": os.getenv("JIRA_AUTO_TRIGGER_ASSIGNEE", "Clayton Lopes"),
        "status_filter": os.getenv("JIRA_AUTO_TRIGGER_STATUS", "In Review"),
        "supplier_filter": os.getenv("JIRA_AUTO_TRIGGER_SUPPLIER", "Evolutive"),
        "processo": os.getenv("AUTO_TRIGGER_PROCESSO", ""),
        "subprocesso": os.getenv("AUTO_TRIGGER_SUBPROCESSO", ""),
    }


@app.get("/api/jira/auto-trigger/preview")
async def api_auto_trigger_preview() -> dict[str, Any]:
    """
    Retorna os tickets JIRA elegíveis para auto-trigger sem criar jobs.
    Útil para validar os critérios antes de executar.
    """
    try:
        tickets = await asyncio.to_thread(fetch_auto_trigger_tickets)
        enriched = []
        for t in tickets:
            key = t.get("key", "")
            updated_at = t.get("updated_at", "")
            already = await asyncio.to_thread(has_active_job_for_ticket, key, updated_at)
            enriched.append({**t, "already_active": already})
        return {
            "tickets_found": len(tickets),
            "tickets": enriched,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/jira/auto-trigger/run")
async def api_auto_trigger_run() -> dict[str, Any]:
    """
    Executa o auto-trigger manualmente: consulta JIRA e cria jobs SAP
    para todos os tickets elegíveis (com proteção anti-duplicação).
    """
    try:
        result = await run_auto_trigger()
        return {"status": "success", **result}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class ForceRunRequest(BaseModel):
    ticket_key: str


@app.post("/api/jira/auto-trigger/force-run")
async def api_auto_trigger_force_run(payload: ForceRunRequest) -> dict[str, Any]:
    """
    Força a execução do auto-trigger para um único ticket específico,
    ignorando o status do ticket, mas validando os outros critérios.
    """
    require_worker_online("sap_cockpit")
    ticket_key = payload.ticket_key.strip().upper()
    if not ticket_key:
        raise HTTPException(status_code=400, detail="Chave do ticket não fornecida.")

    try:
        from web_api.jira_client import fetch_single_ticket_for_trigger
        ticket = await asyncio.to_thread(fetch_single_ticket_for_trigger, ticket_key)
        if not ticket:
            raise HTTPException(status_code=404, detail=f"Ticket {ticket_key} não encontrado no JIRA.")

        category_map = _load_category_map()
        categoria = ticket.get("process", "")  # IT SALSA - Categoria SAP

        # 1. Validar mapeamento de categoria
        sap_config = category_map.get(categoria)
        if not sap_config:
            error_msg = f"Sem mapeamento configurado para a categoria: '{categoria}'"
            await asyncio.to_thread(
                log_auto_trigger_entry, ticket_key, ticket.get("summary", ""), None, "error", error_msg
            )
            raise HTTPException(status_code=400, detail=error_msg)

        processo = sap_config.get("processo", "")
        subprocesso = sap_config.get("subprocesso", "")
        request_option = sap_config.get("request_option", "1")
        ambiente = sap_config.get("ambiente", AUTO_TRIGGER_AMBIENTE)

        # 2. Download do anexo XLSX
        xlsx_files = await asyncio.to_thread(
            download_ticket_attachments_to_dir,
            ticket_key,
            JIRA_DOWNLOAD_DIR_CONTAINER,
            JIRA_DOWNLOAD_DIR_WINDOWS,
            True,   # only_xlsx
            False,  # overwrite: False (se já existe, usa o existente)
        )

        if not xlsx_files:
            error_msg = "Sem ficheiro XLSX anexado ao ticket."
            await asyncio.to_thread(
                log_auto_trigger_entry, ticket_key, ticket.get("summary", ""), None, "error", error_msg
            )
            raise HTTPException(status_code=400, detail=error_msg)

        caminho_ficheiro = xlsx_files[0]

        # 3. Criar job SAP
        job_params: dict[str, Any] = {
            "jira_key": ticket_key,
            "jira_summary": ticket.get("summary", ""),
            "jira_updated_at": ticket.get("updated_at", ""),
            "jira_categoria": categoria,
            "ambiente": ambiente,
            "processo": processo,
            "subprocesso": subprocesso,
            "request_option": request_option,
            "request_number": "",
            "request_desc": f"{ticket_key} | {ticket.get('summary', '')}",
            "request_type": "1",
            "caminho_ficheiro": caminho_ficheiro,
            "transacao": "",
            "auto_triggered": True,
        }

        job = await asyncio.to_thread(create_job, "sap_cockpit", job_params)
        job_id = job["id"]

        await asyncio.to_thread(
            log_auto_trigger_entry, ticket_key, ticket.get("summary", ""), job_id, "triggered", f"Execução manual forçada ({ticket.get('updated_at', '')})"
        )

        return {
            "status": "success",
            "message": f"Job SAP #{job_id[:8]} criado com sucesso para o ticket {ticket_key}.",
            "job_id": job_id,
            "processo": processo,
            "subprocesso": subprocesso,
            "caminho_ficheiro": caminho_ficheiro,
        }

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao forçar execução: {str(exc)}")


@app.get("/api/jira/auto-trigger/log")
def api_auto_trigger_log(limit: int = 50) -> dict[str, Any]:
    """Retorna o histórico de execuções do auto-trigger."""
    try:
        return {"log": list_auto_trigger_log(limit=limit)}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.delete("/api/jira/auto-trigger/log")
def api_clear_auto_trigger_log() -> dict[str, Any]:
    """Limpa todo o histórico de execuções do auto-trigger."""
    try:
        clear_auto_trigger_log()
        return {"status": "success", "message": "Histórico do auto-trigger limpo."}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.delete("/api/jira/auto-trigger/log/{entry_id}")
def api_delete_auto_trigger_log_entry(entry_id: str) -> dict[str, Any]:
    """Elimina uma entrada específica do histórico do auto-trigger."""
    try:
        delete_auto_trigger_log_entry(entry_id)
        return {"status": "success", "message": f"Entrada {entry_id} removida."}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/sap-agent/analyze/{ticket_key}")
def api_sap_agent_analyze(ticket_key: str) -> dict[str, Any]:
    """Cria um job técnico para o worker Windows executar a análise do Agente SAP no ticket indicado."""
    require_worker_online("sap_agent_analysis")
    try:
        job = create_job("sap_agent_analysis", {"ticket_key": ticket_key})
        return {"job_id": job["id"], "state": job["state"]}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


class SapAgentChatRequest(BaseModel):
    ticket_key: str
    message: str
    history: list[dict[str, str]] = []
    company_code: str = ""
    sap_query_enabled: bool = True


class SapQueryRequest(BaseModel):
    object_type: str  # 'internal_order', 'po', 'fi_doc', 'wbs', 'asset'
    object_number: str
    company_code: str = ""


@app.post("/api/sap-agent/chat")
def api_sap_agent_chat(request: SapAgentChatRequest) -> dict[str, Any]:
    """Conversação interativa com o Gemini com base no contexto do ticket e nos sinais SAP extraídos."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=400,
            detail="GEMINI_API_KEY não configurada no ficheiro .env. Por favor, adicione a chave e reinicie o cockpit.",
        )

    # 1. Obter detalhes do ticket no JIRA (inclui anexos e textos extraídos)
    ticket_info = fetch_ticket_details(request.ticket_key)
    summary = ticket_info.get("summary") or "Sem sumário"
    description = ticket_info.get("description") or "Sem descrição"
    comments_list = ticket_info.get("comments") or []
    comments_text = "\n".join(comments_list) if comments_list else "Sem comentários"

    # Contexto de anexos para o prompt
    attachments_raw: list[dict] = ticket_info.get("attachments") or []
    attachment_texts: list[str] = ticket_info.get("attachment_texts") or []

    _ATTACH_TEXT_MAX = 3000  # chars por anexo enviados ao modelo

    def _build_attachments_prompt(
        raw: list[dict], texts: list[str]
    ) -> str:
        if not raw:
            return "Nenhum anexo encontrado neste ticket."
        names = [str(a.get("filename") or "") for a in raw if a.get("filename")]
        lines = [f"Anexos ({len(names)} ficheiro(s)): " + ", ".join(names)]
        if not texts:
            lines.append("Sem texto extraído dos anexos.")
        else:
            lines.append("Conteúdo extraído dos anexos:")
            for block in texts:
                if len(block) > _ATTACH_TEXT_MAX:
                    block = block[:_ATTACH_TEXT_MAX] + f"\n[... conteúdo truncado a {_ATTACH_TEXT_MAX} caracteres]"
                lines.append(block)
        return "\n".join(lines)

    attachments_context = _build_attachments_prompt(attachments_raw, attachment_texts)

    # 2. Obter análise do Agente SAP da base de dados
    analysis_job = get_latest_sap_agent_analysis(request.ticket_key)

    signals_str = "Sem sinais identificados."
    evidences_str = "Sem evidências recolhidas."
    probable_cause = "Sem causa provável diagnosticada."
    proposed_solution = "Sem solução proposta."
    tests_str = "Sem testes sugeridos."

    if analysis_job and analysis_job.get("status"):
        try:
            report = json.loads(analysis_job["status"])
            sig = report.get("signal") or {}
            sig_fields = []
            if sig.get("transaction"): sig_fields.append(f"- Transação: {sig['transaction']}")
            if sig.get("program"): sig_fields.append(f"- Programa/Classe: {sig['program']}")
            if sig.get("message_id"): sig_fields.append(f"- Mensagem SAP: {sig['message_id']} {sig.get('message_number') or ''}")
            if sig.get("company_code"): sig_fields.append(f"- Empresa: {sig['company_code']}")
            if sig.get("document_number"): sig_fields.append(f"- Documento: {sig['document_number']}")
            if sig.get("fiscal_year"): sig_fields.append(f"- Exercício: {sig['fiscal_year']}")
            if sig.get("job_name"): sig_fields.append(f"- Job: {sig['job_name']}")
            if sig.get("user"): sig_fields.append(f"- Utilizador: {sig['user']}")
            if sig_fields:
                signals_str = "\n".join(sig_fields)

            evs = report.get("evidences") or []
            ev_list = []
            for e in evs:
                status_icon = "🟢" if e.get("status") == "ok" else ("🟡" if e.get("status") == "warning" else "🔴")
                ev_list.append(f"- {status_icon} {e.get('name')}: {e.get('details')}")
            if ev_list:
                evidences_str = "\n".join(ev_list)

            probable_cause = report.get("probable_cause") or probable_cause
            proposed_solution = report.get("proposed_solution") or proposed_solution

            tests = report.get("tests_to_execute") or []
            if tests:
                tests_str = "\n".join(f"- {t}" for t in tests)
        except Exception as e:
            print(f"[CHAT ERROR] Erro ao decodificar status do job de análise: {e}")

    # 3. Formular prompt do sistema
    system_prompt = f"""Você é o Assistente Especialista em SAP da Evolutive. Você está inserido no cockpit web para ajudar o Clayton a analisar e resolver um erro específico no ticket JIRA {request.ticket_key}.

Abaixo está o contexto do ticket JIRA:
- Chave: {request.ticket_key}
- Sumário: {summary}
- Descrição:
{description}
- Comentários:
{comments_text}
- Anexos:
{attachments_context}

Abaixo estão as evidências recolhidas pelo Agente SAP (no worker Windows local):
- Sinais Identificados:
{signals_str}
- Evidências recolhidas em SAP:
{evidences_str}
- Possível Causa diagnosticada:
{probable_cause}
- Prévia de Solução:
{proposed_solution}
- Testes sugeridos:
{tests_str}

O utilizador Clayton Lopes (consultor SAP) está a conversar contigo para explorar este ticket, sugerir novas soluções ou analisar erros adicionais. Responde de forma profissional, direta e técnica. Dá recomendações de tabelas SAP, transações (SM30, SM37, SE16N, etc.) e à análise funcional e técnica. Responde no mesmo idioma do utilizador (português).

Tens acesso a uma ferramenta especial: `sap_gui_action`. Quando o utilizador pedir para "abrir", "entrar", "pesquisar" ou "analisar" algo no SAP, usa esta ferramenta para executar a ação directamente no SAP GUI da máquina Windows.
Ações disponíveis:
- se16n_query: Pesquisar numa tabela SAP (EKKO, AUFK, BKPF, EKPO, etc.)
- open_transaction: Abrir qualquer transação SAP
- read_sbar: Ler o status bar da sessão SAP actual
"""

    # 3.5 Detetar intenção de consulta SAP na mensagem do utilizador
    sap_data_context = ""
    sap_query_badge = False
    if request.sap_query_enabled:
        try:
            import sys as _sys
            _project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
            if _project_dir and _project_dir not in _sys.path:
                _sys.path.insert(0, _project_dir)
            from sap_agent.sap_chat_tools import detect_sap_intent, query_sap_object
            obj_type, obj_number = detect_sap_intent(request.message)
            if obj_type and obj_number:
                sap_result = query_sap_object(
                    obj_type,
                    obj_number,
                    company_code=request.company_code or None,
                )
                if sap_result.data_blocks:
                    header = (
                        f"\n\n**📊 Dados reais lidos do SAP (objeto: {sap_result.object_type} — {sap_result.object_number}):**"
                        if sap_result.is_real_data
                        else f"\n\n**📌 Orientação de consulta SAP (objeto: {sap_result.object_type} — {sap_result.object_number}):**"
                    )
                    sap_data_context = header + "\n" + "\n\n".join(sap_result.data_blocks)
                    sap_query_badge = sap_result.is_real_data
        except Exception as _sap_exc:
            print(f"[CHAT SAP QUERY] Aviso ao tentar consultar SAP: {_sap_exc}")

    # 3.6 Actualizar prompt do sistema com os dados SAP detetados
    if sap_data_context:
        system_prompt += sap_data_context
        system_prompt += "\n\nCom base nos dados reais acima lidos do SAP, responde à mensagem do utilizador de forma técnica e precisa."

    # 4. Formular histórico para a chamada da API do Gemini
    contents = []
    for h in request.history:
        role = h.get("role")
        text = h.get("text")
        if role and text:
            contents.append({
                "role": "user" if role == "user" else "model",
                "parts": [{"text": text}]
            })

    # Adicionar mensagem atual do utilizador
    contents.append({
        "role": "user",
        "parts": [{"text": request.message}]
    })

    # 5. Definir as ferramentas SAP GUI para o Gemini (Function Calling)
    sap_gui_tools = [
        {
            "functionDeclarations": [
                {
                    "name": "sap_gui_action",
                    "description": (
                        "Executa uma ação directamente no SAP GUI aberto na máquina Windows. "
                        "Usa para pesquisar tabelas (SE16N), abrir transações, ler status bar."
                    ),
                    "parameters": {
                        "type": "OBJECT",
                        "properties": {
                            "action": {
                                "type": "STRING",
                                "enum": ["se16n_query", "open_transaction", "read_sbar"],
                                "description": "Ação a executar no SAP GUI."
                            },
                            "table": {
                                "type": "STRING",
                                "description": "Nome da tabela SAP (para se16n_query). Ex: EKKO, AUFK, BKPF."
                            },
                            "filters": {
                                "type": "ARRAY",
                                "items": {
                                    "type": "OBJECT",
                                    "properties": {
                                        "field": {"type": "STRING", "description": "Nome do campo SAP"},
                                        "value": {"type": "STRING", "description": "Valor do filtro"}
                                    }
                                },
                                "description": "Filtros a aplicar na pesquisa. Ex: [{\"field\": \"EBELN\", \"value\": \"4500123456\"}]"
                            },
                            "fields": {
                                "type": "ARRAY",
                                "items": {"type": "STRING"},
                                "description": "Campos a mostrar no resultado. Vazio = todos."
                            },
                            "transaction": {
                                "type": "STRING",
                                "description": "Código da transação SAP (para open_transaction). Ex: SE16N, KO03, ME23N."
                            },
                            "max_rows": {
                                "type": "INTEGER",
                                "description": "Número máximo de linhas a retornar (por defeito: 20)."
                            },
                            "description": {
                                "type": "STRING",
                                "description": "Descrição legível da ação para mostrar no chat."
                            }
                        },
                        "required": ["action"]
                    }
                }
            ]
        }
    ]

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": contents,
        "systemInstruction": {
            "parts": [{"text": system_prompt}]
        },
        "tools": sap_gui_tools,
    }

    # Retry com backoff exponencial para erros 503/429 (Gemini sobrecarregado)
    MAX_RETRIES = 3
    RETRY_DELAYS = [2, 4, 8]  # segundos
    response = None

    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=45)

            # Se 503 ou 429, tentar novamente após backoff
            if response.status_code in (503, 429) and attempt < MAX_RETRIES - 1:
                import time as _time
                wait = RETRY_DELAYS[attempt]
                print(f"[GEMINI] Erro {response.status_code} na tentativa {attempt + 1}/{MAX_RETRIES}. A aguardar {wait}s...")
                _time.sleep(wait)
                continue

            response.raise_for_status()
            res_data = response.json()

            candidates = res_data.get("candidates", [])
            if not candidates:
                return {"reply": "Não foi possível obter uma resposta válida do assistente."}

            candidate = candidates[0]
            content = candidate.get("content", {})
            parts = content.get("parts", [])

            # Verificar se o Gemini retornou uma function call (SAP GUI action)
            for part in parts:
                fc = part.get("functionCall")
                if fc and fc.get("name") == "sap_gui_action":
                    fc_args = fc.get("args", {})
                    action_desc = fc_args.get("description") or _build_sap_action_description(fc_args)

                    # Criar job no worker Windows para executar a ação SAP GUI
                    try:
                        require_worker_online("sap_gui_chat_action")
                        job = create_job("sap_gui_chat_action", {
                            **fc_args,
                            "ticket_key": request.ticket_key,
                            "sap_key": "S4PCLNT100",
                        })
                        return {
                            "reply": f"⚙️ A executar no SAP GUI: **{action_desc}**\n\nAguarda enquanto o worker Windows acede ao SAP...",
                            "waiting_sap": True,
                            "job_id": job["id"],
                            "sap_action": fc_args,
                        }
                    except Exception as job_exc:
                        error_msg = str(job_exc)
                        if isinstance(job_exc, HTTPException) and isinstance(job_exc.detail, dict):
                            error_msg = job_exc.detail.get("message", error_msg)
                        return {
                            "reply": f"❌ Não foi possível criar job SAP: {error_msg}\n\nAcesso manual: {action_desc}"
                        }

            # Resposta de texto normal
            if parts:
                reply = parts[0].get("text", "")
                return {"reply": reply}

            return {"reply": "Não foi possível obter uma resposta válida do assistente."}

        except Exception as e:
            if attempt < MAX_RETRIES - 1 and response is not None and response.status_code in (503, 429):
                continue
            detail_msg = str(e)
            if response is not None:
                try:
                    detail_msg = f"{response.status_code} - {response.text}"
                except Exception:
                    pass
            raise HTTPException(
                status_code=500,
                detail=f"Erro ao comunicar com a API do Gemini: {detail_msg}"
            )

    # Esgotadas as tentativas
    detail_msg = ""
    if response is not None:
        try:
            detail_msg = f"{response.status_code} - {response.text}"
        except Exception:
            pass
    raise HTTPException(
        status_code=503,
        detail=f"A API do Gemini está temporariamente indisponível (503). Por favor, tente novamente em alguns segundos. Detalhes: {detail_msg}"
    )


def _build_sap_action_description(fc_args: dict) -> str:
    """Gera descrição legível para uma sap_gui_action."""
    action = fc_args.get("action", "")
    if action == "se16n_query":
        table = fc_args.get("table", "")
        filters = fc_args.get("filters") or []
        filter_str = ", ".join(f"{f.get('field')}={f.get('value')}" for f in filters if f.get("field"))
        return f"SE16N → Tabela {table}" + (f" | Filtros: {filter_str}" if filter_str else "")
    elif action == "open_transaction":
        return f"Abrir transação {fc_args.get('transaction', '')}"
    elif action == "read_sbar":
        return "Ler status bar SAP"
    return str(fc_args)


@app.get("/api/sap-agent/chat-job/{job_id}")
def api_sap_agent_chat_job(job_id: str) -> dict[str, Any]:
    """Polling endpoint: retorna o estado e resultado de um job SAP GUI iniciado pelo chat."""
    try:
        job = get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")

        state = job.get("state", "pending")
        status_raw = job.get("status") or ""

        # Tentar desserializar o resultado JSON do worker
        sap_result = None
        result_text = ""
        rows: list = []
        if state == "succeeded" and status_raw:
            try:
                sap_result = json.loads(status_raw)
                result_text = sap_result.get("result_text", "")
                rows = sap_result.get("rows", [])
            except Exception:
                result_text = status_raw

        return {
            "job_id": job_id,
            "state": state,
            "result_text": result_text,
            "rows": rows,
            "error": sap_result.get("error") if sap_result else None,
            "success": sap_result.get("success", False) if sap_result else False,
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/sap-agent/sap-query")
def api_sap_agent_sap_query(req: SapQueryRequest) -> dict[str, Any]:
    """Consulta direta ao SAP via RFC. Usa as credenciais do .env.
    Retorna os dados brutos do SAP para debug ou uso direto no frontend."""
    try:
        import sys as _sys
        _project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
        if _project_dir and _project_dir not in _sys.path:
            _sys.path.insert(0, _project_dir)
        from sap_agent.sap_chat_tools import query_sap_object
        result = query_sap_object(
            req.object_type,
            req.object_number,
            company_code=req.company_code or None,
        )
        return {
            "object_type": result.object_type,
            "object_number": result.object_number,
            "is_real_data": result.is_real_data,
            "data_blocks": result.data_blocks,
            "error": result.error,
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))

@app.get("/api/environments")
def api_environments() -> dict[str, Any]:
    return {
        "environments": get_available_environments()
    }

# WORKER_REQUIRED_TASKS, get_worker_runtime_status e require_worker_online foram movidos para o topo do ficheiro


@app.get("/api/worker/status")
async def api_worker_status(worker_name: str = None) -> dict[str, Any]:
    status_info = get_worker_runtime_status(worker_name)
    # Add backward compatibility "status" key
    status_info["status"] = status_info["state"]
    return status_info


@app.get("/api/processes")
def api_processes() -> dict[str, Any]:
    return {
        "processes": get_available_processes()
    }


@app.get("/api/subprocesses")
def api_subprocesses(processo: str = "") -> dict[str, Any]:
    return {
        "processo": processo,
        "subprocesses": get_available_subprocesses(processo)
    }



@app.post("/api/upload-file")
async def api_upload_file(file: UploadFile = File(...)) -> dict[str, Any]:
    r"""
    Recebe ficheiro selecionado diretamente no browser.

    O ficheiro é guardado numa pasta montada no Windows:
      host:   C:\workspace\sap-script\sap_script_uploads
      docker: /uploads

    A resposta devolve windows_path, que é o caminho usado pelo worker SAP.
    """
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

    saved_name = _safe_upload_filename(file.filename or "ficheiro")
    target_path = UPLOADS_DIR / saved_name

    content = await file.read()

    if not content:
        raise HTTPException(status_code=400, detail="Ficheiro vazio ou inválido.")

    target_path.write_bytes(content)

    # Clean excel file leading spaces right after upload!
    if saved_name.lower().endswith(".xlsx"):
        try:
            clean_excel_leading_spaces(str(target_path))
        except Exception as exc:
            print(f"[UPLOAD] Erro ao limpar espaços do excel: {exc}")

    return {
        "filename": file.filename,
        "saved_name": saved_name,
        "container_path": str(target_path),
        "windows_path": _windows_upload_path(saved_name),
        "size": len(content),
    }

_KNOWN_JOB_FORM_FIELDS = {
    "task", "ambiente", "processo", "subprocesso",
    "request_option", "request_number", "request_desc",
    "request_type", "caminho_ficheiro", "transacao",
    "nome_pasta", "opcao_processamento",
}

def _validar_e_ajustar_params(params: dict[str, Any]) -> None:
    sub = str(params.get("subprocesso") or "").strip()
    is_cua_remove = sub == "J. CUA_REMOVE.py" or sub == "cua_remove" or "cua_remove" in sub.lower()
    
    if is_cua_remove:
        opcao = params.get("opcao_processamento")
        if not opcao or not str(opcao).strip():
            params["opcao_processamento"] = "sistema_user"
        elif params["opcao_processamento"] not in {"sistema_user", "sistema"}:
            raise ValueError(
                f"Opção de processamento inválida para CUA_REMOVE.\n"
                f"Valores permitidos: sistema_user, sistema."
            )
    else:
        # Se não for CUA_REMOVE, não incluir o parâmetro
        params.pop("opcao_processamento", None)

@app.post("/jobs")
async def create_job_from_form(request: Request) -> dict[str, Any]:
    form = await request.form()
    task = str(form.get("task") or "").strip()
    require_worker_online(task)
    ambiente = str(form.get("ambiente") or "").strip().upper()
    processo = str(form.get("processo") or "").strip()
    subprocesso = str(form.get("subprocesso") or "").strip()
    request_option = str(form.get("request_option") or "4").strip() or "4"
    request_number = str(form.get("request_number") or "").strip().upper()
    request_desc = str(form.get("request_desc") or "").strip()
    request_type = str(form.get("request_type") or "1").strip() or "1"
    caminho_ficheiro = str(form.get("caminho_ficheiro") or "").strip()
    transacao = str(form.get("transacao") or "").strip()
    nome_pasta = str(form.get("nome_pasta") or "").strip()
    opcao_processamento = str(form.get("opcao_processamento") or "").strip()

    params = {
        "ambiente": ambiente,
        "processo": processo,
        "subprocesso": subprocesso,
        "request_option": request_option,
        "request_number": request_number,
        "request_desc": request_desc,
        "request_type": request_type,
        "caminho_ficheiro": caminho_ficheiro,
        "transacao": transacao,
        "nome_pasta": nome_pasta,
        "opcao_processamento": opcao_processamento,
    }

    for key, value in form.multi_items():
        if key not in _KNOWN_JOB_FORM_FIELDS:
            params[key] = str(value).strip()

    try:
        _validar_e_ajustar_params(params)
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err)) from val_err

    return create_job(task=task, params=params)


@app.get("/api/jobs")
def api_list_jobs(limit: int = 50, include_archived: bool = False) -> dict[str, Any]:
    return {"jobs": list_jobs(limit=limit, include_archived=include_archived)}


####################################################################################
# IMPORTANTE:
# Esta rota tem de vir ANTES de /api/jobs/{job_id}
# senão o FastAPI interpreta "next" como job_id e devolve 404.
####################################################################################


@app.get("/api/worker/jobs/next")
def api_worker_claim_next_job(
    worker_name: str = "sap-worker",
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping(worker_name)
    job = claim_next_job(worker_name=worker_name)
    return {"job": job}

@app.get("/api/jobs/next")
def api_claim_next_job(
    worker_name: str = "sap-worker",
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping(worker_name)
    job = claim_next_job(worker_name=worker_name)
    return {"job": job}


@app.get("/api/jobs/{job_id}")
def api_get_job(job_id: str) -> dict[str, Any]:
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado")
    return job


@app.post("/api/jobs/{job_id}/complete")
def api_complete_job(
    job_id: str,
    payload: CompleteJobRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping_for_job(job_id)
    try:
        return complete_job(
            job_id=job_id,
            state=payload.state,
            status=payload.status,
            log=payload.log,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

@app.post("/api/jobs/{job_id}/cancel")
def api_cancel_job(job_id: str) -> dict[str, Any]:
    try:
        return cancel_job(job_id=job_id)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

@app.post("/api/jobs/{job_id}/archive")
def api_archive_job(job_id: str) -> dict[str, Any]:
    try:
        return archive_job(job_id=job_id)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

@app.post("/api/jobs/{job_id}/unarchive")
def api_unarchive_job(job_id: str) -> dict[str, Any]:
    try:
        return unarchive_job(job_id=job_id)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

@app.delete("/api/jobs/{job_id}")
def api_delete_job(job_id: str) -> dict[str, Any]:
    try:
        delete_job(job_id=job_id)
        return {"status": "success", "message": "Job eliminado com sucesso."}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/delete-all")
def api_delete_all_jobs() -> dict[str, Any]:
    try:
        deleted = delete_all_jobs()
        return {"status": "success", "message": "Todos os jobs foram eliminados com sucesso.", "deleted": deleted}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

class AppendLogRequest(BaseModel):
    log_line: str

@app.post("/api/jobs/{job_id}/log")
def api_append_job_log(
    job_id: str,
    payload: AppendLogRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping_for_job(job_id)
    try:
        state = get_job_state(job_id)
        if state is None:
            raise HTTPException(status_code=404, detail="Job não encontrado")
        if state == "failed":
            raise HTTPException(status_code=409, detail="Job has been cancelled or failed.")
        append_job_log(job_id=job_id, log_line=payload.log_line)
        return {
            "status": "ok",
            "job_id": job_id,
            "appended": True
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class SapMetadataRequest(BaseModel):
    sap_system: str
    sap_client: str
    sap_user: str

@app.post("/api/jobs/{job_id}/sap-metadata")
def api_update_sap_metadata(
    job_id: str,
    payload: SapMetadataRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping_for_job(job_id)
    try:
        new_params = {
            "sap_system": payload.sap_system,
            "sap_client": payload.sap_client,
            "sap_user": payload.sap_user,
        }
        return update_job_params(job_id=job_id, new_params=new_params)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class SapLogonDebugRequest(BaseModel):
    sap_logon_debug: dict[str, Any]

@app.post("/api/jobs/{job_id}/sap-logon-debug")
def api_update_sap_logon_debug(
    job_id: str,
    payload: SapLogonDebugRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    update_worker_ping_for_job(job_id)
    try:
        new_params = {
            "sap_logon_debug": payload.sap_logon_debug
        }
        return update_job_params(job_id=job_id, new_params=new_params)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class CreateJobRequest(BaseModel):
    task: str
    params: dict[str, Any] = None

@app.post("/api/jobs")
def api_create_job(payload: CreateJobRequest) -> dict[str, Any]:
    require_worker_online(payload.task)
    try:
        params = payload.params or {}
        _validar_e_ajustar_params(params)
        return create_job(task=payload.task, params=params)
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err)) from val_err
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def validate_worker_token(token: str) -> None:
    if token != WORKER_TOKEN:
        raise HTTPException(status_code=401, detail="Worker token inválido")


class HeartbeatRequest(BaseModel):
    worker_name: str

@app.post("/api/worker/heartbeat")
async def api_worker_heartbeat(
    payload: HeartbeatRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    start_time = time.perf_counter()
    validate_worker_token(x_worker_token)
    
    if not payload.worker_name or not payload.worker_name.strip():
        raise HTTPException(status_code=400, detail="Worker name inválido")
    
    now = time.time()
    
    # Check if the previous ping was delayed
    with worker_last_seen_lock:
        prev_seen = worker_last_seen.get(payload.worker_name)
        
    is_delayed = False
    if prev_seen is not None:
        interval = now - prev_seen
        # If interval between heartbeats > 7.5 seconds (HEARTBEAT_SECONDS * 1.5)
        if interval > 7.5:
            is_delayed = True
            with delayed_heartbeats_lock:
                global delayed_heartbeats_count
                delayed_heartbeats_count += 1
                
    update_worker_ping(payload.worker_name)
    
    duration = time.perf_counter() - start_time
    
    if WORKER_HEARTBEAT_DEBUG:
        now_str = datetime.now().isoformat()
        with worker_last_seen_lock:
            num_workers = len(worker_last_seen)
        print(
            f"[DEBUG HEARTBEAT] Recv: {now_str} | Worker: {payload.worker_name} | "
            f"Duration: {duration:.6f}s | Active Workers: {num_workers} | "
            f"Total Delayed: {delayed_heartbeats_count} (This ping delayed: {is_delayed})"
        )
        
    return {"status": "success", "message": "Heartbeat received"}


# ---------------------------------------------------------------------------
# Agent Context Rules endpoints
# ---------------------------------------------------------------------------

class AgentRuleRequest(BaseModel):
    campo: str
    valor: str
    nome_parametro: str = ""
    processo: str = ""
    subprocesso: str = ""
    transacao_sap: str = ""
    notas: str = ""
    tags: str = ""


def _json_no_store(payload: dict[str, Any], status_code: int = 200) -> JSONResponse:
    response = JSONResponse(content=payload, status_code=status_code)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/api/agent/rules")
def api_list_agent_rules() -> JSONResponse:
    """Lista todas as regras de contexto do Agente SAP."""
    try:
        return _json_no_store({"rules": list_agent_rules()})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/agent/rules")
def api_create_agent_rule(payload: AgentRuleRequest) -> JSONResponse:
    """Cria uma nova regra de contexto."""
    try:
        if (
            not payload.nome_parametro.strip()
            or not payload.campo.strip()
            or not payload.valor.strip()
        ):
            raise HTTPException(
                status_code=400,
                detail="Nome do parametro, Campo e Valor sao obrigatorios.",
            )
        rule = create_agent_rule(
            campo=payload.campo,
            valor=payload.valor,
            nome_parametro=payload.nome_parametro,
            processo=payload.processo,
            subprocesso=payload.subprocesso,
            transacao_sap=payload.transacao_sap,
            notas=payload.notas,
            tags=payload.tags,
        )
        return _json_no_store({"status": "success", "rule": rule})
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.put("/api/agent/rules/{rule_id}")
def api_update_agent_rule(rule_id: str, payload: AgentRuleRequest) -> JSONResponse:
    """Actualiza uma regra de contexto existente."""
    try:
        if (
            not payload.nome_parametro.strip()
            or not payload.campo.strip()
            or not payload.valor.strip()
        ):
            raise HTTPException(
                status_code=400,
                detail="Nome do parametro, Campo e Valor sao obrigatorios.",
            )
        rule = update_agent_rule(
            rule_id=rule_id,
            campo=payload.campo,
            valor=payload.valor,
            nome_parametro=payload.nome_parametro,
            processo=payload.processo,
            subprocesso=payload.subprocesso,
            transacao_sap=payload.transacao_sap,
            notas=payload.notas,
            tags=payload.tags,
        )
        if not rule:
            raise HTTPException(status_code=404, detail="Regra não encontrada.")
        return _json_no_store({"status": "success", "rule": rule})
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.delete("/api/agent/rules/{rule_id}")
def api_delete_agent_rule(rule_id: str) -> JSONResponse:
    """Elimina uma regra de contexto."""
    try:
        delete_agent_rule(rule_id)
        return _json_no_store({"status": "success"})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/api/agent/rules/match")
def api_match_agent_rules(
    processo: str = "",
    ticket_type: str = "",
    stream: str = "",
) -> JSONResponse:
    """Retorna regras que correspondem aos metadados do ticket."""
    try:
        rules = get_agent_rules_for_ticket(
            processo=processo,
            ticket_type=ticket_type,
            stream=stream,
        )
        return _json_no_store({"rules": rules})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
