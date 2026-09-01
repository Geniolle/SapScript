import ast
import dataclasses
import importlib
import json
import os
from contextlib import asynccontextmanager
from datetime import date, timedelta
from pathlib import Path
import sys
import requests

from uuid import uuid4
from typing import Any
import time

from dotenv import load_dotenv

last_worker_ping: float = 0.0
_JIRA_ENV_SOURCE = "process environment"
_JIRA_ENV_KEYS = (
    "JIRA_DADOS_COMP_HASH",
    "JIRA_EMAIL",
    "JIRA_TOKEN",
    "JIRA_DADOS_HASH",
    "JIRA_SYNC_PROJECTS",
)


def _load_project_env() -> None:
    global _JIRA_ENV_SOURCE
    module_dir = Path(__file__).resolve().parent
    project_root = module_dir.parent
    candidate_paths = [
        project_root / ".env",
        project_root.parent / ".env",
        Path("/sap-script/.env"),
        Path("/srv/sap-script-web/.env"),
    ]

    for env_path in candidate_paths:
        try:
            if env_path.is_file():
                load_dotenv(dotenv_path=env_path, override=True)
                _JIRA_ENV_SOURCE = str(env_path)
                return
        except Exception:
            continue

    load_dotenv(override=True)
    _JIRA_ENV_SOURCE = "process environment / fallback .env"


def _log_jira_env_boot_status() -> None:
    present = [key for key in _JIRA_ENV_KEYS if os.getenv(key, "").strip()]
    missing = [key for key in _JIRA_ENV_KEYS if not os.getenv(key, "").strip()]
    print(
        "[JIRA ENV BOOT] source="
        f"{_JIRA_ENV_SOURCE} "
        f"present={present or '[]'} "
        f"missing={missing or '[]'}"
    )


_load_project_env()

from fastapi import FastAPI, File, Header, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from web_api.store import append_job_log, cancel_job, claim_next_job, complete_job, create_job, get_job, init_db, list_jobs, archive_job, unarchive_job, delete_job, update_job_params, save_jira_tickets_to_db, list_jira_tickets, update_jira_ticket_assignee, update_jira_ticket_type_db, update_jira_ticket_status_db, update_jira_ticket_supplier_db, log_auto_trigger_entry, list_auto_trigger_log, has_active_job_for_ticket, clear_auto_trigger_log, delete_auto_trigger_log_entry, get_latest_sap_agent_analysis, save_jira_ticket_batch_only, create_agent_rule, list_agent_rules, update_agent_rule, delete_agent_rule, get_agent_rules_for_ticket, get_transacao_by_processo
from web_api.jira_client import fetch_jira_tickets_from_api, assign_jira_ticket, update_jira_ticket_type, get_jira_issue_transitions, transition_jira_issue, update_jira_ticket_supplier, fetch_ticket_details, add_jira_comment, clean_excel_leading_spaces
import asyncio

WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
SAP_SCRIPT_PROJECT_DIR = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
UPLOADS_DIR = Path(os.getenv("UPLOADS_DIR", "/uploads"))
UPLOADS_WINDOWS_DIR = os.getenv("UPLOADS_WINDOWS_DIR", "").strip()

# Intervalo do loop de sincronização JIRA em segundo plano.
POLL_SECONDS = max(1, int(os.getenv("POLL_SECONDS", "60")))

# Diretório de download de anexos JIRA
# No container Docker: /data/jira  (montado a partir de C:\Jira no host Windows)
JIRA_DOWNLOAD_DIR_CONTAINER = os.getenv("JIRA_DOWNLOAD_DIR_CONTAINER", "/data/jira").strip()
JIRA_DOWNLOAD_DIR_WINDOWS = os.getenv("JIRA_DOWNLOAD_DIR_WINDOWS", r"C:\Jira").strip()

@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    _log_jira_env_boot_status()
    background_tasks = [asyncio.create_task(sync_jira_tickets_loop())]
    try:
        yield
    finally:
        for task in background_tasks:
            task.cancel()
        await asyncio.gather(*background_tasks, return_exceptions=True)


app = FastAPI(title="SAP Script Web", lifespan=lifespan)
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


def _get_fi_default_context() -> dict[str, Any]:
    from sap_rfc.fi_config import get_fi_default_context

    return get_fi_default_context()


def _default_f110_next_due_date() -> str:
    return (date.today() + timedelta(days=1)).isoformat()


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

    if caminho != processos_dir_abs and not caminho.startswith(processos_dir_abs + os.sep):
        return None

    if not os.path.isdir(caminho):
        return None

    return caminho


def get_available_processes() -> list[dict[str, str]]:
    processos_dir = _resolve_processes_dir()

    if not processos_dir:
        return []

    processos: list[dict[str, str]] = []

    for nome in sorted(os.listdir(processos_dir), key=str.casefold):
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
    # Importação ampla: apenas tickets ativos dos projetos configurados.
    # Tickets resolvidos/concluídos ficam fora da sync e não são preservados na BD.
    return fetch_jira_tickets_from_api()


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
        await asyncio.sleep(POLL_SECONDS)


async def historical_jira_sync() -> None:
    """
    Sincronização histórica desativada.
    Tickets resolvidos/concluídos não são importados pelo cockpit.
    """
    print("[JIRA HISTORICAL SYNC] Desativada: apenas tickets ativos são importados.")
    return


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    response = templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "ambientes": get_available_environments(),
            "processos": get_available_processes(),
            "fi_defaults": _get_fi_default_context(),
            "jira_base": os.getenv("JIRA_DADOS_COMP_HASH", "https://salsajeans.atlassian.net").strip(),
            "poll_seconds": POLL_SECONDS,
        },
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


@app.post("/api/sap-agent/analyze/{ticket_key}")
def api_sap_agent_analyze(ticket_key: str) -> dict[str, Any]:
    """Cria um job técnico para o worker Windows executar a análise do Agente SAP no ticket indicado."""
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


class SalsaItPfcgAnalyzeRequest(BaseModel):
    role_name: str


class SalsaItPfcgCreateAnalyzeRequest(BaseModel):
    selection_id: str
    role_name: str


class SalsaItPfcgCreateRfcPreviewRequest(BaseModel):
    role_name: str
    description: str
    tcodes: list[str] = []
    transport_mode: str = "LOCAL"
    request_number: str = ""
    request_description: str = ""


class SalsaItPfcgCreateRfcConfirmRequest(BaseModel):
    preview_job_id: str


class SalsaItPfcgDeleteRfcPreviewRequest(BaseModel):
    role_name: str
    transport_mode: str = "LOCAL"
    request_number: str = ""
    request_description: str = ""


class SalsaItPfcgDeleteRfcConfirmRequest(BaseModel):
    preview_job_id: str


class SalsaItPfcgCompostaPreviewRequest(BaseModel):
    role_name: str
    description: str
    child_roles: list[str] = []
    transport_mode: str = "LOCAL"
    request_number: str = ""
    request_description: str = ""


class SalsaItPfcgCompostaConfirmRequest(BaseModel):
    preview_job_id: str


PFCG_EXCEL_SELECTIONS: dict[str, dict[str, str]] = {}
PFCG_RFC_CREATE_PREVIEWS: dict[str, dict[str, Any]] = {}
PFCG_COMPOSTA_CREATE_PREVIEWS: dict[str, dict[str, Any]] = {}

PFCG_RFC_CREATE_ENVIRONMENT = "DEV"
PFCG_RFC_DELETE_PREVIEWS: dict[str, dict[str, Any]] = {}
PFCG_RFC_DELETE_ENVIRONMENT = "DEV"


def _validate_pfcg_role_name_or_400(role_name: str) -> str:
    try:
        from sap_rfc import validate_role_name
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail="Não foi possível carregar a validação PFCG no backend.",
        ) from exc

    try:
        return validate_role_name(role_name)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def _safe_pfcg_failed_message() -> str:
    return "Não foi possível concluir a análise PFCG."


@app.post("/api/salsa-it-agent/pfcg/analyze")
def api_salsa_it_pfcg_analyze(payload: SalsaItPfcgAnalyzeRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)

    try:
        job = create_job("pfcg_role_analysis", {"role_name": role_name})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store(
        {
            "job_id": job["id"],
            "state": job["state"],
            "role_name": role_name,
        }
    )


@app.get("/api/salsa-it-agent/pfcg/analyze/{job_id}")
def api_salsa_it_pfcg_analyze_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_analysis":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à análise PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state == "failed":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if state != "succeeded":
        return _json_no_store({"state": state, "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    if not status_raw:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    try:
        result = json.loads(status_raw)
    except Exception:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "role": str(result.get("role") or ""),
        "description": result.get("description"),
        "language": result.get("language"),
        "system": result.get("system"),
        "client": result.get("client"),
    }

    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")

    return _json_no_store({"state": "succeeded", "result": safe_result})


def _safe_pfcg_sub_result(result: dict[str, Any], *, items_key: str, item_fields: tuple[str, ...]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "role": str(result.get("role") or ""),
        "count": result.get("count"),
        "system": result.get("system"),
        "client": result.get("client"),
        "is_composite": bool(result.get("is_composite")),
    }
    if safe_result["is_composite"]:
        composite_members = result.get("composite_members")
        safe_result["composite_members"] = composite_members if isinstance(composite_members, list) else []
    if result.get("warning"):
        safe_result["warning"] = result.get("warning")

    raw_items = result.get(items_key)
    if isinstance(raw_items, list):
        safe_result[items_key] = [
            {field: item.get(field) for field in item_fields}
            for item in raw_items
            if isinstance(item, dict)
        ]
    else:
        safe_result[items_key] = []

    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")

    return safe_result


@app.post("/api/salsa-it-agent/pfcg/transactions/analyze")
def api_salsa_it_pfcg_transactions_analyze(payload: SalsaItPfcgAnalyzeRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)

    try:
        job = create_job("pfcg_role_transactions_analysis", {"role_name": role_name})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store(
        {
            "job_id": job["id"],
            "state": job["state"],
            "role_name": role_name,
        }
    )


@app.get("/api/salsa-it-agent/pfcg/transactions/analyze/{job_id}")
def api_salsa_it_pfcg_transactions_analyze_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_transactions_analysis":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à análise de transações PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state == "failed":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if state != "succeeded":
        return _json_no_store({"state": state, "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    if not status_raw:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    try:
        result = json.loads(status_raw)
    except Exception:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_sub_result(result, items_key="transactions", item_fields=("tcode", "description"))
    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/salsa-it-agent/pfcg/users/analyze")
def api_salsa_it_pfcg_users_analyze(payload: SalsaItPfcgAnalyzeRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)

    try:
        job = create_job("pfcg_role_users_analysis", {"role_name": role_name})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store(
        {
            "job_id": job["id"],
            "state": job["state"],
            "role_name": role_name,
        }
    )


@app.get("/api/salsa-it-agent/pfcg/users/analyze/{job_id}")
def api_salsa_it_pfcg_users_analyze_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_users_analysis":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à análise de utilizadores PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state == "failed":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if state != "succeeded":
        return _json_no_store({"state": state, "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    if not status_raw:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    try:
        result = json.loads(status_raw)
    except Exception:
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_sub_result(
        result,
        items_key="users",
        item_fields=("username", "valid_from", "valid_to", "assignment_status"),
    )
    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/salsa-it-agent/pfcg/create/select-excel")
def api_salsa_it_pfcg_create_select_excel() -> JSONResponse:
    try:
        job = create_job("select_excel_file", {})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao criar job de seleção de Excel: {str(exc)}")

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
    })


@app.get("/api/salsa-it-agent/pfcg/create/select-excel/{job_id}")
def api_salsa_it_pfcg_create_select_excel_job(job_id: str) -> JSONResponse:
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado.")

    if job.get("task") != "select_excel_file":
        raise HTTPException(status_code=400, detail="O job não pertence ao fluxo de seleção de Excel.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        failure_message = str(job.get("status") or "Não foi possível selecionar o ficheiro Excel.").strip()
        return _json_no_store({
            "state": "failed",
            "message": failure_message,
        })

    selected_path = str(job.get("status") or "").strip()
    if not selected_path:
        return _json_no_store({
            "state": "failed",
            "message": "Seleção de ficheiro Excel cancelada ou vazia.",
        })

    selection_id = job_id
    PFCG_EXCEL_SELECTIONS[selection_id] = {
        "excel_path": selected_path,
        "file_name": Path(selected_path).name,
    }
    return _json_no_store({
        "state": "succeeded",
        "selection_id": selection_id,
        "file_name": Path(selected_path).name,
    })


@app.post("/api/salsa-it-agent/pfcg/create/analyze")
def api_salsa_it_pfcg_create_analyze(payload: SalsaItPfcgCreateAnalyzeRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)
    selection_id = str(payload.selection_id or "").strip()
    if not selection_id:
        raise HTTPException(status_code=400, detail="Seleção de Excel inválida.")

    selection = PFCG_EXCEL_SELECTIONS.get(selection_id)
    if not selection:
        raise HTTPException(status_code=404, detail="Seleção de Excel não encontrada.")

    excel_path = selection.get("excel_path", "")
    if not excel_path:
        raise HTTPException(status_code=400, detail="Caminho do Excel indisponível.")

    try:
        job = create_job(
            "pfcg_create_excel_analysis",
            {
                "excel_path": excel_path,
                "role_name": role_name,
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao criar job de análise do Excel: {str(exc)}")

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": role_name,
        "file_name": selection.get("file_name") or Path(excel_path).name,
    })


@app.get("/api/salsa-it-agent/pfcg/create/analyze/{job_id}")
def api_salsa_it_pfcg_create_analyze_job(job_id: str) -> JSONResponse:
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado.")

    if job.get("task") != "pfcg_create_excel_analysis":
        raise HTTPException(status_code=400, detail="O job não pertence ao fluxo de preparação do Excel.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        failure_message = str(job.get("status") or "Não foi possível concluir a preparação do Perfil de Autorização.").strip()
        return _json_no_store({
            "state": "failed",
            "message": failure_message,
        })

    raw_status = job.get("status")
    try:
        result = json.loads(raw_status) if isinstance(raw_status, str) else raw_status
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({
            "state": "failed",
            "message": "Resultado de análise inválido.",
        })

    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": result.get("status"),
        "role": result.get("role"),
        "description": result.get("description"),
        "language": result.get("language"),
        "system": result.get("system"),
        "client": result.get("client"),
        "sheet": result.get("sheet"),
        "summary": result.get("summary"),
        "warnings": result.get("warnings") or [],
        "errors": result.get("errors") or [],
    }
    if result.get("role_in_excel") is not None:
        safe_result["role_in_excel"] = result.get("role_in_excel")
    return _json_no_store({
        "state": "succeeded",
        "result": safe_result,
    })



def _safe_pfcg_rfc_delete_result(result: dict[str, Any]) -> dict[str, Any]:
    return {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or "ERROR"),
        "environment": str(result.get("environment") or PFCG_RFC_DELETE_ENVIRONMENT),
        "role": str(result.get("role") or ""),
        "description": result.get("description"),
        "tcodes": result.get("tcodes") or [],
        "tcodes_count": result.get("tcodes_count"),
        "users_count": result.get("users_count"),
        "transport": result.get("transport"),
        "transport_mode": result.get("transport_mode"),
        "transport_request": result.get("transport_request"),
        "error_type": result.get("error_type"),
        "message": result.get("message"),
    }

def _safe_pfcg_rfc_create_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "environment": result.get("environment"),
        "role": result.get("role"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        if result.get("missing_tcodes"):
            safe_result["missing_tcodes"] = result.get("missing_tcodes")
        return safe_result

    # Campos apenas do fluxo de sucesso (preview e/ou criação real)
    for field in (
        "description",
        "tcodes",
        "tcodes_count",
        "tcodes_requested",
        "tcodes_created",
        "profile_generated",
        "transport",
        "transport_mode",
        "transport_request",
        "transport_request_created",
    ):
        if field in result:
            safe_result[field] = result.get(field)
    return safe_result


def _safe_pfcg_transport_search_result(result: dict[str, Any]) -> dict[str, Any]:
    safe_result: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "status": str(result.get("status") or ""),
        "environment": result.get("environment"),
    }
    if not safe_result["ok"]:
        safe_result["error_type"] = result.get("error_type")
        safe_result["message"] = result.get("message")
        return safe_result

    safe_result["owner"] = result.get("owner")
    safe_result["requests_count"] = result.get("requests_count")
    safe_result["requests"] = [
        {
            "request": row.get("request"),
            "description": row.get("description"),
            "trtype": row.get("trtype"),
            "target_system": row.get("target_system"),
            "state": row.get("state"),
        }
        for row in (result.get("requests") or [])
        if isinstance(row, dict)
    ]
    return safe_result




@app.post("/api/salsa-it-agent/pfcg/delete/rfc/preview")
def api_salsa_it_pfcg_delete_rfc_preview(payload: SalsaItPfcgDeleteRfcPreviewRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)

    try:
        job = create_job(
            "pfcg_role_delete_preview",
            {
                "environment": PFCG_RFC_DELETE_ENVIRONMENT,
                "role_name": role_name,
                "transport_mode": str(payload.transport_mode or "LOCAL").strip().upper(),
                "request_number": str(payload.request_number or "").strip(),
                "request_description": str(payload.request_description or "").strip(),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": role_name,
    })


@app.get("/api/salsa-it-agent/pfcg/delete/rfc/preview/{job_id}")
def api_salsa_it_pfcg_delete_rfc_preview_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_delete_preview":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à pré-visualização de eliminação PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_rfc_delete_result(result)

    if safe_result.get("ok") and safe_result.get("status") == "PREVIEW_READY":
        transport_preview = result.get("transport") or {}
        PFCG_RFC_DELETE_PREVIEWS[job_id] = {
            "environment": result.get("environment"),
            "role_name": result.get("role"),
            "transport_mode": str(transport_preview.get("transport_mode") or "LOCAL"),
            "request_number": str(transport_preview.get("request_number") or ""),
            "request_description": str(transport_preview.get("request_description") or ""),
        }

    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/salsa-it-agent/pfcg/delete/rfc/confirm")
def api_salsa_it_pfcg_delete_rfc_confirm(payload: SalsaItPfcgDeleteRfcConfirmRequest) -> JSONResponse:
    preview_job_id = str(payload.preview_job_id or "").strip()
    if not preview_job_id:
        raise HTTPException(status_code=400, detail="Identificador da pré-visualização em falta.")

    validated = PFCG_RFC_DELETE_PREVIEWS.get(preview_job_id)
    if not validated:
        raise HTTPException(
            status_code=404,
            detail="Pré-visualização não encontrada ou expirada. Repita a preparação antes de confirmar.",
        )

    try:
        job = create_job("pfcg_role_delete_rfc", dict(validated))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": validated.get("role_name"),
    })


@app.get("/api/salsa-it-agent/pfcg/delete/rfc/confirm/{job_id}")
def api_salsa_it_pfcg_delete_rfc_confirm_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_delete_rfc":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à confirmação de eliminação PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    return _json_no_store({"state": "succeeded", "result": _safe_pfcg_rfc_delete_result(result)})


@app.post("/api/salsa-it-agent/pfcg/create/rfc/preview")
def api_salsa_it_pfcg_create_rfc_preview(payload: SalsaItPfcgCreateRfcPreviewRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)
    description = str(payload.description or "").strip()
    if not description:
        raise HTTPException(status_code=400, detail="Informe uma descrição para o Perfil de Autorização.")

    try:
        job = create_job(
            "pfcg_role_create_preview",
            {
                "environment": PFCG_RFC_CREATE_ENVIRONMENT,
                "role_name": role_name,
                "description": description,
                "tcodes": list(payload.tcodes or []),
                "transport_mode": str(payload.transport_mode or "LOCAL").strip().upper(),
                "request_number": str(payload.request_number or "").strip(),
                "request_description": str(payload.request_description or "").strip(),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": role_name,
    })


@app.get("/api/salsa-it-agent/pfcg/create/rfc/preview/{job_id}")
def api_salsa_it_pfcg_create_rfc_preview_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_create_preview":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à pré-visualização de criação PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_rfc_create_result(result)

    if safe_result.get("ok") and safe_result.get("status") == "PREVIEW_READY":
        transport_preview = result.get("transport") or {}
        PFCG_RFC_CREATE_PREVIEWS[job_id] = {
            "environment": result.get("environment"),
            "role_name": result.get("role"),
            "description": result.get("description"),
            "tcodes": list(result.get("tcodes") or []),
            "transport_mode": str(transport_preview.get("transport_mode") or "LOCAL"),
            "request_number": str(transport_preview.get("request_number") or ""),
            "request_description": str(transport_preview.get("request_description") or ""),
        }

    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/salsa-it-agent/pfcg/create/rfc/confirm")
def api_salsa_it_pfcg_create_rfc_confirm(payload: SalsaItPfcgCreateRfcConfirmRequest) -> JSONResponse:
    preview_job_id = str(payload.preview_job_id or "").strip()
    if not preview_job_id:
        raise HTTPException(status_code=400, detail="Identificador da pré-visualização em falta.")

    validated = PFCG_RFC_CREATE_PREVIEWS.get(preview_job_id)
    if not validated:
        raise HTTPException(
            status_code=404,
            detail="Pré-visualização não encontrada ou expirada. Repita a preparação antes de confirmar.",
        )

    try:
        job = create_job("pfcg_role_create_rfc", dict(validated))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": validated.get("role_name"),
    })


@app.get("/api/salsa-it-agent/pfcg/create/rfc/confirm/{job_id}")
def api_salsa_it_pfcg_create_rfc_confirm_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_role_create_rfc":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à criação individual PFCG (RFC).")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_rfc_create_result(result)
    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/salsa-it-agent/pfcg/composta/preview")
def api_salsa_it_pfcg_composta_preview(payload: SalsaItPfcgCompostaPreviewRequest) -> JSONResponse:
    role_name = _validate_pfcg_role_name_or_400(payload.role_name)
    description = str(payload.description or "").strip()
    if not description:
        raise HTTPException(status_code=400, detail="Informe uma descrição para a Função Composta.")
    if not payload.child_roles:
        raise HTTPException(status_code=400, detail="Informe pelo menos uma função componente (role filha).")

    try:
        job = create_job(
            "pfcg_composta_create_preview",
            {
                "environment": "DEV",
                "role_name": role_name,
                "description": description,
                "child_roles": [str(r).strip().upper() for r in payload.child_roles if str(r).strip()],
                "transport_mode": str(payload.transport_mode or "LOCAL").strip().upper(),
                "request_number": str(payload.request_number or "").strip(),
                "request_description": str(payload.request_description or "").strip(),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": role_name,
    })


@app.get("/api/salsa-it-agent/pfcg/composta/preview/{job_id}")
def api_salsa_it_pfcg_composta_preview_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_composta_create_preview":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à pré-visualização de Função Composta.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    if result.get("ok") and result.get("status") == "PREVIEW_READY":
        transport_preview = result.get("transport") or {}
        PFCG_COMPOSTA_CREATE_PREVIEWS[job_id] = {
            "environment": "DEV",
            "role_name": result.get("role"),
            "description": result.get("description"),
            "child_roles": list(result.get("child_roles") or []),
            "transport_mode": str(transport_preview.get("transport_mode") or "LOCAL"),
            "request_number": str(transport_preview.get("request_number") or ""),
            "request_description": str(transport_preview.get("request_description") or ""),
        }

    return _json_no_store({"state": "succeeded", "result": result})


@app.post("/api/salsa-it-agent/pfcg/composta/confirm")
def api_salsa_it_pfcg_composta_confirm(payload: SalsaItPfcgCompostaConfirmRequest) -> JSONResponse:
    preview_job_id = str(payload.preview_job_id or "").strip()
    if not preview_job_id:
        raise HTTPException(status_code=400, detail="Identificador da pré-visualização em falta.")

    validated = PFCG_COMPOSTA_CREATE_PREVIEWS.get(preview_job_id)
    if not validated:
        raise HTTPException(
            status_code=404,
            detail="Pré-visualização não encontrada ou expirada. Repita a preparação antes de confirmar.",
        )

    try:
        job = create_job("pfcg_composta_create", dict(validated))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({
        "job_id": job["id"],
        "state": job["state"],
        "role_name": validated.get("role_name"),
    })


@app.get("/api/salsa-it-agent/pfcg/composta/confirm/{job_id}")
def api_salsa_it_pfcg_composta_confirm_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_composta_create":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à criação de Função Composta.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "succeeded", "result": {"ok": True, "status": "SUCESSO", "message": "Função Composta criada com sucesso em DEV."}})

    return _json_no_store({"state": "succeeded", "result": result})


@app.post("/api/salsa-it-agent/pfcg/transport/search")
def api_salsa_it_pfcg_transport_search() -> JSONResponse:
    """Pesquisa (read-only via RFC) das Requests de transporte abertas do utilizador RFC em DEV.

    Endpoint de propósito fixo: não aceita nenhum parâmetro do cliente — o ambiente é sempre
    PFCG_RFC_CREATE_ENVIRONMENT (DEV) e a função RFC a chamar é decidida inteiramente dentro
    de sap_rfc.pfcg_transport_service.
    """
    try:
        job = create_job("pfcg_transport_search", {"environment": PFCG_RFC_CREATE_ENVIRONMENT})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return _json_no_store({"job_id": job["id"], "state": job["state"]})


@app.get("/api/salsa-it-agent/pfcg/transport/search/{job_id}")
def api_salsa_it_pfcg_transport_search_job(job_id: str) -> JSONResponse:
    try:
        job = get_job(job_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} não encontrado.")
    if job.get("task") != "pfcg_transport_search":
        raise HTTPException(status_code=400, detail="O job indicado não pertence à pesquisa de Requests PFCG.")

    state = str(job.get("state") or "pending")
    if state in {"pending", "running"}:
        return _json_no_store({"state": state})

    if state != "succeeded":
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    status_raw = str(job.get("status") or "").strip()
    try:
        result = json.loads(status_raw) if status_raw else None
    except Exception:
        result = None

    if not isinstance(result, dict):
        return _json_no_store({"state": "failed", "message": _safe_pfcg_failed_message()})

    safe_result = _safe_pfcg_transport_search_result(result)
    return _json_no_store({"state": "succeeded", "result": safe_result})


@app.post("/api/sap-agent/chat")
def api_sap_agent_chat(request: SapAgentChatRequest) -> dict[str, Any]:
    """Conversação interativa com o Gemini com base no contexto do ticket e nos sinais SAP extraídos."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=400,
            detail="GEMINI_API_KEY não configurada no ficheiro .env. Por favor, adicione a chave e reinicie o cockpit.",
        )

    # 1. Obter detalhes do ticket no JIRA
    ticket_info = fetch_ticket_details(request.ticket_key)
    summary = ticket_info.get("summary") or "Sem sumário"
    description = ticket_info.get("description") or "Sem descrição"
    comments_list = ticket_info.get("comments") or []
    comments_text = "\n".join(comments_list) if comments_list else "Sem comentários"

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
                        return {
                            "reply": f"❌ Não foi possível criar job SAP: {job_exc}\n\nAcesso manual: {action_desc}"
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

@app.get("/api/worker/status")
def api_worker_status() -> dict[str, Any]:
    global last_worker_ping
    is_online = (time.time() - last_worker_ping) < 15.0
    return {"status": "online" if is_online else "offline"}


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
    """
    Recebe ficheiro selecionado diretamente no browser.

    O ficheiro é guardado numa pasta montada no Windows:
      host:   C:\\workspace\\sap-script\\sap_script_uploads
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
    "nome_pasta",
}

@app.post("/jobs")
async def create_job_from_form(request: Request) -> dict[str, Any]:
    form = await request.form()
    task = str(form.get("task") or "").strip()
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
    }

    for key, value in form.multi_items():
        if key not in _KNOWN_JOB_FORM_FIELDS:
            params[key] = str(value).strip()

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
    global last_worker_ping
    validate_worker_token(x_worker_token)
    last_worker_ping = time.time()
    job = claim_next_job(worker_name=worker_name)
    return {"job": job}

@app.get("/api/jobs/next")
def api_claim_next_job(
    worker_name: str = "sap-worker",
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    global last_worker_ping
    validate_worker_token(x_worker_token)
    last_worker_ping = time.time()
    job = claim_next_job(worker_name=worker_name)
    return {"job": job}


@app.get("/api/jobs/{job_id}")
def api_get_job(job_id: str) -> dict[str, Any]:
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado")
    return job


class UpdateJobParamsRequest(BaseModel):
    params: dict[str, Any]


@app.post("/api/jobs/{job_id}/params")
def api_update_job_params(
    job_id: str,
    payload: UpdateJobParamsRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    global last_worker_ping
    last_worker_ping = time.time()
    try:
        return update_job_params(job_id=job_id, new_params=payload.params)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/jobs/{job_id}/complete")
def api_complete_job(
    job_id: str,
    payload: CompleteJobRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    global last_worker_ping
    last_worker_ping = time.time()
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

class AppendLogRequest(BaseModel):
    log_line: str

@app.post("/api/jobs/{job_id}/log")
def api_append_job_log(
    job_id: str,
    payload: AppendLogRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    global last_worker_ping
    last_worker_ping = time.time()
    try:
        job = get_job(job_id)
        if job and job["state"] == "failed":
            raise HTTPException(status_code=409, detail="Job has been cancelled or failed.")
        return append_job_log(job_id=job_id, log_line=payload.log_line)
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
    global last_worker_ping
    last_worker_ping = time.time()
    try:
        new_params = {
            "sap_system": payload.sap_system,
            "sap_client": payload.sap_client,
            "sap_user": payload.sap_user,
        }
        return update_job_params(job_id=job_id, new_params=new_params)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class CreateJobRequest(BaseModel):
    task: str
    params: dict[str, Any] = None

@app.post("/api/jobs")
def api_create_job(payload: CreateJobRequest) -> dict[str, Any]:
    try:
        return create_job(task=payload.task, params=payload.params or {})
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class FiDefaultDocumentRequest(BaseModel):
    environment: str
    branch: str
    payload: dict[str, Any] = None


async def _wait_for_job_terminal_state(job_id: str, timeout_seconds: int) -> dict[str, Any]:
    deadline = time.monotonic() + max(1, int(timeout_seconds))
    while time.monotonic() < deadline:
        job = await asyncio.to_thread(get_job, job_id)
        if job and str(job.get("state") or "").strip() in {"succeeded", "failed"}:
            return job
        await asyncio.sleep(1.0)
    raise TimeoutError(
        f"Timeout a aguardar o job {job_id} terminar no worker Windows."
    )


@app.post("/api/fi/default-document")
async def api_create_fi_default_document(payload: FiDefaultDocumentRequest) -> JSONResponse:
    try:
        fi_payload = dict(payload.payload or {"data_mode": "default"})
        fi_payload.setdefault("data_mode", "default")
        fi_payload.setdefault("environment", payload.environment)
        fi_payload.setdefault("branch", payload.branch)

        job = create_job(
            task="fi_default_document",
            params={
                "environment": payload.environment,
                "branch": payload.branch,
                "payload": fi_payload,
            },
        )

        timeout_seconds = int(os.getenv("FI_DEFAULT_DOCUMENT_TIMEOUT_SECONDS", "900"))
        finished_job = await _wait_for_job_terminal_state(job["id"], timeout_seconds)
        if str(finished_job.get("state") or "").strip() != "succeeded":
            raise HTTPException(
                status_code=400,
                detail=(
                    str(finished_job.get("status") or "Falha ao executar o documento FI.")
                    + (
                        f"\n{finished_job.get('log')}"
                        if str(finished_job.get("log") or "").strip()
                        else ""
                    )
                ),
            )

        result = (finished_job.get("params") or {}).get("fi_document_result")
        if not isinstance(result, dict):
            raise HTTPException(
                status_code=500,
                detail="Worker Windows concluiu o job, mas não devolveu o resultado FI.",
            )

        return _json_no_store(result)
    except TimeoutError as exc:
        raise HTTPException(status_code=504, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


class F110ProposalRequest(BaseModel):
    environment: str
    operation_type: str
    company_code: str
    payment_method: str
    account_number: str
    posting_date: str
    next_due_date: str = ""
    document_number: str = ""
    source_payload: dict[str, Any] | None = None


@app.post("/api/f110/payment")
async def api_run_f110_payment(payload: F110ProposalRequest) -> JSONResponse:
    """Executa o ciclo de pagamento do F110 via RFF110S."""
    try:
        _prepare_project_imports()

        job = create_job(
            task="f110_payment",
            params={
                "environment": payload.environment,
                "operation_type": payload.operation_type,
                "company_code": payload.company_code,
                "payment_method": payload.payment_method,
                "account_number": payload.account_number,
                "posting_date": payload.posting_date,
                "next_due_date": payload.next_due_date,
                "document_number": payload.document_number,
                "source_payload": payload.source_payload or {},
            },
        )

        timeout_seconds = int(os.getenv("FI_PAYMENT_RUN_TIMEOUT_SECONDS", "900"))
        finished_job = await _wait_for_job_terminal_state(job["id"], timeout_seconds)
        if str(finished_job.get("state") or "").strip() != "succeeded":
            raise HTTPException(
                status_code=400,
                detail=(
                    str(finished_job.get("status") or "Falha ao executar o pagamento F110.")
                    + (
                        f"\n{finished_job.get('log')}"
                        if str(finished_job.get("log") or "").strip()
                        else ""
                    )
                ),
            )

        result = (finished_job.get("params") or {}).get("f110_payment_result")
        if not isinstance(result, dict):
            raise HTTPException(
                status_code=500,
                detail="Worker Windows concluiu o job, mas não devolveu o resultado do pagamento F110.",
            )

        return _json_no_store(result)
    except TimeoutError as exc:
        raise HTTPException(status_code=504, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/f110/proposal")
async def api_run_f110_proposal(payload: F110ProposalRequest) -> JSONResponse:
    """Executa apenas a PROPOSTA (Vorlauf) do F110 via RFF110S. Nunca dispara o pagamento/cobrança real."""
    try:
        _prepare_project_imports()

        next_due_date = str(payload.next_due_date or "").strip() or _default_f110_next_due_date()

        job = create_job(
            task="f110_proposal",
            params={
                "environment": payload.environment,
                "operation_type": payload.operation_type,
                "company_code": payload.company_code,
                "payment_method": payload.payment_method,
                "account_number": payload.account_number,
                "posting_date": payload.posting_date,
                "next_due_date": next_due_date,
                "document_number": payload.document_number,
                "source_payload": payload.source_payload or {},
            },
        )

        timeout_seconds = int(os.getenv("F110_PROPOSAL_TIMEOUT_SECONDS", "900"))
        finished_job = await _wait_for_job_terminal_state(job["id"], timeout_seconds)
        if str(finished_job.get("state") or "").strip() != "succeeded":
            raise HTTPException(
                status_code=400,
                detail=(
                    str(finished_job.get("status") or "Falha ao executar a proposta F110.")
                    + (
                        f"\n{finished_job.get('log')}"
                        if str(finished_job.get("log") or "").strip()
                        else ""
                    )
                ),
            )

        result = (finished_job.get("params") or {}).get("f110_proposal_result")
        if not isinstance(result, dict):
            raise HTTPException(
                status_code=500,
                detail="Worker Windows concluiu o job, mas não devolveu o resultado F110.",
            )

        return _json_no_store(result)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def validate_worker_token(token: str) -> None:
    if token != WORKER_TOKEN:
        raise HTTPException(status_code=401, detail="Worker token inválido")


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
