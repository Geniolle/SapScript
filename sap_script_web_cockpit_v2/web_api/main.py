import importlib
import os
from pathlib import Path
import sys
from uuid import uuid4
from typing import Any
import time

last_worker_ping: float = 0.0

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from web_api.store import append_job_log, cancel_job, claim_next_job, complete_job, create_job, get_job, init_db, list_jobs, archive_job, unarchive_job, delete_job, update_job_params, save_jira_tickets_to_db, list_jira_tickets, update_jira_ticket_assignee, update_jira_ticket_type_db, update_jira_ticket_status_db
from web_api.jira_client import fetch_jira_tickets_from_api, assign_jira_ticket, update_jira_ticket_type, get_jira_issue_transitions, transition_jira_issue
import asyncio

WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
SAP_SCRIPT_PROJECT_DIR = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
UPLOADS_DIR = Path(os.getenv("UPLOADS_DIR", "/uploads"))
UPLOADS_WINDOWS_DIR = os.getenv("UPLOADS_WINDOWS_DIR", "").strip()

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

        subprocessos.append({
            "nome": nome,
            "label": nome,
            "path": caminho,
        })

    return subprocessos



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

async def sync_jira_tickets_loop() -> None:
    """
    Loop em segundo plano que roda a cada 60 segundos buscando os tickets JIRA.
    """
    while True:
        try:
            # Executa a busca HTTP em thread pool para evitar travar o event loop do FastAPI
            tickets = await asyncio.to_thread(fetch_jira_tickets_from_api)
            # Guarda na BD local
            await asyncio.to_thread(save_jira_tickets_to_db, tickets)
        except Exception as exc:
            print(f"[JIRA SYNC LOOP ERROR]: {exc}")
        await asyncio.sleep(60)


@app.on_event("startup")
def startup() -> None:
    init_db()
    asyncio.create_task(sync_jira_tickets_loop())


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    response = templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "ambientes": get_available_environments(),
            "processos": get_available_processes(),
            "jira_base": os.getenv("JIRA_DADOS_COMP_HASH", "https://salsajeans.atlassian.net").strip(),
        },
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/api/jira/tickets")
def api_list_jira_tickets(limit: int = 50) -> dict[str, Any]:
    try:
        return {"tickets": list_jira_tickets(limit=limit)}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/jira/sync")
async def api_force_jira_sync() -> dict[str, Any]:
    try:
        tickets = await asyncio.to_thread(fetch_jira_tickets_from_api)
        await asyncio.to_thread(save_jira_tickets_to_db, tickets)
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

    return {
        "filename": file.filename,
        "saved_name": saved_name,
        "container_path": str(target_path),
        "windows_path": _windows_upload_path(saved_name),
        "size": len(content),
    }

@app.post("/jobs")
def create_job_from_form(
    task: str = Form(...),
    ambiente: str = Form(""),
    processo: str = Form(""),
    subprocesso: str = Form(""),
    request_option: str = Form("4"),
    request_number: str = Form(""),
    request_desc: str = Form(""),
    request_type: str = Form("1"),
    caminho_ficheiro: str = Form(""),
    transacao: str = Form(""),
) -> dict[str, Any]:
    params = {
        "ambiente": ambiente.strip().upper(),
        "processo": processo.strip(),
        "subprocesso": subprocesso.strip(),
        "request_option": request_option.strip() or "4",
        "request_number": request_number.strip().upper(),
        "request_desc": request_desc.strip(),
        "request_type": request_type.strip() or "1",
        "caminho_ficheiro": caminho_ficheiro.strip(),
        "transacao": transacao.strip(),
    }
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


def validate_worker_token(token: str) -> None:
    if token != WORKER_TOKEN:
        raise HTTPException(status_code=401, detail="Worker token inválido")





