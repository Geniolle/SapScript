import importlib
import os
import sys
from typing import Any

from fastapi import FastAPI, Form, Header, HTTPException, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from web_api.store import claim_next_job, complete_job, create_job, get_job, init_db, list_jobs

WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
SAP_SCRIPT_PROJECT_DIR = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()

app = FastAPI(title="SAP Script Web")
app.mount("/static", StaticFiles(directory="web_api/static"), name="static")
templates = Jinja2Templates(directory="web_api/templates")


class CompleteJobRequest(BaseModel):
    state: str
    status: str
    log: str = ""


def _prepare_project_imports() -> None:
    """
    Permite que a API web leia a configuracao real do projeto SAP Script.

    No Docker, o docker-compose monta a raiz do projeto em /sap-script e define
    SAP_SCRIPT_PROJECT_DIR=/sap-script. No Windows/terminal, a mesma variavel
    pode apontar para C:\\workspace\\sap-script ou C:\\SAP Script.
    """
    if SAP_SCRIPT_PROJECT_DIR and SAP_SCRIPT_PROJECT_DIR not in sys.path:
        sys.path.insert(0, SAP_SCRIPT_PROJECT_DIR)


def get_available_environments() -> list[dict[str, str]]:
    """
    Fonte dos ambientes: app/config.py
      AMBIENTES: codigo funcional exibido ao utilizador, ex. DEV/QAD/PRD/CUA
      MAPA_SISTEMA: sistema SAP real, ex. S4D/S4Q/S4P/SPA
      CLIENTES_POR_AMBIENTE: client esperado no login
    """
    _prepare_project_imports()

    try:
        config = importlib.import_module("app.config")
        ambientes = getattr(config, "AMBIENTES", {})
        mapa_sistema = getattr(config, "MAPA_SISTEMA", {})
        clientes = getattr(config, "CLIENTES_POR_AMBIENTE", {})
    except Exception:
        # Fallback defensivo para a pagina continuar funcional mesmo se o volume
        # do projeto SAP Script ainda nao estiver montado no container.
        ambientes = {
            "1": ("DEV", "DESENVOLVIMENTO (S4H)"),
            "2": ("QAD", "QUALIDADE (S4H)"),
            "3": ("PRD", "PRODUCAO (S4H)"),
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

    def sort_key(item: tuple[str, tuple[str, str]]) -> tuple[int, str]:
        numero, valores = item
        try:
            return int(numero), valores[0]
        except Exception:
            return 9999, str(numero)

    result: list[dict[str, str]] = []
    for _numero, valores in sorted(ambientes.items(), key=sort_key):
        codigo = str(valores[0]).strip().upper()
        nome = str(valores[1]).strip()
        sistema = str(mapa_sistema.get(codigo, "")).strip().upper()
        cliente = str(clientes.get(codigo, "")).strip()
        label = f"{codigo} - {nome}"
        if sistema or cliente:
            label += f" | Sistema {sistema or '-'} | Cliente {cliente or '-'}"
        result.append({
            "codigo": codigo,
            "nome": nome,
            "sistema": sistema,
            "cliente": cliente,
            "label": label,
        })

    return result


@app.on_event("startup")
def startup() -> None:
    init_db()


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "ambientes": get_available_environments(),
        },
    )


@app.get("/api/environments")
def api_environments() -> dict[str, Any]:
    return {"environments": get_available_environments()}


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
def api_list_jobs(limit: int = 50) -> dict[str, Any]:
    return {"jobs": list_jobs(limit=limit)}


@app.get("/api/jobs/{job_id}")
def api_get_job(job_id: str) -> dict[str, Any]:
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado")
    return job


@app.get("/api/jobs/next")
def api_claim_next_job(
    worker_name: str = "sap-worker",
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    job = claim_next_job(worker_name=worker_name)
    return {"job": job}


@app.post("/api/jobs/{job_id}/complete")
def api_complete_job(
    job_id: str,
    payload: CompleteJobRequest,
    x_worker_token: str = Header(default=""),
) -> dict[str, Any]:
    validate_worker_token(x_worker_token)
    try:
        return complete_job(
            job_id=job_id,
            state=payload.state,
            status=payload.status,
            log=payload.log,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def validate_worker_token(token: str) -> None:
    if token != WORKER_TOKEN:
        raise HTTPException(status_code=401, detail="Worker token inválido")
