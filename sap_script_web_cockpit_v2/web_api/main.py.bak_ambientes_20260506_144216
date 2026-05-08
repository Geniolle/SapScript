import os
from typing import Any

from fastapi import FastAPI, Form, Header, HTTPException, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from web_api.store import claim_next_job, complete_job, create_job, get_job, init_db, list_jobs

WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")

app = FastAPI(title="SAP Script Web")
app.mount("/static", StaticFiles(directory="web_api/static"), name="static")
templates = Jinja2Templates(directory="web_api/templates")


class CompleteJobRequest(BaseModel):
    state: str
    status: str
    log: str = ""


@app.on_event("startup")
def startup() -> None:
    init_db()


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse("index.html", {"request": request})


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
