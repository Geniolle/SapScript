from __future__ import annotations

import os
import socket
import sys
import time
import traceback
from pathlib import Path
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def _bootstrap_import_paths() -> None:
    worker_dir = Path(__file__).resolve().parent
    cockpit_dir = worker_dir.parent
    repo_root = worker_dir.parent.parent

    candidate_paths = []
    sap_script_project_dir = os.getenv("SAP_SCRIPT_PROJECT_DIR", "").strip()
    if sap_script_project_dir:
        candidate_paths.append(sap_script_project_dir)
    candidate_paths.append(str(repo_root))
    candidate_paths.append(str(cockpit_dir))

    existing_pythonpath = os.environ.get("PYTHONPATH", "").strip()
    pythonpath_parts = [part for part in existing_pythonpath.split(os.pathsep) if part]

    for path in candidate_paths:
        if path not in sys.path:
            sys.path.insert(0, path)
        if path not in pythonpath_parts:
            pythonpath_parts.insert(0, path)

    os.environ["PYTHONPATH"] = os.pathsep.join(pythonpath_parts)


_bootstrap_import_paths()

from sap_tasks import run_sap_task, JobCancelledException

_REQUEST_SESSION = requests.Session()
_REQUEST_SESSION.headers.update({"Connection": "keep-alive"})
_REQUEST_SESSION.mount(
    "http://",
    HTTPAdapter(
        max_retries=Retry(
            total=2,
            connect=2,
            read=2,
            status=2,
            backoff_factor=0.25,
            allowed_methods=frozenset({"GET", "POST"}),
        )
    ),
)
_REQUEST_SESSION.mount(
    "https://",
    HTTPAdapter(
        max_retries=Retry(
            total=2,
            connect=2,
            read=2,
            status=2,
            backoff_factor=0.25,
            allowed_methods=frozenset({"GET", "POST"}),
        )
    ),
)

API_CONNECT_TIMEOUT = float(os.getenv("WORKER_API_CONNECT_TIMEOUT", "3"))
API_READ_TIMEOUT = float(os.getenv("WORKER_API_READ_TIMEOUT", "15"))
JOB_POLL_INTERVAL_SECONDS = int(os.getenv("POLL_SECONDS", "3"))


def _resolve_api_base_url() -> str:
    configured = os.getenv("API_BASE_URL", "").strip().rstrip("/")
    if configured:
        return configured
    return "http://localhost:8010"


API_BASE_URL = _resolve_api_base_url()
WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
WORKER_NAME = os.getenv("WORKER_NAME", socket.gethostname())


def headers() -> dict[str, str]:
    return {"X-Worker-Token": WORKER_TOKEN}


def claim_next_job() -> dict[str, Any] | None:
    response = _REQUEST_SESSION.get(
        f"{API_BASE_URL}/api/worker/jobs/next",
        params={"worker_name": WORKER_NAME},
        headers=headers(),
        timeout=(API_CONNECT_TIMEOUT, API_READ_TIMEOUT),
    )
    response.raise_for_status()
    return response.json().get("job")


def complete_job(job_id: str, state: str, status: str, log: str) -> None:
    response = _REQUEST_SESSION.post(
        f"{API_BASE_URL}/api/jobs/{job_id}/complete",
        headers=headers(),
        json={"state": state, "status": status, "log": log},
        timeout=(API_CONNECT_TIMEOUT, API_READ_TIMEOUT),
    )
    response.raise_for_status()


def process_job(job: dict[str, Any]) -> None:
    try:
        status, log = run_sap_task(job)
        complete_job(job["id"], "succeeded", status, log)
    except JobCancelledException:
        print(f"❌ Job {job['id']} interrompido e cancelado com sucesso no SAP.")
    except BaseException as exc:
        status = str(exc) or "Erro sem mensagem (ou sys.exit)"
        log = traceback.format_exc()
        complete_job(job["id"], "failed", status, log)


def main() -> None:
    api_source = "API_BASE_URL" if os.getenv("API_BASE_URL", "").strip() else "default fallback"
    print(f"Worker {WORKER_NAME} iniciado.")
    print(f"API ativa: {API_BASE_URL} ({api_source})")
    print(f"Timeout API: connect={API_CONNECT_TIMEOUT}s read={API_READ_TIMEOUT}s")
    print("Para terminar, usa CTRL+C.")
    while True:
        try:
            job = claim_next_job()
            if job:
                print(f"A executar job {job['id']} ({job['task']})")
                process_job(job)
            else:
                time.sleep(JOB_POLL_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            print("Worker terminado pelo utilizador.")
            break
        except requests.RequestException as exc:
            print(f"[WORKER API] Falha de ligação a {API_BASE_URL}: {exc}")
            time.sleep(min(JOB_POLL_INTERVAL_SECONDS, 5))
        except BaseException:
            print(traceback.format_exc())
            time.sleep(JOB_POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()

