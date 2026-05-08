from __future__ import annotations

import os
import socket
import time
import traceback
from typing import Any

import requests

from sap_tasks import run_sap_task

API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:8000").rstrip("/")
WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
WORKER_NAME = os.getenv("WORKER_NAME", socket.gethostname())
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "3"))


def headers() -> dict[str, str]:
    return {"X-Worker-Token": WORKER_TOKEN}


def claim_next_job() -> dict[str, Any] | None:
    response = requests.get(
        f"{API_BASE_URL}/api/jobs/next",
        params={"worker_name": WORKER_NAME},
        headers=headers(),
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("job")


def complete_job(job_id: str, state: str, status: str, log: str) -> None:
    response = requests.post(
        f"{API_BASE_URL}/api/jobs/{job_id}/complete",
        headers=headers(),
        json={"state": state, "status": status, "log": log},
        timeout=30,
    )
    response.raise_for_status()


def process_job(job: dict[str, Any]) -> None:
    try:
        status, log = run_sap_task(job)
        complete_job(job["id"], "succeeded", status, log)
    except Exception as exc:
        status = str(exc) or "Erro sem mensagem"
        log = traceback.format_exc()
        complete_job(job["id"], "failed", status, log)


def main() -> None:
    print(f"Worker {WORKER_NAME} ligado a {API_BASE_URL}")
    print("Para terminar, usa CTRL+C.")
    while True:
        try:
            job = claim_next_job()
            if job:
                print(f"A executar job {job['id']} ({job['task']})")
                process_job(job)
            else:
                time.sleep(POLL_SECONDS)
        except KeyboardInterrupt:
            print("Worker terminado pelo utilizador.")
            break
        except Exception:
            print(traceback.format_exc())
            time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
