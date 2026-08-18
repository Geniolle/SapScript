import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import os
import socket
import time
import traceback
import threading
from typing import Any

import requests

from sap_tasks import run_sap_task, JobCancelledException

API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:8010").rstrip("/")
WORKER_TOKEN = os.getenv("WORKER_TOKEN", "change-me")
WORKER_NAME = os.getenv("WORKER_NAME", socket.gethostname())
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "3"))

# Thread stop event
stop_event = threading.Event()

def heartbeat_loop() -> None:
    heartbeat_seconds = float(os.getenv("HEARTBEAT_SECONDS", "5.0"))
    error_cooldown = 60.0
    debug_mode = os.getenv("WORKER_HEARTBEAT_DEBUG", "false").strip().lower() in ("1", "true", "yes", "sim")
    
    consecutive_failures = 0
    last_error_time = 0.0
    
    with requests.Session() as session:
        while not stop_event.is_set():
            try:
                response = session.post(
                    f"{API_BASE_URL}/api/worker/heartbeat",
                    headers=headers(),
                    json={"worker_name": WORKER_NAME},
                    timeout=(2.0, 8.0)
                )
                response.raise_for_status()
                
                if consecutive_failures > 0:
                    print("❇️ Conexão de heartbeat restabelecida com a API.")
                
                consecutive_failures = 0
                
                if debug_mode:
                    print("[DEBUG HEARTBEAT CLIENT] Heartbeat enviado com sucesso.")
            except Exception as e:
                consecutive_failures += 1
                current_time = time.time()
                
                if consecutive_failures == 1:
                    print(f"⚠️ Falha no heartbeat do worker: {e}")
                    last_error_time = current_time
                elif current_time - last_error_time > error_cooldown:
                    print(f"⚠️ Falha contínua no heartbeat do worker (Falhas consecutivas: {consecutive_failures}): {e}")
                    last_error_time = current_time
                
                if debug_mode:
                    print(f"[DEBUG HEARTBEAT CLIENT] Falha no heartbeat. Total consecutivas: {consecutive_failures}. Erro: {e}")
                    
            stop_event.wait(heartbeat_seconds)


def headers() -> dict[str, str]:
    return {"X-Worker-Token": WORKER_TOKEN}


def claim_next_job() -> dict[str, Any] | None:
    response = requests.get(
        f"{API_BASE_URL}/api/worker/jobs/next",
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
    status = "Execução falhou"
    log = ""
    state = "failed"
    try:
        import importlib
        import sap_tasks
        try:
            importlib.reload(sap_tasks)
        except Exception as reload_err:
            print(f"[WORKER] Aviso ao recarregar sap_tasks: {reload_err}")

        run_res = sap_tasks.run_sap_task(job)
        if isinstance(run_res, tuple) and len(run_res) == 3:
            status, log, success_val = run_res
        else:
            status, log = run_res
            # Fallback seguro não-textual de substring genérica
            # Analisa apenas início das linhas para erros reais
            success_val = True
            for line in log.splitlines():
                line_clean = line.strip()
                if line_clean.startswith(("[ERROR]", "[ERRO]", "❌", "Traceback (most recent call last):")):
                    success_val = False
                    break

        has_warnings = "[DOC_WARN]" in log or "[TECHNICAL WARN]" in log or "warning" in status.lower()
        if not success_val:
            state = "failed"
        elif has_warnings:
            state = "succeeded_with_warnings"
        else:
            state = "succeeded"
    except JobCancelledException:
        print(f"❌ Job {job['id']} interrompido e cancelado com sucesso no SAP.")
        status = "Cancelado pelo utilizador"
        log = "O pedido foi cancelado manualmente via interface web."
        state = "failed"
    except BaseException as exc:
        status = str(exc) or "Erro sem mensagem (ou sys.exit)"
        log = traceback.format_exc()
        state = "failed"
    finally:
        try:
            complete_job(job["id"], state, status, log)
        except Exception as e:
            print(f"Erro ao completar job {job['id']}: {e}")


def main() -> None:
    print(f"Worker {WORKER_NAME} ligado a {API_BASE_URL}")
    print("Para terminar, usa CTRL+C.")
    
    # Iniciar heartbeat thread
    heartbeat_thread = threading.Thread(target=heartbeat_loop, daemon=True)
    heartbeat_thread.start()
    
    try:
        while True:
            try:
                job = claim_next_job()
                if job:
                    print(f"A executar job {job['id']} ({job['task']})")
                    process_job(job)
                else:
                    time.sleep(POLL_SECONDS)
            except KeyboardInterrupt:
                raise
            except BaseException:
                print(traceback.format_exc())
                time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        print("Worker terminado pelo utilizador.")
    finally:
        stop_event.set()
        heartbeat_thread.join(timeout=1.0)


if __name__ == "__main__":
    main()

