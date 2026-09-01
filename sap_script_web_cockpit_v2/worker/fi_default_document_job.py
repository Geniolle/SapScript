from __future__ import annotations

import dataclasses
import json
import os
from typing import Any

import requests


class FiDefaultDocumentJobError(RuntimeError):
    pass


def _json_safe(value: Any) -> Any:
    if hasattr(value, "isoformat") and callable(getattr(value, "isoformat")):
        try:
            return value.isoformat()
        except Exception:
            pass
    if isinstance(value, dict):
        return {key: _json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_json_safe(item) for item in value]
    return value


def build_fi_default_document_payload(params: dict[str, Any]) -> tuple[str, str, dict[str, Any]]:
    environment = str(params.get("environment") or "QAD").strip().upper()
    branch = str(params.get("branch") or "cliente").strip().lower()
    fi_payload = dict(params.get("payload") or {"data_mode": "default"})
    fi_payload.setdefault("data_mode", "default")
    fi_payload.setdefault("environment", environment)
    fi_payload.setdefault("branch", branch)
    return environment, branch, fi_payload


def update_job_params_via_api(job_id: str, new_params: dict[str, Any]) -> dict[str, Any]:
    api_base_url = (
        os.getenv("SAP_API_BASE_URL", "").strip().rstrip("/")
        or os.getenv("API_BASE_URL", "").strip().rstrip("/")
    )
    if not api_base_url:
        raise FiDefaultDocumentJobError("API base URL não definido para atualizar o job.")

    worker_token = os.getenv("WORKER_TOKEN", "change-me")
    response = requests.post(
        f"{api_base_url}/api/jobs/{job_id}/params",
        headers={"X-Worker-Token": worker_token},
        json={"params": new_params},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def run_fi_default_document_job(
    *,
    job_id: str,
    params: dict[str, Any],
    post_fi_document: Any,
) -> tuple[str, str]:
    environment, branch, fi_payload = build_fi_default_document_payload(params)
    result = post_fi_document(environment, branch, fi_payload)
    result_payload = _json_safe(dataclasses.asdict(result))
    result_json = json.dumps(result_payload, ensure_ascii=False)

    try:
        update_job_params_via_api(job_id, {"fi_document_result": result_payload})
    except Exception as exc:
        raise FiDefaultDocumentJobError(
            f"Não foi possível gravar o resultado FI no job: {exc}"
        ) from exc

    log = (
        "Documento FI executado pelo worker Windows.\n"
        f"Ambiente: {environment}\n"
        f"Branch: {branch}\n"
        f"Status: {result.status}\n"
        f"Mensagem: {result.message}"
    )
    return result_json, log
