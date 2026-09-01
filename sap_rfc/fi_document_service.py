from __future__ import annotations

import json
import os
import platform
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None

from .fi_config import build_connection_params as fi_build_connection_params
from .fi_payload_builder import (
    _apply_default_payload as fi_apply_default_payload,
    _build_bapi_payload as fi_build_bapi_payload,
    _build_customer_payload as fi_build_customer_payload,
    _build_gl_payload as fi_build_gl_payload,
    _build_vendor_payload as fi_build_vendor_payload,
)


@dataclass
class FiDocumentResult:
    ok: bool
    status: str
    message: str
    branch: str
    company_code: str = ""
    document_number: str = ""
    check_return: list[dict[str, Any]] = field(default_factory=list)
    post_return: list[dict[str, Any]] = field(default_factory=list)
    commit_return: list[dict[str, Any]] = field(default_factory=list)
    payload: dict[str, Any] = field(default_factory=dict)


def _bridge_python_executable() -> Path | None:
    repo_root = Path(__file__).resolve().parents[1]
    candidates = [
        repo_root / ".venv-rfc" / "Scripts" / "python.exe",
        repo_root / ".venv-rfc" / "Scripts" / "pythonw.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _run_post_fi_document_via_bridge(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    python_exe = _bridge_python_executable()
    if python_exe is None:
        raise RuntimeError(
            "Execução FI via bridge indisponível neste runtime "
            f"({platform.platform()}). Configure SAP_FI_BRIDGE_PYTHON ou WORKFLOW_PYTHON_EXEC "
            "com um Python compatível com PyRFC, ou execute este worker no Windows "
            "com .venv-rfc\\Scripts\\python.exe disponível."
        )

    repo_root = Path(__file__).resolve().parents[1]
    bridge_script = (
        "from pathlib import Path\n"
        "import dataclasses\n"
        "import json\n"
        "import sys\n"
        f"repo_root = Path(r'{repo_root}')\n"
        "sys.path.insert(0, str(repo_root))\n"
        "from sap_rfc.fi_document_service import _post_fi_document_core as _post\n"
        "environment = json.loads(sys.stdin.readline())\n"
        "branch = json.loads(sys.stdin.readline())\n"
        "payload = json.loads(sys.stdin.read() or '{}')\n"
        "result = _post(environment, branch, payload)\n"
        "print(json.dumps(dataclasses.asdict(result), ensure_ascii=False, default=str))\n"
    )

    env = os.environ.copy()
    env["SAP_FI_BRIDGE_ACTIVE"] = "1"

    proc = subprocess.run(
        [str(python_exe), "-c", bridge_script],
        input="\n".join(
            [
                json.dumps(environment, ensure_ascii=False, default=str),
                json.dumps(branch, ensure_ascii=False, default=str),
                json.dumps(payload, ensure_ascii=False, default=str),
            ]
        ),
        capture_output=True,
        text=True,
        cwd=str(repo_root),
        env=env,
    )
    stdout = proc.stdout.strip()
    stderr = proc.stderr.strip()
    if proc.returncode != 0:
        detail = stderr or stdout or f"Processo RFC de apoio falhou com código {proc.returncode}."
        raise RuntimeError(detail)
    if not stdout:
        raise RuntimeError("Processo RFC de apoio devolveu resposta vazia.")
    data = json.loads(stdout)
    return FiDocumentResult(**data)


def _require_pyrfc() -> None:
    if Connection is None:
        raise RuntimeError(f"PyRFC indisponível: {_PYRFC_IMPORT_ERROR}")


def _check_return_tables(response: dict[str, Any]) -> list[dict[str, Any]]:
    rows = response.get("RETURN") or response.get("RETURN[]") or []
    if isinstance(rows, dict):
        rows = [rows]
    return [dict(row) for row in rows if isinstance(row, dict)]


def _has_bapi_error(rows: list[dict[str, Any]]) -> bool:
    for row in rows:
        type_ = str(row.get("TYPE") or "").strip().upper()
        if type_ in {"A", "E", "X"}:
            return True
    return False


def _join_return_messages(rows: list[dict[str, Any]]) -> str:
    parts: list[str] = []
    for row in rows:
        message = str(row.get("MESSAGE") or "").strip()
        if message:
            parts.append(message)
    return " | ".join(parts)


def _call_bapi(connection: Any, function_name: str, payload: dict[str, Any]) -> dict[str, Any]:
    response = connection.call(function_name, **payload)
    if not isinstance(response, dict):
        return {}
    return response


def _extract_document_number(response: dict[str, Any]) -> str:
    for key in ("OBJ_KEY", "BELNR", "DOC_NO", "DOCUMENTNUMBER", "DOC_NUMBER"):
        value = str(response.get(key) or "").strip()
        if value:
            return value
    return ""


def _post_fi_document_core(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    if Connection is None:
        if os.getenv("SAP_FI_BRIDGE_ACTIVE") == "1":
            raise RuntimeError(
                "PyRFC indisponível no processo de bridge. Verifique se o Python do bridge "
                "tem SAP NetWeaver RFC SDK + pyrfc instalados."
            )
        return _run_post_fi_document_via_bridge(environment, branch, payload)

    _require_pyrfc()
    connection_params = fi_build_connection_params(environment)
    payload = fi_apply_default_payload(environment, branch, payload)

    connection = Connection(**connection_params)  # type: ignore[misc]
    try:
        bapi_payload = fi_build_bapi_payload(branch, environment, payload, connection=connection)
        check_response = _call_bapi(connection, "BAPI_ACC_DOCUMENT_CHECK", bapi_payload)
        check_return = _check_return_tables(check_response)
        if _has_bapi_error(check_return):
            return FiDocumentResult(
                ok=False,
                status="ERRO",
                message=_join_return_messages(check_return) or "BAPI_ACC_DOCUMENT_CHECK devolveu erro.",
                branch=branch,
                company_code=str(payload.get("company_code") or "").strip().upper(),
                check_return=check_return,
                payload=dict(bapi_payload),
            )

        post_response = _call_bapi(connection, "BAPI_ACC_DOCUMENT_POST", bapi_payload)
        post_return = _check_return_tables(post_response)
        if _has_bapi_error(post_return):
            return FiDocumentResult(
                ok=False,
                status="ERRO",
                message=_join_return_messages(post_return) or "BAPI_ACC_DOCUMENT_POST devolveu erro.",
                branch=branch,
                company_code=str(payload.get("company_code") or "").strip().upper(),
                check_return=check_return,
                post_return=post_return,
                payload=dict(bapi_payload),
            )

        commit_response = _call_bapi(connection, "BAPI_TRANSACTION_COMMIT", {"WAIT": "X"})
        commit_return = _check_return_tables(commit_response)
        document_number = _extract_document_number(post_response) or _extract_document_number(check_response)

        message = _join_return_messages(post_return or check_return) or "Documento FI processado com sucesso."
        return FiDocumentResult(
            ok=True,
            status="SUCESSO",
            message=message,
            branch=branch,
            company_code=str(payload.get("company_code") or "").strip().upper(),
            document_number=document_number,
            check_return=check_return,
            post_return=post_return,
            commit_return=commit_return,
            payload=dict(bapi_payload),
        )
    finally:
        try:
            connection.close()
        except Exception:
            pass


def post_fi_document(environment: str, branch: str, payload: dict[str, Any]) -> FiDocumentResult:
    if Connection is None:
        return _run_post_fi_document_via_bridge(environment, branch, payload)
    return _post_fi_document_core(environment, branch, payload)


# Reexports para manter compatibilidade com imports internos antigos.
build_connection_params = fi_build_connection_params
apply_default_payload = fi_apply_default_payload
build_bapi_payload = fi_build_bapi_payload
_build_customer_payload = fi_build_customer_payload
_build_vendor_payload = fi_build_vendor_payload
_build_gl_payload = fi_build_gl_payload
