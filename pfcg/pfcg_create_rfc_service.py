from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import Any

RFC_VENV_RELATIVE_PYTHON = Path(".venv-rfc") / "Scripts" / "python.exe"
RFC_SDK_HOME = r"C:\nwrfcsdk"
RFC_ALLOWED_EXIT_CODES = {0, 1}
RFC_TIMEOUT_SECONDS = 120


class PfcgCreateRfcBridgeError(Exception):
    pass


def _find_project_root() -> Path:
    # `_rfc_common` não importa pyrfc a nível de módulo, por isso pode ser usado
    # com segurança a partir da venv "normal" (não é preciso a .venv-rfc aqui).
    from sap_rfc._rfc_common import find_project_root

    return find_project_root()


def _build_bridge_env(project_dir: Path) -> dict[str, str]:
    env = os.environ.copy()
    sdk_home = (
        str(env.get("SAP_NWRFC_HOME") or "").strip()
        or str(env.get("SAPNWRFC_HOME") or "").strip()
        or RFC_SDK_HOME
    )
    env["SAP_SCRIPT_PROJECT_DIR"] = str(project_dir)
    env["SAPNWRFC_HOME"] = sdk_home
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    sdk_lib = str(Path(sdk_home) / "lib")
    current_path = env.get("PATH", "")
    if sdk_lib.lower() not in current_path.lower():
        path_entries = [sdk_lib]
        if current_path:
            path_entries.append(current_path)
        env["PATH"] = os.pathsep.join(path_entries)
    return env


def _run_bridge_cli(cli_module: str, args: list[str], *, timeout: int = RFC_TIMEOUT_SECONDS) -> dict[str, Any]:
    project_dir = _find_project_root()
    rfc_python = (project_dir / RFC_VENV_RELATIVE_PYTHON).resolve()
    if not rfc_python.exists():
        raise PfcgCreateRfcBridgeError(f"Python RFC não encontrado: {rfc_python}")

    env = _build_bridge_env(project_dir)
    command = [str(rfc_python), "-m", cli_module, *args]

    try:
        run = subprocess.run(
            command,
            cwd=str(project_dir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
            shell=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PfcgCreateRfcBridgeError(f"Timeout ao executar {cli_module}.") from exc

    stdout = str(run.stdout or "").strip()
    stderr = str(run.stderr or "").strip()

    if run.returncode not in RFC_ALLOWED_EXIT_CODES:
        raise PfcgCreateRfcBridgeError(
            f"Bridge RFC ({cli_module}) falhou com exit code {run.returncode}. stderr={stderr[:500]}"
        )
    if not stdout:
        raise PfcgCreateRfcBridgeError(
            f"Bridge RFC ({cli_module}) não devolveu JSON em stdout. stderr={stderr[:500]}"
        )

    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise PfcgCreateRfcBridgeError(f"Bridge RFC ({cli_module}) devolveu JSON inválido.") from exc

    if not isinstance(payload, dict):
        raise PfcgCreateRfcBridgeError(f"Bridge RFC ({cli_module}) devolveu payload inválido.")

    return payload


def _tcode_args(tcodes: list[str]) -> list[str]:
    args: list[str] = []
    for tcode in tcodes:
        args.extend(["--tcode", str(tcode)])
    return args


def preview_pfcg_role_create_rfc(
    environment: str,
    role_name: str,
    description: str,
    tcodes: list[str],
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Bridge de leitura (preview) para `sap_rfc.pfcg_role_create_service.preview_pfcg_role_create`.

    Executa no subprocesso isolado `.venv-rfc`; nunca escreve em SAP.
    """
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        "--role-name",
        str(role_name or "").strip(),
        "--description",
        str(description or ""),
        *_tcode_args(tcodes),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
    ]
    return _run_bridge_cli("sap_rfc.pfcg_role_create_preview_cli", args)


def create_pfcg_role_rfc(
    environment: str,
    role_name: str,
    description: str,
    tcodes: list[str],
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    """Bridge de ESCRITA para `sap_rfc.pfcg_role_create_service.create_pfcg_role_rfc`.

    Único ponto de entrada de escrita RFC para a criação individual de função PFCG,
    usado tanto pelo fluxo Web (worker) como pelo modo `metodo="RFC"` de
    `Processos/Funções PFCG/A. PFCG_CREATE.py`, para não duplicar a lógica de
    chamada SAP em mais do que um sítio.
    """
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        "--role-name",
        str(role_name or "").strip(),
        "--description",
        str(description or ""),
        *_tcode_args(tcodes),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
        "--confirm",
    ]
    return _run_bridge_cli("sap_rfc.pfcg_role_create_cli", args)


def search_transport_requests_rfc(environment: str) -> dict[str, Any]:
    """Bridge de leitura para `sap_rfc.pfcg_transport_service.search_open_transport_requests`.

    Executa no subprocesso isolado `.venv-rfc`; nunca escreve em SAP.
    """
    args = ["--environment", str(environment or "").strip().upper()]
    return _run_bridge_cli("sap_rfc.pfcg_transport_search_cli", args)


