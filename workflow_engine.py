import json
import logging
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

from sap_session import load_dotenv_manual
from workflow_documentation import WorkflowDocumentation


REQUEST_REGEX = re.compile(r"\b([A-Z0-9]{3,4}K\d{6,})\b")
BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}


class SafeFormatDict(dict):
    def __missing__(self, key: str) -> str:
        return "{" + key + "}"


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in BOOL_TRUE


def _step_capture_evidence_enabled(step: Dict[str, Any]) -> bool:
    if "capture_evidence" in step:
        return _to_bool(str(step.get("capture_evidence", "")))
    if "capture_runtime_snapshot" in step:
        return _to_bool(str(step.get("capture_runtime_snapshot", "")))
    return False


def _pause_before_step_if_enabled(
    *,
    workflow_name: str,
    step_name: str,
    index: int,
    total: int,
    row_context: Dict[str, str],
) -> None:
    if not _to_bool(os.getenv("WORKFLOW_STEP_CONFIRM", "false")):
        return

    ticket_key = str(row_context.get("ticket_key", "") or "").strip() or "-"
    categoria = str(row_context.get("categoria_sap", "") or "").strip() or workflow_name
    request_number = str(row_context.get("request_number", "") or "").strip() or "-"

    message = (
        "\n"
        "================================================================================\n"
        "PAUSA DE VALIDAÇÃO DO WORKFLOW\n"
        f"Ticket: {ticket_key}\n"
        f"Categoria: {categoria}\n"
        f"Workflow: {workflow_name}\n"
        f"Step: {index}/{total} - {step_name}\n"
        f"Request atual: {request_number}\n"
        "Pressiona ENTER para executar este step, ou CTRL+C para interromper.\n"
        "================================================================================\n"
    )

    if not sys.stdin or not sys.stdin.isatty():
        logging.warning(
            "WORKFLOW_STEP_CONFIRM ativo, mas stdin nao e interativo; pausa ignorada para o step '%s'.",
            step_name,
        )
        return

    print(message, flush=True)
    input()


def _load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with open(path, "r", encoding="utf-8-sig") as file_obj:
        return json.load(file_obj)


def _save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as file_obj:
        json.dump(payload, file_obj, ensure_ascii=False, indent=2)


def _format_value(template: str, context: Dict[str, str]) -> str:
    return str(template).format_map(SafeFormatDict(context)).strip()


def _find_ticket_xlsx(download_dir: Path, ticket_key: str) -> str:
    ticket_folder = download_dir / ticket_key
    if not ticket_folder.exists() or not ticket_folder.is_dir():
        return ""

    files = sorted(
        ticket_folder.glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return str(files[0].resolve()) if files else ""


def _normalize_col_key(name: str) -> str:
    normalized = re.sub(r"[^a-zA-Z0-9]+", "_", str(name or "").strip().lower())
    normalized = normalized.strip("_")
    return normalized or "field"


def _build_row_context(
    row: Dict[str, Any],
    *,
    ambiente: str,
    system_name: str,
    sap_client: str,
    sap_connection_name: str,
    download_dir: Path,
) -> Dict[str, str]:
    dados = row.get("dados", {}) or {}
    ticket_key = str(dados.get("Chave", "")).strip().upper()
    categoria = str(dados.get("IT SALSA - Categoria SAP", "")).strip()
    updated = str(dados.get("Atualizado", "")).strip()
    estado = str(dados.get("Estado", "")).strip()
    resumo = str(dados.get("Resumo", "")).strip()
    request_description = " | ".join(part for part in (ticket_key, resumo) if part).strip()

    context: Dict[str, str] = {
        "ticket_key": ticket_key,
        "categoria_sap": categoria,
        "numero_linha": str(row.get("numero_linha", "")),
        "atualizado": updated,
        "estado": estado,
        "resumo": resumo,
        "request_description": request_description,
        "ambiente": ambiente,
        "system_name": system_name,
        "sap_client": sap_client,
        "sap_connection_name": sap_connection_name,
        "request_number": "",
        "xlsx_path": _find_ticket_xlsx(download_dir, ticket_key),
    }

    for raw_key, raw_value in dados.items():
        context[f"col_{_normalize_col_key(raw_key)}"] = str(raw_value or "").strip()

    return context


def _parse_request_number(output: str) -> str:
    marker = re.search(r"REQUEST_NUMBER=([A-Z0-9]{3,4}K\d{6,})", output or "")
    if marker:
        return marker.group(1).strip().upper()

    fallback = REQUEST_REGEX.search(output or "")
    return fallback.group(1).strip().upper() if fallback else ""


def _resolve_sap_runtime_context() -> Dict[str, str]:
    workflow_key = os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100").strip().upper()
    system_name = os.getenv("WORKFLOW_SAP_SYSTEM", "").strip().upper()
    ambiente = os.getenv("WORKFLOW_AMBIENTE", "").strip().upper()
    client = os.getenv("WORKFLOW_SAP_CLIENT", "").strip()

    if not system_name and workflow_key:
        if "CLNT" in workflow_key:
            system_name = workflow_key.split("CLNT", 1)[0].strip().upper()
        else:
            system_name = workflow_key

    if not client and workflow_key:
        client = os.getenv(f"SAP_CLIENT_{workflow_key}", "").strip()
    if not client:
        client = os.getenv("SAP_CLIENT", "").strip()

    if not ambiente:
        map_ambiente = {
            "S4D": "DEV",
            "S4Q": "QAD",
            "S4P": "PRD",
            "SPA": "CUA",
        }
        ambiente = map_ambiente.get(system_name, "QAD")

    connection_name = ""
    if workflow_key:
        connection_name = os.getenv(f"SAP_CONNECTION_{workflow_key}", "").strip()

    return {
        "workflow_sap_key": workflow_key,
        "system_name": system_name,
        "ambiente": ambiente,
        "sap_client": client,
        "sap_connection_name": connection_name,
    }


def _run_step(
    *,
    step: Dict[str, Any],
    step_name: str,
    context: Dict[str, str],
    base_dir: Path,
    python_exec: str,
    documentation: WorkflowDocumentation | None = None,
) -> Tuple[bool, str, Dict[str, str]]:
    script_template = str(step.get("script", "")).strip()
    if not script_template:
        return False, "Step sem campo 'script'.", {}

    required_context = step.get("required_context", []) or []
    missing = [name for name in required_context if not str(context.get(name, "")).strip()]
    if missing:
        return False, f"Contexto em falta: {', '.join(missing)}", {}

    rendered_script = _format_value(script_template, context)
    script_path = Path(rendered_script)
    if not script_path.is_absolute():
        script_path = (base_dir / script_path).resolve()

    if not script_path.exists():
        return False, f"Script nao encontrado: {script_path}", {}

    args = []
    for raw_arg in step.get("args", []) or []:
        value = _format_value(str(raw_arg), context)
        if value:
            args.append(value)

    capture_evidence = _step_capture_evidence_enabled(step)

    command = [python_exec, str(script_path)] + args
    logging.info("Workflow step: %s", " ".join(command))

    step_env = os.environ.copy()
    # Steps disparados daqui pertencem ao fluxo do main/workflow.
    step_env["SAP_CALLED_BY_MAIN"] = "1"
    # Evita erro de encoding em stdout quando scripts imprimem caracteres especiais.
    step_env.setdefault("PYTHONIOENCODING", "utf-8")
    # Sinaliza ao script se deve manter a tela de validacao para captura de evidencia.
    step_env["WORKFLOW_CAPTURE_VALIDATION_SCREEN"] = "1" if capture_evidence else "0"

    run = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=step_env,
        check=False,
        shell=False,
    )
    runtime_snapshot: Dict[str, str] = {}
    stdout = (run.stdout or "").strip()
    stderr = (run.stderr or "").strip()

    if stdout:
        logging.info("Step stdout:\n%s", stdout)
    if stderr:
        logging.warning("Step stderr:\n%s", stderr)

    if step.get("capture_request_number", False):
        req = _parse_request_number(f"{stdout}\n{stderr}")
        if req:
            context["request_number"] = req
            logging.info("Request capturada: %s", req)
        elif step.get("require_request_number", False):
            return False, "Nao foi possivel extrair REQUEST_NUMBER do step.", runtime_snapshot
        else:
            logging.info("Step sem REQUEST_NUMBER; contexto de request mantido.")

    if documentation and documentation.enabled and capture_evidence:
        # Regra de sincronizacao: so avanca para o proximo step apos confirmar
        # a captura final deste step.
        runtime_snapshot = documentation.capture_runtime_snapshot_with_retry(
            step_name=step_name,
            row_context=context,
            attempts=4,
            wait_s=0.4,
        )

    if run.returncode != 0:
        return False, f"Step retornou codigo {run.returncode}", runtime_snapshot

    return True, "", runtime_snapshot


def _run_workflow(
    *,
    workflow_name: str,
    workflow: Dict[str, Any],
    row_context: Dict[str, str],
    base_dir: Path,
    python_exec: str,
    documentation: WorkflowDocumentation | None = None,
) -> Tuple[bool, str]:
    steps = workflow.get("steps", []) or []
    if not steps:
        return False, f"Workflow '{workflow_name}' sem steps."

    for index, step in enumerate(steps, start=1):
        step_name = str(step.get("name", f"step_{index}"))
        capture_evidence = _step_capture_evidence_enabled(step)
        logging.info("Workflow '%s' | Step %s/%s: %s", workflow_name, index, len(steps), step_name)

        _pause_before_step_if_enabled(
            workflow_name=workflow_name,
            step_name=step_name,
            index=index,
            total=len(steps),
            row_context=row_context,
        )

        ok, error, runtime_snapshot = _run_step(
            step=step,
            step_name=step_name,
            context=row_context,
            base_dir=base_dir,
            python_exec=python_exec,
            documentation=documentation,
        )
        if documentation:
            documentation.capture_step(
                step_name=step_name,
                row_context=row_context,
                note="" if ok else f"Falha no step: {error}",
                snapshot_override=runtime_snapshot,
                allow_live_capture=capture_evidence,
            )
        if not ok:
            return False, f"Falha no step '{step_name}': {error}"

    return True, ""


def execute_workflows(rows: List[Dict[str, Any]], base_dir: Path) -> None:
    load_dotenv_manual()

    enabled = _to_bool(os.getenv("WORKFLOW_ENABLED", "true"))
    if not enabled:
        logging.info("WORKFLOW_ENABLED desativado.")
        return

    config_path = Path(
        os.getenv("WORKFLOW_CONFIG_PATH", str((base_dir / "workflows.json").resolve()))
    ).resolve()
    state_path = Path(
        os.getenv("WORKFLOW_STATE_PATH", str((base_dir / "cache" / "workflow_state.json").resolve()))
    ).resolve()
    python_exec = os.getenv("WORKFLOW_PYTHON_EXEC", sys.executable)
    download_dir = Path(os.getenv("JIRA_DOWNLOAD_DIR", r"C:\Jira")).resolve()
    sap_ctx = _resolve_sap_runtime_context()

    logging.info(
        "Contexto SAP workflow | Key=%s | Sistema=%s | Ambiente=%s | Mandante=%s",
        sap_ctx["workflow_sap_key"],
        sap_ctx["system_name"] or "-",
        sap_ctx["ambiente"],
        sap_ctx["sap_client"] or "-",
    )

    workflows = _load_json(config_path, {})
    if not isinstance(workflows, dict):
        raise RuntimeError(f"Formato invalido em {config_path}")

    state = _load_json(state_path, {"processed": {}})
    processed = state.get("processed", {})
    if not isinstance(processed, dict):
        processed = {}
        state["processed"] = processed

    changed = False

    for row in rows:
        dados = row.get("dados", {}) or {}
        category = str(dados.get("IT SALSA - Categoria SAP", "")).strip()
        ticket_key = str(dados.get("Chave", "")).strip().upper()
        if not category or not ticket_key:
            continue

        workflow = workflows.get(category)
        if not workflow:
            continue

        updated = str(dados.get("Atualizado", "")).strip()
        state_id = f"{ticket_key}|{category}|{updated}"

        previous = processed.get(state_id, {})
        if previous.get("status") == "success":
            logging.info("Workflow ja concluido para %s", state_id)
            continue

        context = _build_row_context(
            row,
            ambiente=sap_ctx["ambiente"],
            system_name=sap_ctx["system_name"],
            sap_client=sap_ctx["sap_client"],
            sap_connection_name=sap_ctx["sap_connection_name"],
            download_dir=download_dir,
        )
        documentation = WorkflowDocumentation.from_env(
            base_dir=base_dir,
            row_context=context,
            workflow_name=category,
        )
        logging.info(
            "Iniciando workflow | Ticket=%s | Categoria=%s | XLSX=%s",
            context["ticket_key"],
            context["categoria_sap"],
            context["xlsx_path"] or "(nao encontrado)",
        )

        ok, error = _run_workflow(
            workflow_name=category,
            workflow=workflow,
            row_context=context,
            base_dir=base_dir,
            python_exec=python_exec,
            documentation=documentation,
        )
        doc_path = ""
        if documentation:
            doc_path = documentation.finalize(
                row_context=context,
                success=ok,
                error=error,
            )
            if doc_path:
                logging.info("Documento de evidencias gerado: %s", doc_path)

        now = datetime.now().isoformat(timespec="seconds")
        attempts = int(previous.get("attempts", 0)) + 1
        processed[state_id] = {
            "status": "success" if ok else "failed",
            "attempts": attempts,
            "timestamp": now,
            "ticket_key": ticket_key,
            "category": category,
            "request_number": context.get("request_number", ""),
            "documentation_path": doc_path,
            "error": "" if ok else error,
        }
        changed = True

        if ok:
            logging.info("Workflow concluido com sucesso para %s", state_id)
        else:
            logging.error("Workflow falhou para %s | %s", state_id, error)

    if changed:
        state["processed"] = processed
        _save_json(state_path, state)
