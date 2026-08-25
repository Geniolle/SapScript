from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Callable

WORKER_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = WORKER_DIR.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(WORKER_DIR) not in sys.path:
    sys.path.insert(0, str(WORKER_DIR))

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

ENV_PATH = PROJECT_ROOT / ".env"
if load_dotenv is not None and ENV_PATH.exists():
    load_dotenv(ENV_PATH, override=True)

try:
    from .authorization_rfc_analysis import analyze_user_authorizations_rfc
    from .authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from .authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        format_sap_date_display,
        deduplicate_roles,
        normalize_sap_user,
    )
except ImportError:
    from authorization_rfc_analysis import analyze_user_authorizations_rfc
    from authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        format_sap_date_display,
        deduplicate_roles,
        normalize_sap_user,
    )


def prompt_inputs() -> tuple[str, str]:
    target_user = input("Utilizador SAP a analisar: ").strip().upper()
    while not target_user:
        target_user = input("Utilizador SAP a analisar: ").strip().upper()

    target_system = input("Sistema alvo [default: DEV]: ").strip().upper() or "DEV"
    return target_user, target_system


def print_progress(message: str) -> None:
    print(message, flush=True)


def _emit(progress_logger: Callable[[str], None] | None, message: str) -> None:
    if callable(progress_logger):
        progress_logger(message)
    else:
        print(message, flush=True)


def _read_agr_users_rows(connection: Any, target_user: str) -> tuple[list[dict[str, str]], str]:
    attempts = [
        ("UNAME", "UNAME"),
        ("BNAME", "BNAME"),
    ]
    for _, field in attempts:
        rows = _read_rfc_table(
            connection,
            "AGR_USERS",
            ["AGR_NAME", field, "FROM_DAT", "TO_DAT"],
            [{"field": field, "value": target_user}],
            max_rows=5000,
        )
        rows = [
            row for row in rows
            if str(row.get(field) or "").strip().upper() == target_user
        ]
        if rows:
            return rows, field
    return [], "UNAME"


def _collect_tcodes_for_role(connection: Any, role_name: str) -> list[str]:
    rows = _read_rfc_table(
        connection,
        "AGR_TCODES",
        ["AGR_NAME", "TCODE"],
        [{"field": "AGR_NAME", "value": role_name}],
        max_rows=5000,
    )
    tcodes: list[str] = []
    for row in rows:
        tcode = str(row.get("TCODE") or "").strip().upper()
        if tcode and tcode not in tcodes:
            tcodes.append(tcode)
    return tcodes


def _split_system_key(system_key: str) -> tuple[str, str]:
    clean_key = str(system_key or "").strip().upper()
    if "CLNT" in clean_key:
        system, client = clean_key.split("CLNT", 1)
        return system, client
    return clean_key, ""


def _run_dev_role_analysis(
    target_user: str,
    target_system_key: str,
    progress_logger: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    target_user = normalize_sap_user(target_user)
    target_system_key = target_system_key.strip().upper() or "S4DCLNT100"
    system_name, system_client = _split_system_key(target_system_key)

    _emit(progress_logger, f"[DEV RFC] A abrir ligação RFC ao sistema {target_system_key}...")
    connection = _open_rfc_connection(target_system_key)
    _emit(progress_logger, "[DEV RFC] Ligação RFC estabelecida. Vou consultar AGR_USERS...")

    rows_roles, user_field = _read_agr_users_rows(connection, target_user)

    executed_queries: list[dict[str, Any]] = [
        {
            "table": "AGR_USERS",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_roles),
        }
    ]

    if not rows_roles:
        _emit(progress_logger, f"[DEV RFC] Nenhuma role encontrada para {target_user} em AGR_USERS usando {user_field}.")
        return {
            "success": True,
            "code": "user_not_assigned_to_system",
            "message": f"O utilizador {target_user} não tem roles devolvidas em AGR_USERS para o sistema DEV.",
            "execution_mode": "RFC",
            "execution_system": {
                "key": target_system_key,
                "system": system_name,
                "client": system_client,
            },
            "target_system": {
                "key": target_system_key,
                "system": system_name,
                "client": system_client,
            },
            "target_user": target_user,
            "user_assigned_to_system": False,
            "roles": [],
            "profiles": [],
            "functions": [],
            "summary": {"total_roles": 0, "total_profiles": 0},
            "queries": executed_queries,
            "data_source_verified": True,
            "worker_feature_version": "authorization-tables-v1",
        }

    today_str = "2026-07-23"
    raw_roles: list[dict[str, Any]] = []
    for row in rows_roles:
        role_name = str(row.get("AGR_NAME") or "").strip()
        if not role_name:
            continue
        valid_from_raw = str(row.get("FROM_DAT") or "").strip()
        valid_to_raw = str(row.get("TO_DAT") or "").strip()
        origin = classify_assignment_origin("")
        raw_roles.append({
            "role": role_name,
            "description": "",
            "subsystem": target_system_key,
            "valid_from": format_sap_date_display(valid_from_raw),
            "valid_to": format_sap_date_display(valid_to_raw),
            "validity_status": classify_validity(valid_from_raw, valid_to_raw, today_str),
            "assignment_origin": origin.get("origin", "direct"),
            "assignment_origin_label": origin.get("origin_label", "Direta"),
            "assignment_origin_code": "",
        })

    roles = deduplicate_roles(raw_roles)
    _emit(progress_logger, f"[DEV RFC] Encontradas {len(roles)} roles. Vou consultar AGR_TCODES...")

    role_functions: dict[str, list[str]] = {}
    for role in roles:
        role_name = role["role"]
        tcodes = _collect_tcodes_for_role(connection, role_name)
        role_functions[role_name] = tcodes
        role["functions"] = tcodes

    executed_queries.append({
        "table": "AGR_TCODES",
        "executed": True,
        "filters_applied": True,
        "row_count": sum(len(v) for v in role_functions.values()),
    })

    functions: list[dict[str, str]] = []
    seen_functions: set[str] = set()
    for tcodes in role_functions.values():
        for tcode in tcodes:
            if tcode in seen_functions:
                continue
            seen_functions.add(tcode)
            functions.append({"tcode": tcode})

    summary = build_authorization_summary(roles, [])

    _emit(progress_logger, "[DEV RFC] Análise DEV concluída.")

    return {
        "success": True,
        "code": "analysis_complete",
        "message": f"Análise DEV concluída para {target_user}.",
        "execution_mode": "RFC",
        "execution_system": {
            "key": target_system_key,
            "system": system_name,
            "client": system_client,
        },
        "target_system": {
            "key": target_system_key,
            "system": system_name,
            "client": system_client,
        },
        "target_user": target_user,
        "user_assigned_to_system": True,
        "roles": roles,
        "profiles": [],
        "functions": functions,
        "role_functions": role_functions,
        "summary": summary,
        "queries": executed_queries,
        "data_source_verified": True,
        "worker_feature_version": "authorization-tables-v1",
    }


def run_terminal_authorization_analysis(target_user: str, target_system: str = "DEV") -> dict[str, Any]:
    resolved_system = target_system.strip().upper() or "DEV"
    if resolved_system == "DEV":
        resolved_key = "S4DCLNT100"
    elif resolved_system == "QAD":
        resolved_key = "S4QCLNT100"
    elif resolved_system == "PRD":
        resolved_key = "S4PCLNT100"
    elif resolved_system == "CUA":
        resolved_key = "SPACLNT001"
    else:
        resolved_key = resolved_system

    if resolved_system == "DEV":
        return _run_dev_role_analysis(target_user, resolved_key, progress_logger=print_progress)

    return analyze_user_authorizations_rfc(
        target_user=target_user,
        target_system_key=resolved_key,
        progress_logger=print_progress,
    )


def main() -> int:
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        target_user = sys.argv[1].strip().upper()
    else:
        target_user, _ = prompt_inputs()

    if len(sys.argv) >= 3 and sys.argv[2].strip():
        target_system = sys.argv[2].strip().upper()
    else:
        target_system = "DEV"

    result = run_terminal_authorization_analysis(target_user, target_system)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
