from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

WORKER_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = WORKER_DIR.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(WORKER_DIR) not in sys.path:
    sys.path.insert(0, str(WORKER_DIR))

try:
    from .authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        normalize_sap_user,
        validate_target_system_key,
    )
except ImportError:
    from authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        normalize_sap_user,
        validate_target_system_key,
    )

SIMULATED_SYSTEM_MAP = {
    "DEV": {"key": "S4DCLNT100", "system": "S4D", "client": "100"},
    "QAD": {"key": "S4QCLNT100", "system": "S4Q", "client": "100"},
    "PRD": {"key": "S4PCLNT100", "system": "S4P", "client": "100"},
    "CUA": {"key": "SPACLNT001", "system": "SPA", "client": "001"},
}


@dataclass(frozen=True)
class SimulatedAuthorizationResult:
    success: bool
    code: str
    message: str
    execution_mode: str
    target_user: str
    target_system: dict[str, str]
    execution_system: dict[str, str]
    user_assigned_to_system: bool
    summary: dict[str, int]
    roles: list[dict[str, Any]]
    profiles: list[dict[str, Any]]
    warnings: list[str]
    truncated: bool
    queries: list[dict[str, Any]]
    data_source_verified: bool
    worker_feature_version: str

    def as_dict(self) -> dict[str, Any]:
        return {
            "success": self.success,
            "code": self.code,
            "message": self.message,
            "execution_mode": self.execution_mode,
            "target_user": self.target_user,
            "target_system": self.target_system,
            "execution_system": self.execution_system,
            "user_assigned_to_system": self.user_assigned_to_system,
            "summary": self.summary,
            "roles": self.roles,
            "profiles": self.profiles,
            "warnings": self.warnings,
            "truncated": self.truncated,
            "queries": self.queries,
            "data_source_verified": self.data_source_verified,
            "worker_feature_version": self.worker_feature_version,
        }


def resolve_simulated_system(system_input: str | None = None, default_choice: str = "DEV") -> dict[str, str]:
    raw = str(system_input or "").strip().upper()
    choice = raw or default_choice.strip().upper()

    if choice in SIMULATED_SYSTEM_MAP:
        return {"choice": choice, **SIMULATED_SYSTEM_MAP[choice]}

    normalized = validate_target_system_key(choice)
    system = normalized.split("CLNT", 1)[0]
    client = normalized.split("CLNT", 1)[1] if "CLNT" in normalized else ""
    return {
        "choice": choice,
        "key": normalized,
        "system": system,
        "client": client,
    }


def prompt_authorization_inputs(default_system: str = "DEV") -> tuple[str, str]:
    target_user = input("Utilizador SAP a analisar: ").strip().upper()
    while not target_user:
        target_user = input("Utilizador SAP a analisar: ").strip().upper()

    system_prompt = f"Sistema alvo [default: {default_system.upper()}]: "
    target_system = input(system_prompt).strip().upper() or default_system.upper()
    return target_user, target_system


def _emit(progress_logger: Callable[[str], None] | None, message: str) -> None:
    if callable(progress_logger):
        progress_logger(message)
    else:
        print(message)


def simulate_authorization_rfc_analysis(
    target_user: str,
    target_system_input: str = "DEV",
    progress_logger: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    target_user = normalize_sap_user(target_user)
    system_info = resolve_simulated_system(target_system_input, default_choice="DEV")

    _emit(progress_logger, f"[SIM RFC] A iniciar a ligação RFC ao sistema {system_info['choice']}...")
    _emit(
        progress_logger,
        f"[SIM RFC] Ligação simulada resolvida para {system_info['system']}/{system_info['client']} ({system_info['key']}).",
    )
    _emit(progress_logger, "[SIM RFC] Vou consultar a USZBVSYS...")

    today_str = "2026-07-23"
    direct_origin = classify_assignment_origin("")
    org_origin = classify_assignment_origin("X")

    roles = deduplicate_roles([
        {
            "role": "Z_SIM_ROLE_01",
            "description": "Role simulada para validação RFC",
            "subsystem": system_info["key"],
            "valid_from": "2026-01-01",
            "valid_to": "2026-12-31",
            "validity_status": classify_validity("2026-01-01", "2026-12-31", today_str),
            "assignment_origin": direct_origin["origin"],
            "assignment_origin_label": direct_origin["origin_label"],
            "assignment_origin_code": "",
        },
        {
            "role": "Z_SIM_ROLE_02",
            "description": "Role simulada de segunda linha",
            "subsystem": system_info["key"],
            "valid_from": "2026-01-01",
            "valid_to": "2026-12-31",
            "validity_status": classify_validity("2026-01-01", "2026-12-31", today_str),
            "assignment_origin": org_origin["origin"],
            "assignment_origin_label": org_origin["origin_label"],
            "assignment_origin_code": "X",
        },
    ])

    profiles = [{"profile": "Z_SIM_PROF_01", "subsystem": system_info["key"]}]
    summary = build_authorization_summary(roles, profiles)
    executed_queries = [
        {"table": "USZBVSYS", "executed": True, "filters_applied": True, "row_count": 1},
        {"table": "USLA04", "executed": True, "filters_applied": True, "row_count": len(roles)},
        {"table": "USL04", "executed": True, "filters_applied": True, "row_count": len(profiles)},
    ]

    _emit(progress_logger, "[SIM RFC] USZBVSYS lida.")
    _emit(progress_logger, "[SIM RFC] Vou consultar as roles em USLA04...")
    _emit(progress_logger, "[SIM RFC] Vou consultar os perfis em USL04...")
    _emit(progress_logger, "[SIM RFC] Análise concluída.")

    result = SimulatedAuthorizationResult(
        success=True,
        code="analysis_complete",
        message="Simulação RFC concluída com sucesso.",
        execution_mode="RFC",
        target_user=target_user,
        target_system={
            "key": system_info["key"],
            "system": system_info["system"],
            "client": system_info["client"],
        },
        execution_system={
            "key": system_info["key"],
            "system": system_info["system"],
            "client": system_info["client"],
        },
        user_assigned_to_system=True,
        summary=summary,
        roles=roles,
        profiles=profiles,
        warnings=["Simulação: dados fictícios, sem ligação a SAP real."],
        truncated=False,
        queries=executed_queries,
        data_source_verified=True,
        worker_feature_version="authorization-tables-v1",
    )
    return result.as_dict()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Simulador do fluxo de autorizações RFC.")
    parser.add_argument("--user", help="Utilizador SAP a analisar.")
    parser.add_argument("--system", default="DEV", help="Sistema alvo. Default: DEV.")
    args = parser.parse_args(argv)

    target_user = args.user or ""
    system_choice = args.system or "DEV"
    if not target_user:
        target_user, system_choice = prompt_authorization_inputs(default_system=system_choice)

    result = simulate_authorization_rfc_analysis(target_user, system_choice)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
