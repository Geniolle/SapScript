import os
from typing import Any

AUTHORIZATION_EXECUTION_ENVIRONMENT = "RFC"

ANALYSIS_TYPES = {
    "master_data": {
        "label": "Dados de Utilizador",
        "description": (
            "Estado do utilizador, validade, bloqueios e "
            "informações gerais da conta SAP."
        ),
        "handler": None,
    },
    "authorizations": {
        "label": "Autorizações",
        "description": (
            "Roles, perfis, transações e objetos de autorização "
            "atribuídos ao utilizador."
        ),
        "handler": None,
    },
}

def get_analysis_types() -> list[dict[str, Any]]:
    return [
        {
            "key": key,
            "label": val["label"],
            "description": val["description"]
        }
        for key, val in ANALYSIS_TYPES.items()
    ]

def get_analysis_type(type_key: str) -> dict[str, Any] | None:
    return ANALYSIS_TYPES.get(type_key)

def validate_analysis_selection(type_key: str) -> bool:
    return type_key in ANALYSIS_TYPES


def get_execution_mode_for_system_key(system_key: str) -> str:
    cua_key = os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001").strip().upper()
    sys_up = str(system_key or "").strip().upper()

    if os.getenv("AUTHORIZATION_FORCE_RFC", "").strip().lower() in {"1", "true", "yes", "sim"}:
        return "RFC"

    if sys_up == cua_key or sys_up in {"CUA", "SPA", "SPACLNT001"}:
        return "CUA"

    return "RFC"
