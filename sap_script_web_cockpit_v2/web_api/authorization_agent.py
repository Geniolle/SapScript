from typing import Any

AUTHORIZATION_EXECUTION_ENVIRONMENT = "CUA"

ANALYSIS_TYPES = {
    "master_data": {
        "label": "Dados mestre",
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
