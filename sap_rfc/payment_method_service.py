"""
payment_method_service.py - Serviço RFC para consulta e análise de Meios de Pagamento por País (T042Z)
"""
from __future__ import annotations

import string
from typing import Any, Dict, List, Optional

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    classify_rfc_error,
    find_project_root,
    load_project_env,
    make_read_only_guard,
    read_table,
)


def suggest_next_code(used_codes: set[str]) -> tuple[Optional[str], List[str]]:
    """
    Dada a lista de códigos de 1 caractere em uso, sugere o próximo código disponível.
    Sequência prioritária de busca:
    1. Letras maiúsculas de A a Z
    2. Dígitos de 1 a 9 e 0
    """
    all_possible = list(string.ascii_uppercase) + [str(i) for i in range(1, 10)] + ["0"]
    available = [c for c in all_possible if c not in used_codes]
    
    next_code = available[0] if available else None
    return next_code, available


def get_payment_method_interval_analysis(
    environment: str = "DEV",
    country: str = "PT"
) -> Dict[str, Any]:
    """
    Pesquisa os Meios de Pagamento existentes na tabela T042Z para um país via RFC,
    identifica o último código atribuído e sugere o próximo código disponível.
    """
    root = find_project_root()
    load_project_env(root)

    params = build_connection_params_for_env(environment)
    from pyrfc import Connection

    conn = Connection(**params)
    guard = make_read_only_guard(allowed_tables=("T042Z",))

    try:
        country_code = country.upper().strip()
        # Ler registros da T042Z para o país
        rows_t042z = read_table(
            conn,
            guard,
            table_name="T042Z",
            fields=["LAND1", "ZLSCH", "TEXT1"],
            options=[{"TEXT": f"LAND1 = '{country_code}'"}],
            rowcount=0,
        )
        conn.close()

        items = []
        used_codes = set()
        for row in rows_t042z:
            land1 = row[0].strip()
            zlsch = row[1].strip()
            text1 = row[2].strip()
            if zlsch:
                used_codes.add(zlsch)
                items.append({
                    "country": land1,
                    "code": zlsch,
                    "description": text1
                })

        # Ordenar itens existentes pelo código
        items.sort(key=lambda x: x["code"])
        sorted_codes = [x["code"] for x in items]
        last_code = sorted_codes[-1] if sorted_codes else None

        next_code, available_codes = suggest_next_code(used_codes)

        return {
            "success": True,
            "environment": environment,
            "country": country_code,
            "total_existing": len(items),
            "used_codes": sorted_codes,
            "last_used_code": last_code,
            "next_suggested_code": next_code,
            "remaining_available_count": len(available_codes),
            "items": items,
        }

    except Exception as exc:
        try:
            conn.close()
        except Exception:
            pass
        code, msg = classify_rfc_error(exc)
        return {
            "success": False,
            "environment": environment,
            "country": country,
            "error_code": code,
            "error_message": msg,
            "detail": str(exc),
        }
