from __future__ import annotations

from datetime import datetime
from typing import Any, Callable

try:
    from .authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from .authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        format_sap_date_display,
        normalize_sap_user,
    )
except ImportError:
    from authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        format_sap_date_display,
        normalize_sap_user,
    )


RFC_FEATURE_VERSION = "authorization-tables-v1"


def _split_system_key(system_key: str) -> tuple[str, str]:
    clean_key = str(system_key or "").strip().upper()
    if "CLNT" in clean_key:
        system, client = clean_key.split("CLNT", 1)
        return system, client
    return clean_key, ""


def _emit(progress_logger: Callable[[str], None] | None, message: str) -> None:
    if callable(progress_logger):
        progress_logger(message)


def _read_agr_users_rows(connection: Any, target_user: str, max_rows: int) -> tuple[list[dict[str, str]], str]:
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
            max_rows=max_rows,
        )
        rows = [
            row for row in rows
            if str(row.get(field) or "").strip().upper() == target_user
        ]
        if rows:
            return rows, field
    return [], "UNAME"


def _collect_tcodes_for_roles_batch(
    connection: Any,
    role_names: list[str],
    progress_logger: Callable[[str], None] | None = None,
) -> dict[str, list[str]]:
    _emit(progress_logger, f"[AUTH RFC] A consultar AGR_TCODES em lote para {len(role_names)} roles...")
    role_functions: dict[str, list[str]] = {role: [] for role in role_names}
    if not role_names:
        return role_functions

    chunk_size = 15
    for i in range(0, len(role_names), chunk_size):
        chunk = role_names[i:i + chunk_size]
        options = []
        for idx, r_name in enumerate(chunk):
            prefix = "" if idx == 0 else "OR "
            options.append({"TEXT": f"{prefix}AGR_NAME = '{r_name}'"})

        try:
            res = connection.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="AGR_TCODES",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": "AGR_NAME"}, {"FIELDNAME": "TCODE"}],
                OPTIONS=options,
                ROWCOUNT=5000,
            )
            raw_rows = res.get("DATA") or []
            for raw in raw_rows:
                wa = str(raw.get("WA") or "")
                parts = wa.split("|")
                if len(parts) >= 2:
                    r_name = parts[0].strip()
                    tcode = parts[1].strip().upper()
                    if r_name in role_functions and tcode and tcode not in role_functions[r_name]:
                        role_functions[r_name].append(tcode)
        except Exception as exc:
            for r_name in chunk:
                rows = _read_rfc_table(
                    connection,
                    "AGR_TCODES",
                    ["AGR_NAME", "TCODE"],
                    [{"field": "AGR_NAME", "value": r_name}],
                    max_rows=5000,
                )
                tcodes = []
                for row in rows:
                    tcode = str(row.get("TCODE") or "").strip().upper()
                    if tcode and tcode not in tcodes:
                        tcodes.append(tcode)
                role_functions[r_name] = tcodes

    return role_functions


def analyze_user_authorizations_rfc_dev(
    target_user: str,
    target_system_key: str,
    max_rows: int = 5000,
    progress_logger: Any | None = None,
    connection_params: dict[str, str] | None = None,
) -> dict[str, Any]:
    try:
        target_user = normalize_sap_user(target_user)
        target_system_key = str(target_system_key or "").strip().upper() or "S4DCLNT100"
        system_name, system_client = _split_system_key(target_system_key)
    except Exception as exc:
        return {
            "success": False,
            "code": "invalid_input",
            "message": str(exc),
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": RFC_FEATURE_VERSION,
        }

    connection = None

    try:
        _emit(progress_logger, f"[AUTH RFC] A abrir ligação RFC ao sistema {target_system_key}...")
        connection = _open_rfc_connection(target_system_key, connection_params=connection_params)
        _emit(progress_logger, "[AUTH RFC] Ligação RFC validada. A consultar AGR_USERS...")
    except Exception as exc:
        return {
            "success": False,
            "code": "rfc_connection_failed",
            "message": str(exc),
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": RFC_FEATURE_VERSION,
        }

    executed_queries: list[dict[str, Any]] = []

    try:
        rows_roles, user_field = _read_agr_users_rows(connection, target_user, max_rows)
        executed_queries.append({
            "table": "AGR_USERS",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_roles),
        })

        if not rows_roles:
            _emit(progress_logger, f"[AUTH RFC] Nenhuma role encontrada para {target_user} em AGR_USERS usando {user_field}.")
            return {
                "success": True,
                "code": "user_not_assigned_to_system",
                "message": f"Utilizador {target_user} sem roles devolvidas em AGR_USERS para o sistema {system_name}.",
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
                "queries": executed_queries,
                "data_source_verified": True,
                "source": "RFC_AGR_USERS",
                "worker_feature_version": RFC_FEATURE_VERSION,
            }

        today_str = datetime.now().strftime("%Y-%m-%d")
        raw_roles: list[dict[str, Any]] = []
        for row in rows_roles:
            role_name = str(row.get("AGR_NAME") or "").strip()
            if not role_name:
                continue

            valid_from_raw = str(row.get("FROM_DAT") or "").strip()
            valid_to_raw = str(row.get("TO_DAT") or "").strip()
            origin_info = classify_assignment_origin("")
            raw_roles.append({
                "role": role_name,
                "description": "",
                "subsystem": target_system_key,
                "valid_from": format_sap_date_display(valid_from_raw),
                "valid_to": format_sap_date_display(valid_to_raw),
                "validity_status": classify_validity(valid_from_raw, valid_to_raw, today_str),
                "assignment_origin": origin_info["origin"],
                "assignment_origin_label": origin_info["origin_label"],
                "assignment_origin_code": "",
            })

        roles = deduplicate_roles(raw_roles)
        _emit(progress_logger, f"[AUTH RFC] Encontradas {len(roles)} roles. A consultar AGR_TCODES em lote...")

        role_names = [role["role"] for role in roles]
        role_functions = _collect_tcodes_for_roles_batch(connection, role_names, progress_logger=progress_logger)
        for role in roles:
            role_name = role["role"]
            tcodes = role_functions.get(role_name, [])
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

        _emit(progress_logger, "[AUTH RFC] Análise concluída com sucesso.")

        return {
            "success": True,
            "code": "analysis_complete",
            "analysis_type": "authorizations",
            "source": "RFC_AGR_USERS",
            "message": f"Análise concluída para {target_user} no sistema {system_name}.",
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
            "worker_feature_version": RFC_FEATURE_VERSION,
        }
    except Exception as exc:
        return {
            "success": False,
            "code": "rfc_table_read_failed",
            "message": str(exc),
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": RFC_FEATURE_VERSION,
        }
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
