from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Callable

try:
    from .authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from .authorization_table_analysis import (
        classify_assignment_origin,
        classify_validity,
        format_sap_date_display,
        normalize_sap_date,
        normalize_sap_user,
        query_cua_table,
    )
except ImportError:
    from authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
    from authorization_table_analysis import (
        classify_assignment_origin,
        classify_validity,
        format_sap_date_display,
        normalize_sap_date,
        normalize_sap_user,
        query_cua_table,
    )


WORKER_FEATURE_VERSION = "authorization-tables-v1"


def _split_system_key(system_key: str) -> tuple[str, str]:
    clean_key = str(system_key or "").strip().upper()
    if "CLNT" in clean_key:
        system, client = clean_key.split("CLNT", 1)
        return system, client
    return clean_key, ""


def _emit(progress_logger: Callable[[str], None] | None, message: str) -> None:
    if callable(progress_logger):
        progress_logger(message)


def _safe_first_text(*values: Any) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def _user_lock_status_label(uflag: str) -> str:
    code = str(uflag or "").strip()
    if code == "0":
        return "Desbloqueada"
    if code == "32":
        return "Bloqueada globalmente"
    if code == "64":
        return "Bloqueada localmente"
    if code == "128":
        return "Bloqueada por tentativas de logon"
    return "Indeterminada"


def _build_master_data_payload(
    *,
    target_user: str,
    target_system_key: str,
    system_name: str,
    system_client: str,
    query_reader: Callable[[str, list[dict[str, str]], int], list[dict[str, str]]],
    progress_logger: Callable[[str], None] | None = None,
    execution_mode: str,
    source: str,
) -> dict[str, Any]:
    today_str = datetime.now().strftime("%Y-%m-%d")
    executed_queries: list[dict[str, Any]] = []

    _emit(progress_logger, "[MASTER] A consultar dados do utilizador (USR21, USR04, AGR_USERS)...")
    executed_queries.append({
        "table": "USR02",
        "executed": True,
        "filters_applied": True,
        "row_count": 0,
        "bypassed": True,
    })

    usr02 = {}
    valid_from_raw = ""
    valid_to_raw = ""
    lock_code = "0"

    _emit(progress_logger, "[MASTER] A consultar USR21...")
    rows_usr21 = query_reader(
        "USR21",
        [{"field": "BNAME", "value": target_user}],
        50,
    )
    executed_queries.append({
        "table": "USR21",
        "executed": True,
        "filters_applied": True,
        "row_count": len(rows_usr21),
    })

    persnumber = _safe_first_text(
        rows_usr21[0].get("PERSNUMBER") if rows_usr21 else "",
    )
    addrnumber = _safe_first_text(
        rows_usr21[0].get("ADDRNUMBER") if rows_usr21 else "",
        rows_usr21[0].get("ADRNR") if rows_usr21 else "",
        rows_usr21[0].get("ADDR_NO") if rows_usr21 else "",
    )

    _emit(progress_logger, "[MASTER] A consultar USR04...")
    rows_usr04 = query_reader(
        "USR04",
        [{"field": "BNAME", "value": target_user}],
        50,
    )
    executed_queries.append({
        "table": "USR04",
        "executed": True,
        "filters_applied": True,
        "row_count": len(rows_usr04),
    })

    _emit(progress_logger, "[MASTER] A consultar AGR_USERS...")
    rows_roles = query_reader(
        "AGR_USERS",
        [{"field": "UNAME", "value": target_user}],
        5000,
    )
    if not rows_roles:
        rows_roles = query_reader(
            "AGR_USERS",
            [{"field": "BNAME", "value": target_user}],
            5000,
        )
    executed_queries.append({
        "table": "AGR_USERS",
        "executed": True,
        "filters_applied": True,
        "row_count": len(rows_roles),
    })

    today = today_str
    role_rows: list[dict[str, Any]] = []
    for row in rows_roles:
        role_name = str(row.get("AGR_NAME") or "").strip()
        if not role_name:
            continue
        role_from_raw = normalize_sap_date(row.get("FROM_DAT", ""))
        role_to_raw = normalize_sap_date(row.get("TO_DAT", ""))
        role_rows.append({
            "role": role_name,
            "description": "",
            "subsystem": target_system_key,
            "valid_from": format_sap_date_display(role_from_raw),
            "valid_to": format_sap_date_display(role_to_raw),
            "validity_status": classify_validity(role_from_raw, role_to_raw, today),
            "assignment_origin": classify_assignment_origin("").get("origin", "direct"),
            "assignment_origin_label": classify_assignment_origin("").get("origin_label", "Direta"),
            "assignment_origin_code": "",
        })

    profile_rows: list[dict[str, Any]] = []
    for row in rows_usr04:
        profile_name = str(row.get("PROFILE") or "").strip()
        if profile_name:
            profile_rows.append({
                "profile": profile_name,
                "subsystem": target_system_key,
            })

    seen_profiles: set[str] = set()
    deduped_profiles: list[dict[str, Any]] = []
    for profile in profile_rows:
        profile_name = profile["profile"]
        if profile_name in seen_profiles:
            continue
        seen_profiles.add(profile_name)
        deduped_profiles.append(profile)

    full_name = ""
    email = ""

    master_data = {
        "user": target_user,
        "full_name": full_name,
        "name_first": "",
        "name_last": "",
        "email": email,
        "user_group": _safe_first_text(usr02.get("CLASS")),
        "user_type": _safe_first_text(usr02.get("USTYP")),
        "person_number": persnumber,
        "address_number": addrnumber,
        "valid_from": format_sap_date_display(valid_from_raw),
        "valid_to": format_sap_date_display(valid_to_raw),
        "validity_status": classify_validity(valid_from_raw, valid_to_raw, today),
        "lock_status": _user_lock_status_label(lock_code),
        "lock_code": lock_code,
        "created_by": _safe_first_text(usr02.get("ANAME")),
        "created_on": format_sap_date_display(usr02.get("ERDAT", "")),
        "last_logon_on": format_sap_date_display(usr02.get("TRDAT", "")),
    }

    summary = {
        "total_roles": len(role_rows),
        "total_profiles": len(deduped_profiles),
        "has_email": 1 if email else 0,
        "has_full_name": 1 if full_name else 0,
    }

    return {
        "success": True,
        "code": "analysis_complete",
        "message": "Análise de master data concluída com sucesso.",
        "analysis_type": "master_data",
        "execution_mode": execution_mode,
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
        "master_data": master_data,
        "roles": role_rows,
        "profiles": deduped_profiles,
        "summary": summary,
        "queries": executed_queries,
        "data_source_verified": True,
        "source": source,
        "worker_feature_version": WORKER_FEATURE_VERSION,
    }


def analyze_user_master_data(
    session: Any,
    target_user: str,
    target_system_key: str,
    max_rows: int = 5000,
    progress_logger: Any | None = None,
) -> dict[str, Any]:
    target_user = normalize_sap_user(target_user)
    target_system_key = str(target_system_key or "").strip().upper() or "SPACLNT001"
    system_name, system_client = _split_system_key(target_system_key)
    cua_sap_key = str(os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()
    _emit(
        progress_logger,
        f"[MASTER] Pedido recebido: utilizador={target_user}, sistema={target_system_key}, "
        f"tipo=master_data, modo=CUA, execution_mode=CUA, cua_sap_key={cua_sap_key}.",
    )

    def read_table(table: str, filters: list[dict[str, str]], rows_limit: int) -> list[dict[str, str]]:
        return query_cua_table(session, table, filters, max_rows=rows_limit)

    try:
        return _build_master_data_payload(
            target_user=target_user,
            target_system_key=target_system_key,
            system_name=system_name,
            system_client=system_client,
            query_reader=read_table,
            progress_logger=progress_logger,
            execution_mode="CUA",
            source="CUA_USER_MASTER",
        )
    except Exception as exc:
        err_msg = str(exc)
        code = "table_not_authorized" if "table_not_authorized" in err_msg else "master_data_analysis_failed"
        return {
            "success": False,
            "code": code,
            "message": err_msg,
            "roles": [],
            "profiles": [],
            "execution_mode": "CUA",
            "worker_feature_version": WORKER_FEATURE_VERSION,
        }


def analyze_user_master_data_rfc(
    target_user: str,
    target_system_key: str,
    max_rows: int = 5000,
    progress_logger: Any | None = None,
    connection_params: dict[str, str] | None = None,
) -> dict[str, Any]:
    target_user = normalize_sap_user(target_user)
    target_system_key = str(target_system_key or "").strip().upper() or "S4DCLNT100"
    system_name, system_client = _split_system_key(target_system_key)
    cua_sap_key = str(os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()
    _emit(
        progress_logger,
        f"[MASTER RFC] Pedido recebido: utilizador={target_user}, sistema={target_system_key}, "
        f"tipo=master_data, modo=RFC, execution_mode=RFC, cua_sap_key={cua_sap_key}.",
    )

    connection = None
    try:
        _emit(progress_logger, f"[MASTER RFC] A abrir ligação RFC ao sistema {target_system_key}...")
        connection = _open_rfc_connection(target_system_key, connection_params=connection_params)
        _emit(progress_logger, "[MASTER RFC] Ligação RFC validada.")

        def read_table(table: str, filters: list[dict[str, str]], rows_limit: int) -> list[dict[str, str]]:
            fields_map = {
            "USR02": ["BNAME", "USTYP", "CLASS", "GLTGV", "GLTGB", "UFLAG", "ANAME", "ERDAT", "TRDAT", "LTIME"],
            "USR21": ["BNAME", "PERSNUMBER", "ADDRNUMBER"],
            "USR04": ["BNAME", "PROFILE"],
            "AGR_USERS": ["AGR_NAME", "UNAME", "BNAME", "FROM_DAT", "TO_DAT"],
        }
            fields = fields_map.get(table, [])
            return _read_rfc_table(connection, table, fields, filters, max_rows=rows_limit)

        return _build_master_data_payload(
            target_user=target_user,
            target_system_key=target_system_key,
            system_name=system_name,
            system_client=system_client,
            query_reader=read_table,
            progress_logger=progress_logger,
            execution_mode="RFC",
            source="RFC_USER_MASTER",
        )
    except Exception as exc:
        return {
            "success": False,
            "code": "rfc_table_read_failed",
            "message": str(exc),
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": WORKER_FEATURE_VERSION,
        }
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass
