from __future__ import annotations

import os
from datetime import datetime
from typing import Any

try:
    from pyrfc import Connection
except Exception as exc:  # pragma: no cover - import guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None

try:
    from .authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        format_sap_date_display,
        normalize_sap_date,
        normalize_sap_user,
        validate_target_system_key,
    )
except ImportError:
    from authorization_table_analysis import (
        build_authorization_summary,
        classify_assignment_origin,
        classify_validity,
        deduplicate_roles,
        format_sap_date_display,
        normalize_sap_date,
        normalize_sap_user,
        validate_target_system_key,
    )

RFC_FEATURE_VERSION = "authorization-tables-v1"
AUTHORIZATION_ENV_ALIAS_MAP = {
    "S4D": "DEV",
    "S4Q": "QAD",
    "S4P": "PRD",
    "SPA": "CUA",
}


def _split_system_key(system_key: str) -> tuple[str, str]:
    clean_key = validate_target_system_key(system_key)
    if "CLNT" in clean_key:
        system, client = clean_key.split("CLNT", 1)
        return system, client
    return clean_key, ""


def _authorization_env_alias(system_key: str) -> str:
    system = _split_system_key(validate_target_system_key(system_key))[0]
    return AUTHORIZATION_ENV_ALIAS_MAP.get(system, "")


def _first_env_value(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


def _rfc_literal(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def _build_rfc_connection_params(target_system_key: str) -> dict[str, str]:
    system_key = validate_target_system_key(target_system_key)
    system, client = _split_system_key(system_key)
    alias = _authorization_env_alias(system_key)

    ashost = _first_env_value(
        f"SAP_ASHOST_{system_key}",
        f"SAP_ASHOST_{system}",
        f"SAP_ASHOST_{alias}" if alias else "",
        "SAP_ASHOST",
    )
    sysnr = _first_env_value(
        f"SAP_SYSNR_{system_key}",
        f"SAP_SYSNR_{system}",
        f"SAP_SYSNR_{alias}" if alias else "",
        "SAP_SYSNR",
    ) or "00"
    user = _first_env_value(
        f"SAP_USER_{system_key}",
        f"SAP_USER_{system}",
        f"SAP_USER_{alias}" if alias else "",
        "SAP_USER",
    )
    password = _first_env_value(
        f"SAP_PASSWORD_{system_key}",
        f"SAP_PASSWORD_{system}",
        f"SAP_PASSWORD_{alias}" if alias else "",
        "SAP_PASSWORD",
    )
    language = _first_env_value(
        f"SAP_RFC_LANGUAGE_{system_key}",
        f"SAP_RFC_LANGUAGE_{system}",
        f"SAP_RFC_LANGUAGE_{alias}" if alias else "",
        f"SAP_LANGUAGE_{system_key}",
        f"SAP_LANGUAGE_{system}",
        f"SAP_LANGUAGE_{alias}" if alias else "",
        "SAP_RFC_LANGUAGE",
        "SAP_LANGUAGE",
    ) or "PT"

    if not ashost:
        raise RuntimeError(
            f"Variável SAP_ASHOST não encontrada ou vazia no .env. "
            f"Esperado um destes nomes: SAP_ASHOST_{system_key}, SAP_ASHOST_{system}"
            + (f", SAP_ASHOST_{alias}" if alias else "")
            + ", SAP_ASHOST."
        )
    if not sysnr:
        raise RuntimeError(
            f"Variável SAP_SYSNR não encontrada ou vazia no .env. "
            f"Esperado um destes nomes: SAP_SYSNR_{system_key}, SAP_SYSNR_{system}"
            + (f", SAP_SYSNR_{alias}" if alias else "")
            + ", SAP_SYSNR."
        )
    if not user:
        raise RuntimeError(
            f"Variável SAP_USER não encontrada ou vazia no .env. "
            f"Esperado um destes nomes: SAP_USER_{system_key}, SAP_USER_{system}"
            + (f", SAP_USER_{alias}" if alias else "")
            + ", SAP_USER."
        )
    if not password:
        raise RuntimeError(
            f"Variável de password RFC não encontrada ou vazia no .env. "
            f"Esperado um destes nomes: SAP_PASSWORD_{system_key}, SAP_PASSWORD_{system}"
            + (f", SAP_PASSWORD_{alias}" if alias else "")
            + ", SAP_PASSWORD."
        )

    return {
        "ashost": ashost,
        "sysnr": sysnr,
        "client": client,
        "user": user,
        "passwd": password,
        "lang": language,
        "_system": system,
    }


def _open_rfc_connection(target_system_key: str, connection_params: dict[str, str] | None = None) -> Any:
    if Connection is None:
        raise RuntimeError(f"pyrfc não está disponível no worker: {_PYRFC_IMPORT_ERROR}")

    params = dict(connection_params or _build_rfc_connection_params(target_system_key))
    if not params.get("ashost") or not params.get("sysnr") or not params.get("user") or not params.get("passwd"):
        params = _build_rfc_connection_params(target_system_key)
    conn_params = {key: value for key, value in params.items() if not key.startswith("_")}

    try:
        return Connection(**conn_params)
    except Exception as exc:
        raise RuntimeError(f"Não foi possível abrir a ligação RFC: {exc}") from exc


def _read_rfc_table(
    connection: Any,
    table: str,
    fields: list[str],
    filters: list[dict[str, str]],
    max_rows: int = 5000,
) -> list[dict[str, str]]:
    where_clauses = [
        f"{str(item.get('field') or '').strip().upper()} = {_rfc_literal(str(item.get('value') or '').strip())}"
        for item in filters
        if str(item.get("field") or "").strip() and str(item.get("value") or "").strip() != ""
    ]

    try:
        result = connection.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": field} for field in fields],
            OPTIONS=[{"TEXT": clause} for clause in where_clauses],
            ROWCOUNT=max_rows,
        )
    except Exception as exc:
        err_text = str(exc)
        no_data_markers = (
            "TABLE_WITHOUT_DATA",
            "NO_DATA",
            "tabela sem dados",
            "sem dados",
        )
        if any(marker.lower() in err_text.lower() for marker in no_data_markers):
            return []
        raise RuntimeError(f"rfc_table_read_failed: Erro ao consultar a tabela {table}: {exc}") from exc

    raw_rows = result.get("DATA") or []
    parsed_rows: list[dict[str, str]] = []

    for raw in raw_rows:
        wa = str(raw.get("WA") or "")
        values = wa.split("|")
        if len(values) < len(fields):
            values.extend([""] * (len(fields) - len(values)))
        parsed_rows.append({
            field: str(values[idx]).strip() if idx < len(values) else ""
            for idx, field in enumerate(fields)
        })

    return parsed_rows


def analyze_user_authorizations_rfc(
    target_user: str,
    target_system_key: str,
    max_rows: int = 5000,
    progress_logger: Any | None = None,
    connection_params: dict[str, str] | None = None,
) -> dict[str, Any]:
    try:
        target_user = normalize_sap_user(target_user)
        target_system_key = validate_target_system_key(target_system_key)
        system_name, system_client = _split_system_key(target_system_key)
    except Exception as exc:
        return {
            "success": False,
            "code": "invalid_input",
            "message": f"Entrada inválida: {exc}",
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": RFC_FEATURE_VERSION,
        }

    if callable(progress_logger):
        cua_sap_key = str(os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()
        progress_logger(
            f"[AUTH RFC] Pedido recebido: utilizador={target_user}, sistema={target_system_key}, "
            f"tipo=authorizations, modo=RFC, execution_mode=RFC, cua_sap_key={cua_sap_key}."
        )

    connection = None

    try:
        if callable(progress_logger):
            progress_logger("[AUTH RFC] A abrir ligação RFC...")
        if connection_params:
            resolved = {
                "ashost": str(connection_params.get("ashost") or "").strip(),
                "sysnr": str(connection_params.get("sysnr") or "").strip(),
                "client": str(connection_params.get("client") or "").strip(),
                "user": str(connection_params.get("user") or "").strip(),
                "lang": str(connection_params.get("lang") or "").strip() or "PT",
            }
            if callable(progress_logger):
                progress_logger(
                    "[AUTH RFC] Destino resolvido: "
                    f"{resolved['ashost']}:{resolved['sysnr']}/{resolved['client']} "
                    f"utilizador={resolved['user']}."
                )
        connection = _open_rfc_connection(target_system_key, connection_params=connection_params)
    except Exception as exc:
        return {
            "success": False,
            "code": "rfc_connection_failed",
            "message": f"Não foi possível abrir a ligação RFC: {exc}",
            "roles": [],
            "profiles": [],
            "execution_mode": "RFC",
            "worker_feature_version": RFC_FEATURE_VERSION,
        }

    executed_queries: list[dict[str, Any]] = []

    try:
        if callable(progress_logger):
            progress_logger("[AUTH RFC] Ligação RFC validada. A consultar USZBVSYS...")

        rows_sys = _read_rfc_table(
            connection,
            "USZBVSYS",
            ["BNAME", "SUBSYSTEM"],
            [
                {"field": "BNAME", "value": target_user},
            ],
            max_rows=max_rows,
        )
        rows_sys = [
            row for row in rows_sys
            if str(row.get("SUBSYSTEM") or "").strip().upper() == target_system_key
        ]
        executed_queries.append({
            "table": "USZBVSYS",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_sys),
        })

        if not rows_sys:
            if callable(progress_logger):
                progress_logger("[AUTH RFC] Utilizador não associado ao sistema alvo via RFC.")
            return {
                "success": True,
                "code": "user_not_assigned_to_system",
                "message": f"Utilizador {target_user} não associado ao sistema {system_name} via RFC.",
                "user_assigned_to_system": False,
                "roles": [],
                "profiles": [],
                "queries": executed_queries,
                "data_source_verified": True,
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
                "worker_feature_version": RFC_FEATURE_VERSION,
            }

        if callable(progress_logger):
            progress_logger("[AUTH RFC] Utilizador validado. A consultar USLA04 via RFC...")

        rows_roles = _read_rfc_table(
            connection,
            "USLA04",
            ["BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG"],
            [
                {"field": "BNAME", "value": target_user},
            ],
            max_rows=max_rows,
        )
        rows_roles = [
            row for row in rows_roles
            if str(row.get("SUBSYSTEM") or "").strip().upper() == target_system_key
        ]
        executed_queries.append({
            "table": "USLA04",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_roles),
        })

        if callable(progress_logger):
            progress_logger("[AUTH RFC] A consultar USL04 via RFC...")

        rows_profiles = _read_rfc_table(
            connection,
            "USL04",
            ["BNAME", "SUBSYSTEM", "PROFILE"],
            [
                {"field": "BNAME", "value": target_user},
            ],
            max_rows=max_rows,
        )
        rows_profiles = [
            row for row in rows_profiles
            if str(row.get("SUBSYSTEM") or "").strip().upper() == target_system_key
        ]
        executed_queries.append({
            "table": "USL04",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_profiles),
        })

        today_str = datetime.now().strftime("%Y-%m-%d")
        raw_roles: list[dict[str, Any]] = []
        for row in rows_roles:
            role_name = str(row.get("AGR_NAME") or "").strip()
            if not role_name:
                continue

            valid_from_raw = normalize_sap_date(row.get("FROM_DAT", ""))
            valid_to_raw = normalize_sap_date(row.get("TO_DAT", ""))
            org_flag = str(row.get("ORG_FLAG") or "").strip()
            origin_info = classify_assignment_origin(org_flag)

            raw_roles.append({
                "role": role_name,
                "description": "",
                "subsystem": target_system_key,
                "valid_from": format_sap_date_display(valid_from_raw),
                "valid_to": format_sap_date_display(valid_to_raw),
                "validity_status": classify_validity(valid_from_raw, valid_to_raw, today_str),
                "assignment_origin": origin_info["origin"],
                "assignment_origin_label": origin_info["origin_label"],
                "assignment_origin_code": org_flag,
            })

        deduped_roles = deduplicate_roles(raw_roles)

        raw_profiles: list[dict[str, Any]] = []
        for row in rows_profiles:
            profile_name = str(row.get("PROFILE") or "").strip()
            if not profile_name:
                continue
            raw_profiles.append({
                "profile": profile_name,
                "subsystem": target_system_key,
            })

        seen_profiles: set[str] = set()
        deduped_profiles: list[dict[str, Any]] = []
        for profile in raw_profiles:
            profile_name = profile["profile"]
            if profile_name in seen_profiles:
                continue
            seen_profiles.add(profile_name)
            deduped_profiles.append(profile)

        summary = build_authorization_summary(deduped_roles, deduped_profiles)
        truncated = len(rows_roles) >= max_rows or len(rows_profiles) >= max_rows
        warnings = []
        if truncated:
            warnings.append(f"A consulta atingiu o limite máximo de {max_rows} linhas.")

        if callable(progress_logger):
            progress_logger("[AUTH RFC] Análise concluída com sucesso.")

        return {
            "success": True,
            "code": "analysis_complete",
            "message": "Análise de autorizações concluída com sucesso.",
            "analysis_type": "authorizations",
            "source": "RFC_USLA04",
            "target_user": target_user,
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
            "user_assigned_to_system": True,
            "summary": summary,
            "roles": deduped_roles,
            "profiles": deduped_profiles,
            "warnings": warnings,
            "truncated": truncated,
            "queries": executed_queries,
            "data_source_verified": True,
            "worker_feature_version": RFC_FEATURE_VERSION,
        }
    except Exception as exc:
        err_msg = str(exc)
        code = "rfc_table_read_failed"
        if ":" in err_msg:
            prefix, rest = err_msg.split(":", 1)
            if prefix.strip() in {"table_not_authorized", "filter_not_applied", "rfc_table_read_failed"}:
                code = prefix.strip()
                err_msg = rest.strip()
        return {
            "success": False,
            "code": code,
            "message": err_msg,
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
