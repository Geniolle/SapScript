import re
from datetime import datetime
from typing import Any
from sap_agent.sap_gui_actions import se16_query_with_session

def normalize_sap_user(user: str) -> str:
    cleaned = str(user or "").strip().upper()
    if not cleaned:
        raise ValueError("Utilizador a analisar não foi informado.")
    if len(cleaned) > 40:
        raise ValueError("Utilizador a analisar excede o limite de 40 caracteres.")
    if not re.match(r"^[A-Z0-9.\-_]+$", cleaned):
        raise ValueError("Utilizador contém caracteres inválidos.")
    return cleaned

def validate_target_system_key(system_key: str) -> str:
    cleaned = str(system_key or "").strip().upper()
    if not cleaned:
        raise ValueError("Sistema alvo da análise não foi informado.")
    if not re.match(r"^[A-Z0-9_]+CLNT[0-9]+$", cleaned):
        raise ValueError("Formato de sistema lógico inválido.")
    return cleaned

def query_cua_table(
    session: Any,
    table: str,
    filters: list[dict[str, str]],
    max_rows: int = 5000,
) -> list[dict[str, str]]:
    res = se16_query_with_session(
        session,
        table=table,
        filters=filters,
        max_rows=max_rows,
        strict_filters=True,
    )
    if not res.success:
        err_msg = str(res.error or "").lower()
        if "não está na allowlist" in err_msg or "não autorizado" in err_msg:
            raise RuntimeError(f"table_not_authorized: O utilizador técnico não possui autorização para consultar a tabela necessária no CUA ({table}).")
        if "não foi possível aplicar de forma segura" in err_msg:
            raise RuntimeError(f"filter_not_applied: Não foi possível aplicar de forma segura os filtros da análise na tabela {table}.")
        raise RuntimeError(f"table_read_failed: Erro ao consultar a tabela {table}: {res.error}")
    return res.rows

def normalize_sap_date(date_str: str) -> str:
    date_str = str(date_str or "").strip()
    if not date_str or date_str in {"00000000", "0000-00-00", "00.00.0000", "00/00/0000"}:
        return ""
    
    if re.match(r"^\d{8}$", date_str):
        return f"{date_str[0:4]}-{date_str[4:6]}-{date_str[6:8]}"
    
    if re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
        return date_str
        
    if re.match(r"^\d{2}\.\d{2}\.\d{4}$", date_str):
        parts = date_str.split(".")
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
        
    if re.match(r"^\d{2}/\d{2}/\d{4}$", date_str):
        parts = date_str.split("/")
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
        
    return ""

def classify_validity(valid_from: str, valid_to: str, today_str: str) -> str:
    from_cleaned = str(valid_from or "").strip()
    to_cleaned = str(valid_to or "").strip()
    
    from_d = normalize_sap_date(valid_from)
    if from_cleaned and from_cleaned not in {"00000000", "0000-00-00", "00.00.0000", "00/00/0000"} and not from_d:
        return "undetermined"
        
    to_d = normalize_sap_date(valid_to)
    if to_cleaned and to_cleaned not in {"00000000", "0000-00-00", "00.00.0000", "00/00/0000"} and not to_d:
        return "undetermined"
    
    if to_d == "9999-12-31":
        to_d = ""
        
    try:
        today = datetime.strptime(today_str, "%Y-%m-%d").date()
    except Exception:
        return "undetermined"
        
    try:
        from_date = datetime.strptime(from_d, "%Y-%m-%d").date() if from_d else None
    except Exception:
        return "undetermined"
        
    try:
        to_date = datetime.strptime(to_d, "%Y-%m-%d").date() if to_d else None
    except Exception:
        return "undetermined"

    if from_date and from_date > today:
        return "future"
    if to_date and to_date < today:
        return "expired"
    return "active"

def classify_assignment_origin(org_flag: str) -> dict[str, str]:
    code = str(org_flag or "").strip().upper()
    if not code:
        return {"origin": "direct", "origin_label": "Direta"}
    if code == "X":
        return {"origin": "organizational_management", "origin_label": "Organização RH"}
    if code == "C":
        return {"origin": "composite_role", "origin_label": "Role composta"}
    if code == "E":
        return {"origin": "enterprise_portal", "origin_label": "Enterprise Portal"}
    return {"origin": "other", "origin_label": "Outra origem"}

def deduplicate_roles(roles: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen = set()
    deduped = []
    for r in roles:
        key = (
            r["role"],
            r["subsystem"],
            r["valid_from"],
            r["valid_to"],
            r.get("assignment_origin_code", "")
        )
        if key not in seen:
            seen.add(key)
            deduped.append(r)
    return deduped

def build_authorization_summary(roles: list[dict[str, Any]], profiles: list[dict[str, Any]]) -> dict[str, int]:
    total_roles = len(roles)
    active_roles = sum(1 for r in roles if r["validity_status"] == "active")
    expired_roles = sum(1 for r in roles if r["validity_status"] == "expired")
    future_roles = sum(1 for r in roles if r["validity_status"] == "future")
    undetermined_roles = sum(1 for r in roles if r["validity_status"] == "undetermined")
    
    direct_roles = sum(1 for r in roles if r["assignment_origin"] == "direct")
    indirect_roles = total_roles - direct_roles
    
    return {
        "total_roles": total_roles,
        "active_roles": active_roles,
        "expired_roles": expired_roles,
        "future_roles": future_roles,
        "undetermined_roles": undetermined_roles,
        "direct_roles": direct_roles,
        "indirect_roles": indirect_roles,
        "total_profiles": len(profiles)
    }

def analyze_user_authorizations(
    session: Any,
    target_user: str,
    target_system_key: str,
    max_rows: int = 5000,
) -> dict[str, Any]:
    try:
        sys_name = str(session.Info.SystemName or "").strip().upper()
        client = str(session.Info.Client or "").strip()
    except Exception as exc:
        return {
            "success": False,
            "code": "invalid_cua_session",
            "message": f"Erro ao ler informações da sessão SAP: {exc}",
            "roles": [],
            "profiles": []
        }

    if sys_name != "SPA" or client != "001":
        return {
            "success": False,
            "code": "invalid_cua_session",
            "message": f"A sessão disponível ({sys_name}/{client}) não corresponde ao CUA SPA/001.",
            "roles": [],
            "profiles": []
        }

    executed_queries = []

    # 1. USZBVSYS
    try:
        filters_sys = [
            {"field": "BNAME", "value": target_user},
            {"field": "SUBSYSTEM", "value": target_system_key}
        ]
        rows_sys = query_cua_table(session, "USZBVSYS", filters_sys, max_rows=max_rows)
        executed_queries.append({
            "table": "USZBVSYS",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_sys)
        })
    except Exception as exc:
        err_msg = str(exc)
        code = "table_read_failed"
        if ":" in err_msg:
            prefix, rest = err_msg.split(":", 1)
            if prefix.strip() in {"table_not_authorized", "filter_not_applied"}:
                code = prefix.strip()
                err_msg = rest.strip()
        return {
            "success": False,
            "code": code,
            "message": err_msg,
            "roles": [],
            "profiles": []
        }

    if not rows_sys:
        simple_sys = target_system_key
        if "CLNT" in target_system_key:
            simple_sys = target_system_key.split("CLNT", 1)[0]
        return {
            "success": True,
            "code": "user_not_assigned_to_system",
            "message": f"O utilizador {target_user} não está associado ao sistema {simple_sys} no CUA.",
            "user_assigned_to_system": False,
            "roles": [],
            "profiles": [],
            "queries": executed_queries,
            "data_source_verified": True,
            "worker_feature_version": "authorization-tables-v1"
        }

    # 2. USLA04
    try:
        filters_roles = [
            {"field": "BNAME", "value": target_user},
            {"field": "SUBSYSTEM", "value": target_system_key}
        ]
        rows_roles = query_cua_table(session, "USLA04", filters_roles, max_rows=max_rows)
        executed_queries.append({
            "table": "USLA04",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_roles)
        })
    except Exception as exc:
        err_msg = str(exc)
        code = "table_read_failed"
        if ":" in err_msg:
            prefix, rest = err_msg.split(":", 1)
            if prefix.strip() in {"table_not_authorized", "filter_not_applied"}:
                code = prefix.strip()
                err_msg = rest.strip()
        return {
            "success": False,
            "code": code,
            "message": err_msg,
            "roles": [],
            "profiles": []
        }

    # 3. USL04
    try:
        filters_profiles = [
            {"field": "BNAME", "value": target_user},
            {"field": "SUBSYSTEM", "value": target_system_key}
        ]
        rows_profiles = query_cua_table(session, "USL04", filters_profiles, max_rows=max_rows)
        executed_queries.append({
            "table": "USL04",
            "executed": True,
            "filters_applied": True,
            "row_count": len(rows_profiles)
        })
    except Exception as exc:
        err_msg = str(exc)
        code = "table_read_failed"
        if ":" in err_msg:
            prefix, rest = err_msg.split(":", 1)
            if prefix.strip() in {"table_not_authorized", "filter_not_applied"}:
                code = prefix.strip()
                err_msg = rest.strip()
        return {
            "success": False,
            "code": code,
            "message": err_msg,
            "roles": [],
            "profiles": []
        }

    # Processar Roles
    today_str = datetime.now().strftime("%Y-%m-%d")
    raw_roles = []
    for r in rows_roles:
        role_name = str(r.get("AGR_NAME") or "").strip()
        if not role_name:
            continue
        
        valid_from = normalize_sap_date(r.get("FROM_DAT", ""))
        valid_to = normalize_sap_date(r.get("TO_DAT", ""))
        org_flag = str(r.get("ORG_FLAG") or "").strip()
        
        status_info = classify_validity(valid_from, valid_to, today_str)
        origin_info = classify_assignment_origin(org_flag)
        
        raw_roles.append({
            "role": role_name,
            "description": "",
            "subsystem": target_system_key,
            "valid_from": valid_from,
            "valid_to": valid_to,
            "validity_status": status_info,
            "assignment_origin": origin_info["origin"],
            "assignment_origin_label": origin_info["origin_label"],
            "assignment_origin_code": org_flag
        })

    deduped_roles = deduplicate_roles(raw_roles)

    # Processar Perfis
    raw_profiles = []
    for p in rows_profiles:
        profile_name = str(p.get("PROFILE") or "").strip()
        if not profile_name:
            continue
        raw_profiles.append({
            "profile": profile_name,
            "subsystem": target_system_key
        })

    # Deduplicar perfis
    seen_profiles = set()
    deduped_profiles = []
    for p in raw_profiles:
        if p["profile"] not in seen_profiles:
            seen_profiles.add(p["profile"])
            deduped_profiles.append(p)

    # Calcular resumo
    summary = build_authorization_summary(deduped_roles, deduped_profiles)

    # Truncated check
    truncated = len(rows_roles) >= max_rows or len(rows_profiles) >= max_rows
    warnings = []
    if truncated:
        warnings.append(f"A consulta atingiu o limite máximo de {max_rows} linhas.")

    return {
        "success": True,
        "code": "analysis_complete",
        "analysis_type": "authorizations",
        "source": "CUA_USLA04",
        "target_user": target_user,
        "execution_system": {
            "key": "SPACLNT001",
            "system": "SPA",
            "client": "001"
        },
        "target_system": {
            "key": target_system_key,
            "system": target_system_key.split("CLNT", 1)[0] if "CLNT" in target_system_key else target_system_key,
            "client": target_system_key.split("CLNT", 1)[1] if "CLNT" in target_system_key else ""
        },
        "user_assigned_to_system": True,
        "summary": summary,
        "roles": deduped_roles,
        "profiles": deduped_profiles,
        "warnings": warnings,
        "truncated": truncated,
        "queries": executed_queries,
        "data_source_verified": True,
        "worker_feature_version": "authorization-tables-v1"
    }
