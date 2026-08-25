import re
import os
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


def format_sap_date_display(date_str: str) -> str:
    normalized = normalize_sap_date(date_str)
    if not normalized:
        return ""
    try:
        dt = datetime.strptime(normalized, "%Y-%m-%d").date()
    except Exception:
        return normalized
    return dt.strftime("%d/%m/%Y")

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
    progress_logger: Any | None = None,
    target_subsystem_key: str | None = None,
) -> dict[str, Any]:
    try:
        sys_name = str(session.Info.SystemName or "").strip().upper()
        client = str(session.Info.Client or "").strip()
    except Exception as exc:
        return {
            "success": False,
            "code": "invalid_cua_session",
            "message": f"Não foi possível validar a sessão SAP: {exc}",
            "roles": [],
            "profiles": []
        }

    if sys_name != "SPA" or client != "001":
        return {
            "success": False,
            "code": "invalid_cua_session",
            "message": f"Sessão CUA inválida: {sys_name}/{client} (esperado SPA/001).",
            "roles": [],
            "profiles": []
        }

    subsystem_to_match = str(target_subsystem_key or target_system_key).strip().upper()

    if callable(progress_logger):
        cua_sap_key = str(os.getenv("AUTHORIZATION_CUA_SAP_KEY", "SPACLNT001")).strip().upper()
        progress_logger(
            f"[AUTH] Pedido recebido: utilizador={target_user}, sistema={target_system_key}, subsystem={subsystem_to_match}, "
            f"tipo=authorizations, modo=CUA, execution_mode=CUA, cua_sap_key={cua_sap_key}."
        )

    executed_queries = []

    # 1. USLA04 (Tabela principal e única necessária para funções CUA)
    try:
        if callable(progress_logger):
            progress_logger("[AUTH] A abrir sessão CUA...")
            sub_msg = subsystem_to_match if subsystem_to_match and subsystem_to_match.upper() not in {"", "ALL", "TODOS", "SPA", "SPACLNT001"} else "todos os subsistemas"
            progress_logger(f"[AUTH] A consultar tabela USLA04 para {target_user} ({sub_msg})...")
        
        filters_roles = [
            {"field": "BNAME", "value": target_user}
        ]
        if subsystem_to_match and subsystem_to_match.upper() not in {"", "ALL", "TODOS", "SPA", "SPACLNT001"}:
            filters_roles.append({"field": "SUBSYSTEM", "value": subsystem_to_match})

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
            "message": f"Falha ao consultar USLA04 no CUA: {err_msg}",
            "roles": [],
            "profiles": []
        }

    # Processar Roles
    today_str = datetime.now().strftime("%Y-%m-%d")
    raw_roles = []
    for r in rows_roles:
        role_name = ""
        # 1. Tentar chave exata ou variantes de cabeçalhos SAP GUI (PT/EN)
        for key, val in r.items():
            k_upper = str(key or "").strip().upper()
            if k_upper in {
                "AGR_NAME", "ROLE", "FUNÇÃO", "FUNCAO", "NOME DA FUNÇÃO", "NOME DA FUNCAO",
                "AGRUPAD.PERFIS", "AGRUPAD. PERFIS", "AGRUPAMENTO DE PERFIS", "CONJ.PERFIS", "AGRUPAMENTO"
            }:
                role_name = str(val or "").strip()
                if role_name:
                    break

        # 2. Fallback: procurar por valor textual que siga o formato de role SAP (ex: Z... ou Y...)
        if not role_name:
            for val in r.values():
                v_str = str(val or "").strip()
                if len(v_str) >= 3 and not re.match(r"^\d+$", v_str) and v_str not in {subsystem_to_match, target_system_key, target_user, "SPACLNT001"}:
                    if re.match(r"^[A-Z0-9_/:\-]+$", v_str) and not re.match(r"^\d{4}-\d{2}-\d{2}$", v_str) and not re.match(r"^\d{8}$", v_str):
                        role_name = v_str
                        break

        if not role_name:
            continue

        row_subsystem = ""
        for key, val in r.items():
            k_upper = str(key or "").strip().upper()
            if k_upper in {"SUBSYSTEM", "SUBSYS", "SISTEMA", "SUBSISTEMA", "SISTEMA ALVO", "SUBSYSTEMA"}:
                row_subsystem = str(val or "").strip()
                if row_subsystem:
                    break
        if not row_subsystem:
            row_subsystem = str(r.get("SUBSYSTEM") or subsystem_to_match or "").strip()

        if subsystem_to_match and subsystem_to_match.upper() not in {"", "ALL", "TODOS", "SPA", "SPACLNT001"}:
            sub_target = subsystem_to_match.upper().split("CLNT")[0]
            sub_row = row_subsystem.upper().split("CLNT")[0]
            if sub_target != sub_row and subsystem_to_match.upper() != row_subsystem.upper():
                continue

        valid_from_raw = ""
        for key, val in r.items():
            k_upper = str(key or "").strip().upper()
            if k_upper in {"FROM_DAT", "VÁLIDO DE", "VALIDO DE", "DE", "DATA INÍCIO", "DATA INICIO"}:
                valid_from_raw = str(val or "").strip()
                if valid_from_raw:
                    break
        if not valid_from_raw:
            valid_from_raw = str(r.get("FROM_DAT") or "").strip()

        valid_to_raw = ""
        for key, val in r.items():
            k_upper = str(key or "").strip().upper()
            if k_upper in {"TO_DAT", "VÁLIDO ATÉ", "VALIDO ATE", "ATÉ", "ATE", "DATA FIM"}:
                valid_to_raw = str(val or "").strip()
                if valid_to_raw:
                    break
        if not valid_to_raw:
            valid_to_raw = str(r.get("TO_DAT") or "").strip()

        org_flag = ""
        for key, val in r.items():
            k_upper = str(key or "").strip().upper()
            if k_upper in {"ORG_FLAG", "ORG", "FLAG", "ORIGEM", "TIPO"}:
                org_flag = str(val or "").strip()
                if org_flag:
                    break
        if not org_flag:
            org_flag = str(r.get("ORG_FLAG") or "").strip()

        status_info = classify_validity(valid_from_raw, valid_to_raw, today_str)
        origin_info = classify_assignment_origin(org_flag)
        
        raw_roles.append({
            "role": role_name,
            "description": "",
            "subsystem": row_subsystem,
            "valid_from": format_sap_date_display(valid_from_raw),
            "valid_to": format_sap_date_display(valid_to_raw),
            "validity_status": status_info,
            "assignment_origin": origin_info["origin"],
            "assignment_origin_label": origin_info["origin_label"],
            "assignment_origin_code": org_flag
        })

    deduped_roles = deduplicate_roles(raw_roles)
    deduped_profiles = []

    # Calcular resumo global e agrupamento por sistema
    summary = build_authorization_summary(deduped_roles, deduped_profiles)

    systems_summary_map = {}
    for r_item in deduped_roles:
        sys_key = r_item.get("subsystem") or "OUTROS"
        if sys_key not in systems_summary_map:
            sys_name = sys_key.split("CLNT", 1)[0] if "CLNT" in sys_key else sys_key
            systems_summary_map[sys_key] = {
                "subsystem": sys_key,
                "system": sys_name,
                "roles_count": 0
            }
        systems_summary_map[sys_key]["roles_count"] += 1
    systems_summary = list(systems_summary_map.values())

    # Truncated check
    truncated = len(rows_roles) >= max_rows
    warnings = []
    if truncated:
        warnings.append(f"A consulta atingiu o limite máximo de {max_rows} linhas.")

    if callable(progress_logger):
        progress_logger("[AUTH] Análise concluída com sucesso.")

    return {
        "success": True,
        "code": "analysis_complete",
        "message": f"Leitura da tabela USLA04 concluída com sucesso. Encontradas {len(deduped_roles)} funções.",
        "analysis_type": "authorizations",
        "source": "CUA_USLA04",
        "target_user": target_user,
        "execution_mode": "CUA",
        "execution_system": {
            "key": "SPACLNT001",
            "system": "SPA",
            "client": "001",
        },
        "target_system": {
            "key": subsystem_to_match,
            "system": subsystem_to_match.split("CLNT", 1)[0] if "CLNT" in subsystem_to_match else subsystem_to_match,
            "client": subsystem_to_match.split("CLNT", 1)[1] if "CLNT" in subsystem_to_match else "",
        },
        "user_assigned_to_system": len(deduped_roles) > 0,
        "summary": summary,
        "systems_summary": systems_summary,
        "roles": deduped_roles,
        "profiles": [],
        "warnings": warnings,
        "truncated": truncated,
        "queries": executed_queries,
        "data_source_verified": True,
        "worker_feature_version": "authorization-tables-v1"
    }
