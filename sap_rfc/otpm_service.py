"""
otpm_service.py - Mapeamento e validação de ficheiros Excel OTPM com auditoria RFC (TPAMA/T028P).
Reutilizado pelo Agente Salsa IT (web_api e worker).
"""
from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd

from sap_rfc.otpm_validation_service import (
    OtpmValidationService,
    sanitize_panam,
    T028pStatus,
    TpamaStatus,
)


COL_NECESSARIAS: Dict[str, Set[str]] = {
    "NOME CADEIA DE PESQUISA": {
        "NOME CADEIA DE PESQUISA",
        "NOME DA CADEIA DE PESQUISA",
        "NOME CADEIA",
        "CADEIA",
        "PANAM",
    },
    "EMPRESA": {"EMPRESA", "BUKRS", "COMPANY CODE"},
    "BANCO": {"BANCO", "HBKID", "ID BANCO"},
    "ID CONTA": {"ID CONTA", "HKTID", "ID DA CONTA", "ID-CONTA"},
    "OPERAÇÃO EXTERNA": {"OPERAÇÃO EXTERNA", "OPERACAO EXTERNA", "VGEXT", "OP EXTERNA"},
    "SINAL (+/-)": {"SINAL (+/-)", "SINAL", "VOZPM", "SINAL +-"},
    "NOME DO CAMPO DESTINO": {"NOME DO CAMPO DESTINO", "NOME CAMPO DESTINO", "ALVO", "DESTINO"},
    "CAMPO DESTINO": {"CAMPO DESTINO", "TARGFI", "CAMPO DE DESTINO"},
    "BASE DE MAPEAMENTO": {"BASE DE MAPEAMENTO", "PREFIX", "BASE MAPEAMENTO"},
}

MAP_CAMPOS_REAIS: Dict[str, str] = {
    "REGRA DE CONTABILIZAÇÃO": "EBVGINT",
    "REGRA DE CONTABILIZACAO": "EBVGINT",
    "REGRA CONTABILIZAÇÃO": "EBVGINT",
    "REGRA CONTABILIZACAO": "EBVGINT",
    "REGRA": "EBVGINT",
    "BP": "EBAVKON",
    "PARCEIRO DE NEGÓCIOS": "EBAVKON",
    "PARCEIRO DE NEGOCIOS": "EBAVKON",
    "CONTA DO RAZÃO": "EBAVKOA",
    "CONTA DO RAZAO": "EBAVKOA",
    "CONTA RAZÃO": "EBAVKOA",
    "CONTA RAZAO": "EBAVKOA",
    "TIPO DE CONTA": "EBAVKOA",
    "CENTRO": "EBWERKS",
    "CENTRO DE LUCRO": "EBPRCTR",
}

BASE_TO_ALLOWED_SIGNS: Dict[str, Set[str]] = {
    "CP0100000000000": {"-"},
    "Z001": {"+"},
    "Z002": {"+", "-"},
    "Z006": {"+"},
    "Z007": {"-"},
    "Z021": {"-"},
    "Z022": {"+"},
    "Z030": {"+"},
    "Z031": {"-"},
    "Z032": {"+"},
    "Z033": {"-"},
    "Z050": {"+"},
    "Z051": {"-"},
    "Z056": {"+"},
    "Z057": {"-"},
    "Z060": {"+"},
    "Z061": {"-"},
    "Z062": {"+"},
    "Z063": {"-"},
    "Z118": {"+"},
    "Z119": {"-"},
    "ZDD3": {"+"},
    "ZDD4": {"-"},
}


def _norm(s: Any) -> str:
    if s is None:
        return ""
    st = str(s)
    st = "".join(c for c in unicodedata.normalize("NFKD", st) if not unicodedata.combining(c))
    st = st.replace("\u00A0", " ")
    st = re.sub(r"\s+", " ", st.strip())
    return st.upper()


def _safe_str(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    st = str(val).strip()
    return "" if st.lower() in {"nan", "none"} else st


def _map_sign_free(val: str) -> str:
    v = (val or "").strip().upper()
    if v in {"+", "PLUS", "POS", "P"}:
        return "+"
    if v in {"-", "MINUS", "NEG", "M"}:
        return "-"
    return ""


def parse_otpm_excel(file_path: str) -> Tuple[List[Dict[str, Any]], List[str], List[str]]:
    """Lê todas as folhas do ficheiro Excel, normaliza colunas e extrai linhas com cadeias."""
    if not os.path.exists(file_path):
        return [], [], [f"Ficheiro não encontrado: {file_path}"]

    try:
        excel_file = pd.ExcelFile(file_path)
    except Exception as exc:
        return [], [], [f"Erro ao abrir ficheiro Excel: {exc}"]

    parsed_rows: List[Dict[str, Any]] = []
    ignored_examples: List[str] = []
    errors: List[str] = []

    for sheet_name in excel_file.sheet_names:
        try:
            df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str).fillna("")
        except Exception as sheet_exc:
            errors.append(f"Erro ao ler folha '{sheet_name}': {sheet_exc}")
            continue

        if df_sheet.empty:
            continue

        # Resolver colunas
        cols_norm = {_norm(c): c for c in df_sheet.columns}
        resolved: Dict[str, str] = {}
        for canon, syns in COL_NECESSARIAS.items():
            for syn in syns:
                norm_syn = _norm(syn)
                if norm_syn in cols_norm:
                    resolved[canon] = cols_norm[norm_syn]
                    break

        # A folha só é relevante se tiver a coluna da cadeia
        if "NOME CADEIA DE PESQUISA" not in resolved:
            continue

        col_cadeia = resolved["NOME CADEIA DE PESQUISA"]
        col_empresa = resolved.get("EMPRESA")
        col_banco = resolved.get("BANCO")
        col_id_conta = resolved.get("ID CONTA")
        col_op_ext = resolved.get("OPERAÇÃO EXTERNA")
        col_sinal = resolved.get("SINAL (+/-)")
        col_nome_dest = resolved.get("NOME DO CAMPO DESTINO")
        col_campo_dest = resolved.get("CAMPO DESTINO")
        col_base_map = resolved.get("BASE DE MAPEAMENTO")

        for idx, row in df_sheet.iterrows():
            row_num = idx + 2  # Cabeçalho na linha 1
            raw_cadeia = _safe_str(row.get(col_cadeia))
            if not raw_cadeia:
                continue

            # Identificar exemplos a ignorar
            cadeia_upper = _norm(raw_cadeia)
            if "EXEMPLO" in cadeia_upper or "FIDELIDADE" in cadeia_upper:
                ignored_examples.append(
                    f"Folha '{sheet_name}', Linha {row_num}: '{raw_cadeia}' ignorado (exemplo/modelo)."
                )
                continue

            empresa = _safe_str(row.get(col_empresa)) if col_empresa else ""
            banco = _safe_str(row.get(col_banco)) if col_banco else ""
            id_conta = _safe_str(row.get(col_id_conta)) if col_id_conta else ""
            op_ext = _safe_str(row.get(col_op_ext)) if col_op_ext else ""
            sinal_raw = _safe_str(row.get(col_sinal)) if col_sinal else ""
            nome_campo_dest = _safe_str(row.get(col_nome_dest)) if col_nome_dest else ""
            campo_dest_raw = _safe_str(row.get(col_campo_dest)) if col_campo_dest else ""
            base_map = _safe_str(row.get(col_base_map)) if col_base_map else ""

            # Normalização de PANAM
            panam = sanitize_panam(raw_cadeia)
            if not panam:
                errors.append(f"Folha '{sheet_name}', Linha {row_num}: Nome de cadeia inválido após normalização.")
                continue

            # Resolução do campo de destino (TARGFI)
            targfi = campo_dest_raw.strip().upper()
            if not targfi:
                targfi = MAP_CAMPOS_REAIS.get(_norm(nome_campo_dest), "")

            # Resolução do sinal (VOZPM)
            sinal = _map_sign_free(sinal_raw)
            if not sinal and base_map:
                allowed_signs = BASE_TO_ALLOWED_SIGNS.get(base_map.strip().upper())
                if allowed_signs and len(allowed_signs) == 1:
                    sinal = list(allowed_signs)[0]

            parsed_rows.append({
                "sheet": sheet_name,
                "row_index": row_num,
                "raw_cadeia": raw_cadeia,
                "panam": panam,
                "empresa": empresa.upper(),
                "banco": banco.upper(),
                "id_conta": id_conta.upper(),
                "op_externa": op_ext.upper(),
                "sinal": sinal,
                "nome_campo_destino": nome_campo_dest,
                "targfi": targfi,
                "base_mapeamento": base_map,
            })

    return parsed_rows, ignored_examples, errors


def validate_otpm_excel_via_rfc(
    file_path: str,
    environment: str = "DEV",
) -> Dict[str, Any]:
    """Processa o Excel e valida cada cadeia e respetivas atribuições via RFC no sistema SAP indicado."""
    parsed_rows, ignored_examples, parse_errors = parse_otpm_excel(file_path)
    if parse_errors and not parsed_rows:
        return {
            "ok": False,
            "status": "PARSE_ERROR",
            "message": f"Erro ao analisar o Excel: {'; '.join(parse_errors)}",
            "errors": parse_errors,
        }

    # Agrupar linhas por cadeia PANAM
    cadeias_map: Dict[str, List[Dict[str, Any]]] = {}
    for r in parsed_rows:
        cadeias_map.setdefault(r["panam"], []).append(r)

    results_by_chain: List[Dict[str, Any]] = []
    overall_ok = True
    total_actions_needed = 0

    with OtpmValidationService(environment=environment) as svc:
        for panam, rows in cadeias_map.items():
            # 1. Validação TPAMA
            tpama_res = svc.check_tpama(panam)

            # Construir lista esperada para T028P
            expected_t028p: List[Dict[str, Any]] = []
            for row in rows:
                expected_t028p.append({
                    "BUKRS": row["empresa"],
                    "HBKID": row["banco"],
                    "HKTID": row["id_conta"],
                    "VGEXT": row["op_externa"],
                    "VOZPM": row["sinal"],
                    "TARGFI": row["targfi"],
                    "PREFIX": row["base_mapeamento"],
                    "ENABLED": "X",
                })

            # 2. Validação T028P
            t028p_res = svc.validate_t028p_assignments(panam, expected_t028p)

            # Determinar se requer criação/atribuição
            requires_tpama_create = tpama_res.status == TpamaStatus.NAO_EXISTE
            requires_t028p_assign = (
                t028p_res.status in (T028pStatus.NAO_EXISTE, T028pStatus.PARCIAL)
                or len(t028p_res.chaves_em_falta) > 0
            )

            if requires_tpama_create or requires_t028p_assign:
                total_actions_needed += 1

            if tpama_res.status == TpamaStatus.ERRO_RFC or t028p_res.status == T028pStatus.ERRO_RFC:
                overall_ok = False

            # Montar detalhe das atribuições
            assignments_detail: List[Dict[str, Any]] = []
            for row, exp in zip(rows, expected_t028p):
                # Verificar se esta atribuição específica existe
                exp_key = (
                    exp["BUKRS"],
                    exp["HBKID"],
                    exp["HKTID"],
                    exp["VGEXT"],
                    exp["VOZPM"],
                    panam,
                    exp["TARGFI"],
                    exp["PREFIX"],
                )
                missing = exp_key in t028p_res.chaves_em_falta
                assignments_detail.append({
                    "row_index": row["row_index"],
                    "sheet": row["sheet"],
                    "empresa": row["empresa"],
                    "banco": row["banco"],
                    "id_conta": row["id_conta"],
                    "op_externa": row["op_externa"],
                    "sinal": row["sinal"],
                    "targfi": row["targfi"],
                    "nome_campo_destino": row["nome_campo_destino"],
                    "base_mapeamento": row["base_mapeamento"],
                    "exists_in_sap": not missing and t028p_res.status != T028pStatus.NAO_EXISTE,
                    "action_needed": "CRIAR_ATRIBUICAO" if (missing or t028p_res.status == T028pStatus.NAO_EXISTE) else "NENHUMA",
                })

            results_by_chain.append({
                "panam": panam,
                "raw_name": rows[0]["raw_cadeia"],
                "tpama": {
                    "status": tpama_res.status.value,
                    "message": tpama_res.message,
                    "regex": tpama_res.record.regex if tpama_res.record else "",
                    "mapex": tpama_res.record.mapex if tpama_res.record else "",
                    "note": tpama_res.record.note if tpama_res.record else "",
                },
                "t028p": {
                    "status": t028p_res.status.value,
                    "total_esperado": t028p_res.total_esperado,
                    "total_encontrado": t028p_res.total_encontrado,
                    "message": t028p_res.message,
                    "divergencias": t028p_res.divergencias,
                    "chaves_em_falta_count": len(t028p_res.chaves_em_falta),
                },
                "assignments": assignments_detail,
                "action_needed": (
                    "CRIAR_CADEIA_E_ATRIBUICOES" if requires_tpama_create
                    else ("ATRIBUIR_FALTANTES" if requires_t028p_assign else "NENHUMA_JA_EXISTE")
                ),
            })

    return {
        "ok": overall_ok,
        "status": "VALIDATED",
        "environment": environment,
        "file_name": Path(file_path).name,
        "total_chains": len(results_by_chain),
        "total_actions_needed": total_actions_needed,
        "ignored_examples": ignored_examples,
        "parse_errors": parse_errors,
        "chains": results_by_chain,
    }
