"""
zterm_service.py - Serviço RFC para consulta e suporte a Condições de Pagamento (T052/T052U)
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional
from pathlib import Path
import sys

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    classify_rfc_error,
    make_read_only_guard,
    read_table,
)


def get_zterm_interval_analysis(environment: str = "DEV", prefix: str = "Z") -> Dict[str, Any]:
    """
    Pesquisa as Condições de Pagamento existentes na tabela T052/T052U via RFC
    e identifica o próximo código disponível no intervalo (ex: prefixo 'Z').
    """
    from sap_rfc._rfc_common import find_project_root, load_project_env
    root = find_project_root()
    load_project_env(root)

    params = build_connection_params_for_env(environment)
    from pyrfc import Connection

    conn = Connection(**params)
    guard = make_read_only_guard(allowed_tables=("T052", "T052U"))

    try:
        # 1. Ler todas as ZTERMs de T052
        rows_t052 = read_table(
            conn,
            guard,
            table_name="T052",
            fields=["ZTERM", "ZTAG1", "ZPRZ1", "ZTAG2", "ZPRZ2", "ZTAG3"],
            options=[],
            rowcount=0,
        )

        existing_zterms = set()
        zterm_map = {}

        for row in rows_t052:
            zterm = row[0].strip()
            if not zterm:
                continue
            existing_zterms.add(zterm)
            zterm_map[zterm] = {
                "zterm": zterm,
                "ztag1": row[1].strip(),
                "zprz1": row[2].strip(),
                "ztag2": row[3].strip(),
                "zprz2": row[4].strip(),
                "ztag3": row[5].strip(),
                "text": "",
            }

        # 2. Ler descrições de T052U
        rows_t052u = read_table(
            conn,
            guard,
            table_name="T052U",
            fields=["ZTERM", "SPRAS", "TEXT1"],
            options=[],
            rowcount=0,
        )

        for row in rows_t052u:
            zt = row[0].strip()
            lang = row[1].strip()
            txt = row[2].strip()
            if zt in zterm_map:
                # Priorizar PT ('P' ou 'PT') ou preencher se vazio
                if not zterm_map[zt]["text"] or lang in ("P", "PT"):
                    zterm_map[zt]["text"] = txt

        conn.close()

        # 3. Filtrar pelo prefixo (ex: Z*)
        pref = prefix.upper().strip()
        matched_terms = [zt for zt in existing_zterms if zt.startswith(pref)]
        matched_terms.sort()

        # 4. Calcular o próximo código disponível
        # Se os códigos forem da forma Z001, Z002, Z030 -> extrai números
        numeric_suffixes = []
        pattern = re.compile(rf"^{re.escape(pref)}(\d+)$")

        for zt in matched_terms:
            m = pattern.match(zt)
            if m:
                numeric_suffixes.append(int(m.group(1)))

        if numeric_suffixes:
            max_num = max(numeric_suffixes)
            next_num = max_num + 1
            # Mantém a mesma formatação numérica de largura (ex: 3 dígitos Z031 ou 2 dígitos Z31)
            # Descobre o número de dígitos mais comum
            sample_digits = len(str(numeric_suffixes[0]))
            fmt = f"{{:0{max(sample_digits, 3)}d}}"
            next_code = f"{pref}{fmt.format(next_num)}"
            if len(next_code) > 4:
                next_code = next_code[:4]
        else:
            next_code = f"{pref}001"

        # Organiza amostra dos termos encontrados
        sample_list = []
        for zt in matched_terms:
            item = zterm_map.get(zt, {})
            sample_list.append({
                "zterm": zt,
                "text": item.get("text", ""),
                "ztag1": item.get("ztag1", "0"),
                "zprz1": item.get("zprz1", "0.000"),
            })

        return {
            "success": True,
            "environment": environment,
            "prefix": pref,
            "total_existing": len(existing_zterms),
            "matched_count": len(matched_terms),
            "last_used_code": matched_terms[-1] if matched_terms else None,
            "next_available_code": next_code,
            "items": sample_list,
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
            "error_code": code,
            "error_message": msg,
            "detail": str(exc),
        }
