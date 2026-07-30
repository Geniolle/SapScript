from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Callable

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

try:
    from .authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table
except ImportError:
    from authorization_rfc_analysis import _open_rfc_connection, _read_rfc_table


def _clean_str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def format_sap_user_id(user_val: str) -> str:
    if not user_val:
        return ""
    clean = str(user_val).strip().upper()
    if not clean:
        return ""
    if clean.startswith("S"):
        return clean
    if clean.isdigit():
        clean = clean.lstrip("0")
        return f"S{clean}"
    return f"S{clean}"


def search_hr_user_data_rfc(
    query: str,
    target_system_key: str = "S4PCLNT100",
    max_results: int = 10,
    progress_logger: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """
    Pesquisa e extrai dados de colaboradores na tabela de RH (ou master data de utilizadores)
    no sistema produtivo (S4PCLNT100).

    Tabelas consultadas:
    1. PA0002 (Dados pessoais: Nome, Apelido, PERNR)
    2. PA0001 (Atribuição org: Equipa/Unidade Org)
    3. PA0105 (Comunicação: Email no SUBTY 0010 e Utilizador SAP no SUBTY 0001)
    4. Fallback USR21 + ADRP + ADR6 + USR02 (Caso PA* não devoluva registos)
    """
    clean_query = _clean_str(query).upper()
    if not clean_query:
        return {
            "success": False,
            "message": "Termo de pesquisa vazio.",
            "data": [],
            "query": query,
            "total": 0,
        }

    results: list[dict[str, Any]] = []

    try:
        if callable(progress_logger):
            progress_logger(f"[RH RFC] A abrir ligação RFC a {target_system_key}...")

        conn = _open_rfc_connection(target_system_key)
    except Exception as exc:
        return {
            "success": False,
            "message": f"Falha na ligação RFC a {target_system_key}: {exc}",
            "data": [],
            "query": query,
            "total": 0,
        }

    try:
        # Tratar utilizadores com prefixo S (ex: S80001996 -> 80001996)
        clean_pernr = clean_query[1:].lstrip("0") if (clean_query.startswith("S") and clean_query[1:].isdigit()) else clean_query.lstrip("0")
        is_numeric = clean_pernr.isdigit()
        pa0002_filters = []
        if is_numeric:
            padded_pernr = clean_pernr.zfill(8)
            pa0002_filters.append({"field": "PERNR", "value": padded_pernr})
        
        if callable(progress_logger):
            progress_logger("[RH RFC] A consultar tabela de RH PA0002...")

        rows_pa0002: list[dict[str, str]] = []
        try:
            # DE/PARA Mapeado pelo utilizador:
            # CUA-NAME_FIRST -> PA0002-VORNA
            # CUA-NAME_LAST  -> PA0002-NACHN
            # CUA-SMTP_ADDR  -> PA0002-USRID_LONG / PA0105-USRID_LONG
            rows_pa0002 = _read_rfc_table(
                conn,
                "PA0002",
                ["PERNR", "VORNA", "NACHN", "CNAME"],
                pa0002_filters,
                max_rows=max_results,
            )
        except Exception:
            rows_pa0002 = []

        # Se encontrou no PA0002 por PERNR ou filtrando no cliente
        matching_pa0002 = []
        for r in rows_pa0002:
            pernr = _clean_str(r.get("PERNR"))
            vorna = _clean_str(r.get("VORNA"))  # CUA-NAME_FIRST
            nachn = _clean_str(r.get("NACHN"))  # CUA-NAME_LAST
            cname = _clean_str(r.get("CNAME")) or f"{vorna} {nachn}".strip()
            usrid_long_pa2 = _clean_str(r.get("USRID_LONG"))  # CUA-SMTP_ADDR (PA0002)

            if (
                is_numeric
                or clean_query in pernr
                or clean_query in vorna.upper()
                or clean_query in nachn.upper()
                or clean_query in cname.upper()
            ):
                matching_pa0002.append({
                    "pernr": pernr,
                    "first_name": vorna,
                    "last_name": nachn,
                    "full_name": cname,
                    "usrid_long_pa2": usrid_long_pa2,
                })

        # Se encontrou no RH (PA0002), cruzar com PA0001 (Equipa) e PA0105 (Email/User)
        if matching_pa0002:
            for item in matching_pa0002[:max_results]:
                pernr = item["pernr"]
                email = item.get("usrid_long_pa2") or ""  # Prioridade PA0002-USRID_LONG
                sap_user = ""
                team = ""

                # PA0105 - Email (CUA-SMTP_ADDR) e Utilizador SAP com filtros: SUBTY=0010 e ENDDA >= hoje (YYYYMMDD)
                try:
                    today_str = datetime.now().strftime("%Y%m%d")
                    rows_pa0105 = _read_rfc_table(
                        conn,
                        "PA0105",
                        ["PERNR", "SUBTY", "USRID", "USRID_LONG", "BEGDA", "ENDDA"],
                        [{"field": "PERNR", "value": pernr}],
                        max_rows=50,
                    )
                    active_emails = []
                    active_users = []
                    for r105 in rows_pa0105:
                        subty = _clean_str(r105.get("SUBTY"))
                        endda = _clean_str(r105.get("ENDDA")) or "99991231"
                        begda = _clean_str(r105.get("BEGDA")) or "19000101"
                        # Filtro: ENDDA maior ou igual a hoje (YYYYMMDD)
                        if endda >= today_str:
                            if subty == "0010":
                                val_email = _clean_str(r105.get("USRID_LONG")) or _clean_str(r105.get("USRID"))
                                if val_email:
                                    active_emails.append((begda, endda, val_email))
                            elif subty == "0001":
                                val_user = _clean_str(r105.get("USRID"))
                                if val_user:
                                    active_users.append((begda, endda, val_user))

                    if active_emails:
                        active_emails.sort(key=lambda x: x[0], reverse=True)
                        email = active_emails[0][2]

                    if active_users:
                        active_users.sort(key=lambda x: x[0], reverse=True)
                        sap_user = active_users[0][2]
                except Exception:
                    pass

                # PA0001 - Equipa / Unidade Org
                try:
                    rows_pa0001 = _read_rfc_table(
                        conn,
                        "PA0001",
                        ["PERNR", "ORGEH", "PLSTX", "STELL"],
                        [{"field": "PERNR", "value": pernr}],
                        max_rows=5,
                    )
                    if rows_pa0001:
                        team = _clean_str(rows_pa0001[0].get("PLSTX")) or _clean_str(rows_pa0001[0].get("ORGEH"))
                except Exception:
                    pass

                results.append({
                    "pernr": pernr,
                    "user_id": format_sap_user_id(sap_user or pernr),
                    "full_name": item["full_name"],
                    "first_name": item["first_name"],
                    "last_name": item["last_name"],
                    "email": email,
                    "team": team or "Recursos Humanos / Operacional",
                    "system": "S4P",
                    "source": "PA_HR_TABLES",
                })

        # 2. Se não encontrou dados no PA0002 ou a tabela PA* estiver vazia, consultar master data de utilizadores (USR21, ADRP, ADR6, USR02)
        if not results:
            if callable(progress_logger):
                progress_logger("[RH RFC] A consultar utilizadores em USR21 / USR02...")

            usr21_filters = []
            if not is_numeric:
                usr21_filters.append({"field": "BNAME", "value": clean_query})

            rows_usr21: list[dict[str, str]] = []
            try:
                rows_usr21 = _read_rfc_table(
                    conn,
                    "USR21",
                    ["BNAME", "PERSNUMBER", "ADDRNUMBER"],
                    usr21_filters,
                    max_rows=max_results,
                )
            except Exception:
                rows_usr21 = []

            # Se a busca exata por BNAME não deu resultados e a query não é puramente numérica, buscar USR02 para listar utilizadores
            if not rows_usr21 and not is_numeric:
                try:
                    rows_usr02 = _read_rfc_table(
                        conn,
                        "USR02",
                        ["BNAME", "CLASS", "USTYP"],
                        [],
                        max_rows=50,
                    )
                    matching_users = [
                        r.get("BNAME") for r in rows_usr02
                        if clean_query in _clean_str(r.get("BNAME")).upper()
                    ]
                    for u in matching_users[:max_results]:
                        u_rows = _read_rfc_table(
                            conn,
                            "USR21",
                            ["BNAME", "PERSNUMBER", "ADDRNUMBER"],
                            [{"field": "BNAME", "value": u}],
                            max_rows=1,
                        )
                        rows_usr21.extend(u_rows)
                except Exception:
                    pass

            for u_row in rows_usr21[:max_results]:
                bname = _clean_str(u_row.get("BNAME"))
                persnum = _clean_str(u_row.get("PERSNUMBER"))
                addrnum = _clean_str(u_row.get("ADDRNUMBER"))

                full_name = bname
                first_name = ""
                last_name = ""
                email = ""
                team = ""

                # ADRP - Nome
                if persnum:
                    try:
                        rows_adrp = _read_rfc_table(
                            conn,
                            "ADRP",
                            ["PERSNUMBER", "NAME_TEXT", "NAME_FIRST", "NAME_LAST"],
                            [{"field": "PERSNUMBER", "value": persnum}],
                            max_rows=1,
                        )
                        if rows_adrp:
                            full_name = _clean_str(rows_adrp[0].get("NAME_TEXT")) or full_name
                            first_name = _clean_str(rows_adrp[0].get("NAME_FIRST"))
                            last_name = _clean_str(rows_adrp[0].get("NAME_LAST"))
                    except Exception:
                        pass

                # ADR6 - Email
                if persnum or addrnum:
                    try:
                        adr6_filters = []
                        if persnum:
                            adr6_filters.append({"field": "PERSNUMBER", "value": persnum})
                        elif addrnum:
                            adr6_filters.append({"field": "ADDRNUMBER", "value": addrnum})

                        rows_adr6 = _read_rfc_table(
                            conn,
                            "ADR6",
                            ["ADDRNUMBER", "PERSNUMBER", "SMTP_ADDR"],
                            adr6_filters,
                            max_rows=1,
                        )
                        if rows_adr6:
                            email = _clean_str(rows_adr6[0].get("SMTP_ADDR"))
                    except Exception:
                        pass

                # USR02 - Classe / Departamento
                try:
                    rows_u02 = _read_rfc_table(
                        conn,
                        "USR02",
                        ["BNAME", "CLASS"],
                        [{"field": "BNAME", "value": bname}],
                        max_rows=1,
                    )
                    if rows_u02:
                        team = _clean_str(rows_u02[0].get("CLASS"))
                except Exception:
                    pass

                results.append({
                    "pernr": persnum or bname,
                    "user_id": format_sap_user_id(bname),
                    "full_name": full_name,
                    "first_name": first_name,
                    "last_name": last_name,
                    "email": email,
                    "team": team or "Geral",
                    "system": "S4P",
                    "source": "SAP_USER_MASTER",
                })

        return {
            "success": True,
            "message": f"Consulta RH concluída. Encontrados {len(results)} registo(s).",
            "data": results,
            "query": query,
            "total": len(results),
        }

    except Exception as exc:
        return {
            "success": False,
            "message": f"Erro ao processar consulta de RH: {exc}",
            "data": [],
            "query": query,
            "total": 0,
        }
    finally:
        try:
            conn.close()
        except Exception:
            pass
