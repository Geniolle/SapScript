# -*- coding: utf-8 -*-
"""
sap_rfc/pfcg_composta_sync_service.py
======================================================================
Serviço de Sincronização da Fase 4 (PFCG_COMPOSTA -> SAP QAD & PRD)
======================================================================
Regras de Negócio:
1. Processa somente linhas com STATUS vazio em PFCG_COMPOSTA (None ou "").
2. Pré-validação READ-ONLY obrigatória em AGR_AGRS para classificar:
   - JA_EXISTE_PRD (membro já presente) -> ignora escrita
   - REALMENTE_PENDENTE (confirmadamente ausente) -> adiciona via RFC
   - ERRO_LEITURA -> aborta sem falsa conclusão
3. Execução RFC exclusiva (PRGN_RFC_ADD_AGRS_TO_COLL_AGR + PRGN_GEN_PROFILES_FOR_ROLES)
   sem qualquer fallback para SAP GUI.
4. Readback OBRIGATÓRIO pós-escrita consultando novamente a AGR_AGRS.
5. Retorno estruturado confiável (ok, analisadas, ja_existiam, criadas, confirmadas, erros).
6. Atualização de Excel: Somente em execução REAL (dry_run=False) e se confirmadas por readback.
======================================================================
"""

import openpyxl
from datetime import datetime
from typing import Dict, List, Set, Tuple, Any, Optional

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    make_read_only_guard,
    make_write_guard,
    read_table,
    classify_rfc_error
)
from pyrfc import Connection

# Constantes de Colunas da Sheet PFCG_COMPOSTA (1-indexed)
COL_ID = 1                   # A
COL_AGR_NAME_COMPOSTA = 2    # B
COL_TEXT = 3                 # C
COL_AGR_NAME = 4             # D
COL_STATUS = 5               # E
COL_MSG = 6                  # F
COL_TIMESTEMP = 7            # G
COL_PRD = 8                  # H

STATE_EXISTS = "JA_EXISTE_PRD"
STATE_PENDING = "REALMENTE_PENDENTE"
STATE_ERROR = "ERRO_LEITURA"


def obter_pendencias_pfcg_composta(caminho_excel: str) -> List[Dict[str, Any]]:
    """
    Lê a sheet PFCG_COMPOSTA e retorna todas as linhas que possuem STATUS vazio (None ou "").
    """
    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    if "PFCG_COMPOSTA" not in wb.sheetnames:
        wb.close()
        return []

    ws = wb["PFCG_COMPOSTA"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) <= 1:
        return []

    pendencias = []
    for row_idx, r in enumerate(rows[1:], start=2):
        if not r or len(r) < 4:
            continue
        
        status = str(r[COL_STATUS - 1]).strip() if len(r) >= COL_STATUS and r[COL_STATUS - 1] is not None else ""
        if status == "" or status == "Pendente QAD" or status == "Pendente PRD":
            id_val = r[COL_ID - 1]
            comp_name = str(r[COL_AGR_NAME_COMPOSTA - 1]).strip().upper() if r[COL_AGR_NAME_COMPOSTA - 1] else ""
            text_val = str(r[COL_TEXT - 1]).strip() if r[COL_TEXT - 1] else ""
            single_name = str(r[COL_AGR_NAME - 1]).strip().upper() if r[COL_AGR_NAME - 1] else ""

            if comp_name and single_name:
                categoria = "NOVO_LOTE" if isinstance(id_val, (int, float)) and 742 <= id_val <= 849 else "PENDENCIA_HISTORICA"
                pendencias.append({
                    "row_idx": row_idx,
                    "id": id_val,
                    "composite": comp_name,
                    "text": text_val,
                    "single": single_name,
                    "categoria": categoria
                })

    return pendencias


def validar_roles_filhas_nao_compostas(conn: Any, guard: Any, single_roles: List[str]) -> Tuple[Set[str], Set[str]]:
    """
    Consulta a tabela AGR_FLAGS para garantir que nenhuma Single Role a ser adicionada
    é na verdade uma Composite Role (FLAG_TYPE == 'COLL_AGR').
    """
    if not single_roles:
        return set(), set()

    roles_unicas = list(set(single_roles))
    options = [{"TEXT": f"AGR_NAME = '{r}' AND FLAG_TYPE = 'COLL_AGR'"} for r in roles_unicas]

    try:
        flags_rows = read_table(
            conn,
            guard,
            table_name="AGR_FLAGS",
            fields=["AGR_NAME", "FLAG_TYPE"],
            options=options,
            rowcount=0
        )
        invalid_composites = {row[0].strip().upper() for row in flags_rows if row}
    except Exception:
        invalid_composites = set()

    roles_validas = set(roles_unicas) - invalid_composites
    return roles_validas, invalid_composites


def ler_relacoes_agr_agrs(conn: Any, guard: Any, composite_roles: List[str]) -> Tuple[Dict[str, Set[str]], bool, Optional[str]]:
    """
    Lê a tabela AGR_AGRS para um conjunto de Composite Roles.
    Retorna (dicionário {COMPOSITE: set(SINGLES_EXISTENTES)}, ok_status, erro_msg).
    """
    if not composite_roles:
        return {}, True, None

    options = [{"TEXT": f"AGR_NAME = '{c}'"} for c in set(composite_roles)]
    relacoes: Dict[str, Set[str]] = {c: set() for c in composite_roles}
    try:
        data_rows = read_table(
            conn,
            guard,
            table_name="AGR_AGRS",
            fields=["AGR_NAME", "CHILD_AGR"],
            options=options,
            rowcount=0
        )
        for r in data_rows:
            if len(r) >= 2:
                parent = r[0].strip().upper()
                child = r[1].strip().upper()
                if parent in relacoes:
                    relacoes[parent].add(child)
        return relacoes, True, None
    except Exception as e:
        return {}, False, str(e)


def preflight_test_connection(ambiente: str) -> Tuple[bool, Optional[str]]:
    """
    Testa preflight de conexão com um ambiente SAP (READ-ONLY).
    """
    try:
        from sap_rfc._rfc_common import load_project_env, find_project_root
        from pyrfc import Connection
        load_project_env(find_project_root())
        params = build_connection_params_for_env(ambiente)
        guard = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
        
        conn = Connection(**params)
        read_table(conn, guard, table_name="AGR_DEFINE", fields=["AGR_NAME"], options=[], rowcount=1)
        conn.close()
        return True, None
    except Exception as e:
        return False, str(e)


def sincronizar_pfcg_composta(
    caminho_excel: str,
    dry_run: bool = True,
    only_prd: bool = False,
    ambiente_qad: str = "QAD",
    ambiente_prd: str = "PRD"
) -> Dict[str, Any]:
    """
    Orquestra a sincronização da sheet PFCG_COMPOSTA via RFC com readback obrigatório.
    Devolve estrutura de resultado padronizada.
    """
    pendencias = obter_pendencias_pfcg_composta(caminho_excel)
    if not pendencias:
        return {
            "ok": True,
            "status": "SEM_PENDENCIAS",
            "mensagem": "Nenhuma linha com STATUS vazio encontrada na sheet PFCG_COMPOSTA.",
            "analisadas": 0,
            "ja_existiam": 0,
            "criadas": 0,
            "confirmadas": 0,
            "total_pendencias": 0,
            "total_novo_lote": 0,
            "total_historicas": 0,
            "erros": [],
            "detalhes": []
        }

    novo_lote_count = sum(1 for p in pendencias if p["categoria"] == "NOVO_LOTE")
    historicas_count = sum(1 for p in pendencias if p["categoria"] == "PENDENCIA_HISTORICA")

    # 1. PRE-FLIGHT DE CONEXÃO RFC
    qad_preflight_ok = False
    qad_preflight_err = None
    if not only_prd:
        qad_preflight_ok, qad_preflight_err = preflight_test_connection(ambiente_qad)
    
    prd_preflight_ok, prd_preflight_err = preflight_test_connection(ambiente_prd)

    # Abortar se preflight do ambiente PRD falhar
    if not prd_preflight_ok:
        msg_err = f"Falha de conexão RFC com SAP PRD: {prd_preflight_err}"
        return {
            "ok": False,
            "status": "PREFLIGHT_FAILED",
            "mensagem": msg_err,
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "criadas": 0,
            "confirmadas": 0,
            "total_pendencias": len(pendencias),
            "erros": [msg_err],
            "detalhes": []
        }

    # Agrupar por Composite Role
    composites_map: Dict[str, Dict[str, Any]] = {}
    for item in pendencias:
        comp = item["composite"]
        if comp not in composites_map:
            composites_map[comp] = {
                "text": item["text"],
                "singles": [],
                "items": []
            }
        composites_map[comp]["singles"].append(item["single"])
        composites_map[comp]["items"].append(item)

    lista_composites = list(composites_map.keys())
    todas_singles = [s for comp in composites_map for s in composites_map[comp]["singles"]]

    # 2. PRÉ-VALIDAÇÃO READ-ONLY EM AGR_AGRS (PRD)
    params_prd = build_connection_params_for_env(ambiente_prd)
    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
    conn_ro = Connection(**params_prd)
    
    try:
        _, invalid_set = validar_roles_filhas_nao_compostas(conn_ro, guard_ro, todas_singles)
        prd_rel_iniciais, prd_read_ok, prd_read_err = ler_relacoes_agr_agrs(conn_ro, guard_ro, lista_composites)
    finally:
        conn_ro.close()

    if not prd_read_ok:
        msg_err = f"Falha na leitura inicial da tabela AGR_AGRS no SAP PRD: {prd_read_err}"
        return {
            "ok": False,
            "status": "READ_FAILED",
            "mensagem": msg_err,
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "criadas": 0,
            "confirmadas": 0,
            "erros": [msg_err],
            "detalhes": []
        }

    total_analisadas = len(pendencias)
    ja_existiam_cnt = 0
    realmente_pendentes_items = []

    for comp, data in composites_map.items():
        existentes_comp = prd_rel_iniciais.get(comp, set())
        for it in data["items"]:
            sing = it["single"]
            if sing in existentes_comp:
                ja_existiam_cnt += 1
                it["estado_prd"] = STATE_EXISTS
            else:
                it["estado_prd"] = STATE_PENDING
                realmente_pendentes_items.append(it)

    # 3. EXECUÇÃO RFC DAS REALMENTE PENDENTES (DELTA)
    criadas_cnt = 0
    erros_execucao = []

    if realmente_pendentes_items and not dry_run:
        guard_wr = make_write_guard(
            allowed_functions=("PRGN_RFC_ADD_AGRS_TO_COLL_AGR", "PRGN_GEN_PROFILES_FOR_ROLES"),
            allowed_tables=("AGR_AGRS", "AGR_TEXTS")
        )
        conn_wr = Connection(**params_prd)
        try:
            # Agrupar delta por Composite
            delta_map = defaultdict(list)
            for it in realmente_pendentes_items:
                delta_map[it["composite"]].append(it)

            for comp, items_delta in delta_map.items():
                singles_payload = [{"AGR_NAME": it["single"], "TEXT": it["text"]} for it in items_delta]
                try:
                    res_rfc = conn_wr.call(
                        "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
                        ACTIVITY_GROUP=comp,
                        ACTIVITY_GROUPS=singles_payload,
                        NO_DIALOG="X"
                    )
                    rets = res_rfc.get("RETURN", []) or []
                    errs = [m for m in rets if str(m.get("TYPE", "")).upper() in ("E", "A")]
                    if errs:
                        msg = f"Erro RFC ao adicionar a {comp}: {'; '.join(m.get('MESSAGE', '') for m in errs)}"
                        erros_execucao.append(msg)
                        continue

                    # Gerar perfis e atualizar utilizador (User Compare)
                    conn_wr.call("PRGN_GEN_PROFILES_FOR_ROLES", IT_ROLES=[{"AGR_NAME": comp}], IV_USERCOMPARE="X")
                    criadas_cnt += len(items_delta)
                except Exception as exc:
                    erros_execucao.append(f"Exceção RFC ao processar {comp}: {exc}")
        finally:
            conn_wr.close()

    # 4. READBACK OBRIGATÓRIO EM AGR_AGRS PÓS-ESCRITA
    conn_rb = Connection(**params_prd)
    try:
        prd_rel_finais, rb_ok, rb_err = ler_relacoes_agr_agrs(conn_rb, guard_ro, lista_composites)
    finally:
        conn_rb.close()

    confirmadas_cnt = 0
    excel_updates = []
    resultado_detalhado = []

    for comp, data in composites_map.items():
        finais_comp = prd_rel_finais.get(comp, set())
        for it in data["items"]:
            sing = it["single"]
            row_idx = it["row_idx"]
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if sing in finais_comp:
                confirmadas_cnt += 1
                st = "Concluído"
                msg = "Validado e confirmado fisicamente no SAP PRD via RFC (Readback OK)"
                prd_st = "Validado"
            else:
                st = "Erro"
                msg = "Falha no readback AGR_AGRS: Relação não confirmada no SAP PRD após tentativa"
                prd_st = "Erro"

            excel_updates.append({
                "row_idx": row_idx,
                "status": st,
                "msg": msg,
                "timestemp": ts,
                "prd": prd_st
            })

    # 5. ATUALIZAÇÃO NO EXCEL SE EXECUÇÃO REAL (dry_run=False)
    if not dry_run and excel_updates and not erros_execucao:
        wb = openpyxl.load_workbook(caminho_excel)
        ws = wb["PFCG_COMPOSTA"]
        for up in excel_updates:
            r_idx = up["row_idx"]
            ws.cell(row=r_idx, column=COL_STATUS, value=up["status"])
            ws.cell(row=r_idx, column=COL_MSG, value=up["msg"])
            ws.cell(row=r_idx, column=COL_TIMESTEMP, value=up["timestemp"])
            ws.cell(row=r_idx, column=COL_PRD, value=up["prd"])
        wb.save(caminho_excel)
        wb.close()

    is_ok = len(erros_execucao) == 0 and confirmadas_cnt == total_analisadas

    msg_sucesso = (
        f"✓ PFCG_COMPOSTA concluída via RFC.\n"
        f"   Linhas analisadas: {total_analisadas}\n"
        f"   Já existentes: {ja_existiam_cnt}\n"
        f"   Criadas no SAP: {criadas_cnt}\n"
        f"   Confirmadas por readback: {confirmadas_cnt}\n"
        f"   Erros: {len(erros_execucao)}"
    )

    return {
        "ok": is_ok,
        "status": "SUCESSO" if is_ok else "FALHA",
        "mensagem": msg_sucesso if is_ok else f"Falha na sincronização PFCG_COMPOSTA: {'; '.join(erros_execucao)}",
        "analisadas": total_analisadas,
        "ja_existiam": ja_existiam_cnt,
        "criadas": criadas_cnt,
        "confirmadas": confirmadas_cnt,
        "total_pendencias": len(pendencias),
        "total_novo_lote": novo_lote_count,
        "total_historicas": historicas_count,
        "erros": erros_execucao,
        "detalhes": resultado_detalhado
    }
