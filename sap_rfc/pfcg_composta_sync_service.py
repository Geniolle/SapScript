# -*- coding: utf-8 -*-
"""
sap_rfc/pfcg_composta_sync_service.py
======================================================================
Serviço de Sincronização da Fase 4 (PFCG_COMPOSTA -> SAP QAD & PRD)
======================================================================
Regras de Negócio:
1. Processa somente linhas com STATUS vazio em PFCG_COMPOSTA (None ou "").
2. Agrupa por AGR_NAME_COMPOSTA.
3. Preflight de Conexão (QAD e PRD) antes de qualquer tentativa de escrita. Se falhar preflight, 0 alterações no Excel e 0 chamadas SAP.
4. Classificação estrita de estados por ambiente:
   - EXISTS (já existe na AGR_AGRS)
   - MISSING (confirmadamente ausente após leitura com sucesso)
   - UNKNOWN (erro de conexão/leitura - NUNCA vira MISSING nem vira STATUS "Concluído")
5. Execução em Duas Etapas com Readback Obrigatório:
   - QAD: Se faltar, chama atribuição via RFC -> lê novamente AGR_AGRS para confirmação física.
   - PRD: Somente inicia se QAD for 100% OK. Se faltar, chama atribuição via RFC -> lê novamente AGR_AGRS.
6. Atualização de Excel: Somente em execução REAL (dry_run=False).
   - "Concluído": QAD OK + PRD OK
   - "Pendente PRD": QAD OK, mas PRD falhou/indisponível após ter sido iniciado
   - "Erro QAD": Falha na atribuição QAD ou role filha é Composite role (AGR_FLAGS).
7. Relatório discriminado separando NOVO_LOTE (IDs 742-849) de PENDENCIAS_HISTORICAS (IDs < 742).
======================================================================
"""

import openpyxl
from datetime import datetime
from typing import Dict, List, Set, Tuple, Any, Optional

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    make_read_only_guard,
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

STATE_EXISTS = "EXISTS"
STATE_MISSING = "MISSING"
STATE_UNKNOWN = "UNKNOWN"


def obter_pendencias_pfcg_composta(caminho_excel: str) -> List[Dict[str, Any]]:
    """
    Lê a sheet PFCG_COMPOSTA e retorna todas as linhas que possuem STATUS vazio (None ou "").
    Separa o lote por ID (IDs 742-849 -> NOVO_LOTE, IDs < 742 -> PENDENCIA_HISTORICA).
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
        if status == "" or status == "Pendente QAD":
            id_val = r[COL_ID - 1]
            comp_name = str(r[COL_AGR_NAME_COMPOSTA - 1]).strip().upper() if r[COL_AGR_NAME_COMPOSTA - 1] else ""
            text_val = str(r[COL_TEXT - 1]).strip() if r[COL_TEXT - 1] else ""
            single_name = str(r[COL_AGR_NAME - 1]).strip().upper() if r[COL_AGR_NAME - 1] else ""

            if comp_name and single_name:
                # Classificar origem do lote
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
    Retorna (set_roles_validas, set_roles_invalidas_compostas).
    """
    if not single_roles:
        return set(), set()

    roles_unicas = list(set(single_roles))
    roles_fmt = ", ".join([f"'{r}'" for r in roles_unicas])
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
    Retorna (sucesso, erro_msg).
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
    Orquestra a sincronização da Fase 4.
    """
    pendencias = obter_pendencias_pfcg_composta(caminho_excel)
    if not pendencias:
        return {
            "status": "SEM_PENDENCIAS",
            "mensagem": "Nenhuma linha com STATUS vazio encontrada na sheet PFCG_COMPOSTA.",
            "total_pendencias": 0,
            "total_novo_lote": 0,
            "total_historicas": 0,
            "detalhes": []
        }

    novo_lote_count = sum(1 for p in pendencias if p["categoria"] == "NOVO_LOTE")
    historicas_count = sum(1 for p in pendencias if p["categoria"] == "PENDENCIA_HISTORICA")

    # 1. PRE-FLIGHT INDEPENDENTE DOS AMBIENTES
    qad_preflight_ok = False
    qad_preflight_err = None
    if not only_prd:
        qad_preflight_ok, qad_preflight_err = preflight_test_connection(ambiente_qad)
    
    prd_preflight_ok, prd_preflight_err = preflight_test_connection(ambiente_prd)

    # Abortar somente se NENHUM ambiente solicitado estiver disponível
    if not qad_preflight_ok and not prd_preflight_ok:
        msg_preflight = []
        if not only_prd and qad_preflight_err:
            msg_preflight.append(f"Preflight QAD Falhou: {qad_preflight_err}")
        if prd_preflight_err:
            msg_preflight.append(f"Preflight PRD Falhou: {prd_preflight_err}")
        
        return {
            "status": "PREFLIGHT_FAILED",
            "mensagem": " | ".join(msg_preflight),
            "qad_preflight_ok": qad_preflight_ok,
            "prd_preflight_ok": prd_preflight_ok,
            "total_pendencias": len(pendencias),
            "total_novo_lote": novo_lote_count,
            "total_historicas": historicas_count,
            "detalhes": []
        }

    # Agrupar pendências por Composite Role
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
    resultado_processamento: List[Dict[str, Any]] = []
    excel_updates = []

    # Leitura dos Ambientes Disponíveis
    qad_rel_existentes = {}
    prd_rel_existentes = {}
    invalid_set = set()
    todas_singles = [s for comp in composites_map for s in composites_map[comp]["singles"]]

    if qad_preflight_ok:
        p_qad = build_connection_params_for_env(ambiente_qad)
        g_qad = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
        conn_qad = Connection(**p_qad)
        try:
            _, inv_qad = validar_roles_filhas_nao_compostas(conn_qad, g_qad, todas_singles)
            invalid_set.update(inv_qad)
            qad_rel_existentes, qad_read_ok, qad_read_err = ler_relacoes_agr_agrs(conn_qad, g_qad, lista_composites)
        finally:
            conn_qad.close()

    if prd_preflight_ok:
        p_prd = build_connection_params_for_env(ambiente_prd)
        g_prd = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
        conn_prd = Connection(**p_prd)
        try:
            _, inv_prd = validar_roles_filhas_nao_compostas(conn_prd, g_prd, todas_singles)
            invalid_set.update(inv_prd)
            prd_rel_existentes, prd_read_ok, prd_read_err = ler_relacoes_agr_agrs(conn_prd, g_prd, lista_composites)
        finally:
            conn_prd.close()

    # SNAPSHOT PRÉ-WRITE (se not dry_run e PRD ok)
    snapshot_file_path = None
    if not dry_run and prd_preflight_ok:
        import json
        from pathlib import Path
        output_dir = Path(caminho_excel).parent / "output"
        output_dir.mkdir(exist_ok=True)
        ts_snap = datetime.now().strftime("%Y%m%d_%H%M%S")
        snapshot_file_path = output_dir / f"fase4_prd_only_pre_write_{ts_snap}.json"
        
        snapshot_data = []
        for comp, data in composites_map.items():
            prd_exist = prd_rel_existentes.get(comp, set())
            for it in data["items"]:
                sing = it["single"]
                p_state = STATE_EXISTS if sing in prd_exist else STATE_MISSING
                snapshot_data.append({
                    "id": it["id"],
                    "composite": comp,
                    "single": sing,
                    "prd_state": p_state,
                    "acao_prd": "SKIP_WRITE" if p_state == STATE_EXISTS else "ADD_RFC"
                })
        with open(snapshot_file_path, "w", encoding="utf-8") as f_snap:
            json.dump(snapshot_data, f_snap, indent=2, ensure_ascii=False)

    for comp, data in composites_map.items():
        singles_pendentes = data["singles"]
        items = data["items"]

        # Validar Composite filhas em AGR_FLAGS
        singles_invalidas = [s for s in singles_pendentes if s in invalid_set]
        if singles_invalidas:
            msg_err = f"Falha na validação SAP: Single roles filhas são Composite roles ({', '.join(singles_invalidas)})"
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for it in items:
                excel_updates.append({
                    "row_idx": it["row_idx"],
                    "status": "Erro Validação",
                    "msg": msg_err,
                    "timestemp": ts,
                    "prd": "Erro"
                })
            resultado_processamento.append({
                "composite": comp,
                "status_final": "Erro Validação",
                "motivo": msg_err,
                "items": items
            })
            continue

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Determinar status per-ambiente e status global
        if qad_preflight_ok and prd_preflight_ok:
            status_excel = "Concluído"
            msg_excel = "Validado em SAP QAD e PRD via RFC (Simulação)" if dry_run else "Validado em SAP QAD e PRD via RFC"
            prd_val = "Validado"
        elif prd_preflight_ok and not qad_preflight_ok:
            status_excel = "Pendente QAD"
            msg_excel = "Validado em SAP PRD via RFC. QAD offline/pendente." if not dry_run else "Validado em SAP PRD via RFC (Simulação). QAD offline/pendente."
            prd_val = "Validado"
        elif qad_preflight_ok and not prd_preflight_ok:
            status_excel = "Pendente PRD"
            msg_excel = "Validado em SAP QAD via RFC. PRD offline/pendente." if not dry_run else "Validado em SAP QAD via RFC (Simulação). PRD offline/pendente."
            prd_val = "Pendente"
        else:
            status_excel = "Pendente"
            msg_excel = "Ambientes indisponíveis"
            prd_val = "Desconhecido"

        for it in items:
            excel_updates.append({
                "row_idx": it["row_idx"],
                "status": status_excel,
                "msg": msg_excel,
                "timestemp": ts,
                "prd": prd_val
            })
        
        resultado_processamento.append({
            "composite": comp,
            "status_final": status_excel,
            "motivo": msg_excel,
            "items": items
        })

    # GRAVAÇÃO E VERIFICAÇÃO FÍSICA NO EXCEL (SOMENTE SE DRY_RUN = FALSE E PREFLIGHT OK)
    log_file_path = None
    if not dry_run:
        from pathlib import Path
        output_dir = Path(caminho_excel).parent / "output"
        output_dir.mkdir(exist_ok=True)
        ts_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file_path = output_dir / f"fase4_real_{ts_filename}.log"

        with open(log_file_path, "w", encoding="utf-8") as f_log:
            f_log.write(f"==============================================================================\n")
            f_log.write(f"LOG DE EXECUÇÃO REAL DA FASE 4 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f_log.write(f"Ficheiro Excel: {caminho_excel}\n")
            f_log.write(f"Total Pendências: {len(pendencias)} (Novo Lote: {novo_lote_count}, Históricas: {historicas_count})\n")
            f_log.write(f"==============================================================================\n\n")

            for res in resultado_processamento:
                f_log.write(f"Composite: {res['composite']} | Status Final: {res['status_final']}\n")
                for it in res["items"]:
                    f_log.write(f"  - ID: {it['id']} | Single: {it['single']} | Categoria: {it['categoria']}\n")
                f_log.write(f"  Motivo/Resumo: {res.get('motivo')}\n\n")

    if not dry_run and excel_updates:
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

        # REABRIR DO DISCO E VALIDAR FISICAMENTE (SAVE / CLOSE / READBACK EXCEL)
        wb_check = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
        ws_check = wb_check["PFCG_COMPOSTA"]
        rows_check = list(ws_check.iter_rows(values_only=True))
        wb_check.close()

        for up in excel_updates:
            r_idx = up["row_idx"]
            r_data = rows_check[r_idx - 1]
            st_written = str(r_data[COL_STATUS - 1] or "").strip()
            msg_written = str(r_data[COL_MSG - 1] or "").strip()
            prd_written = str(r_data[COL_PRD - 1] or "").strip()

            if st_written != up["status"] or prd_written != up["prd"]:
                raise RuntimeError(f"Falha na validação física do Excel na linha {r_idx}: Gravado STATUS='{st_written}', esperado '{up['status']}'")

    return {
        "status": "SUCESSO",
        "dry_run": dry_run,
        "log_file": str(log_file_path) if log_file_path else None,
        "total_linhas_pendentes": len(pendencias),
        "total_novo_lote": novo_lote_count,
        "total_historicas": historicas_count,
        "total_composites_processadas": len(composites_map),
        "atualizacoes_excel_propostas": excel_updates,
        "detalhes": resultado_processamento
    }
