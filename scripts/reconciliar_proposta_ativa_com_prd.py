# -*- coding: utf-8 -*-
"""
scripts/reconciliar_proposta_ativa_com_prd.py
======================================================================
Script de Reconciliação Integrada de 4 Fases para o Projeto Perfil:
FASE 1: População de Utilizadores Ativos (CONTROLO -> Sheet Departamental -> User SAP).
FASE 2: Incorporation de Single Roles Diretas Válidas do SAP PRD (Read-Only) -> Proposta Ativa (aplicando filtro DEFINIÇÕES).
FASE 3A: Mapeamento de X das Matrizes Departamentais -> Proposta -> Proposta Ativa (Relação 1:N TCODE -> Roles).
FASE 3B: Alinhamento e União Aditiva de Composite Roles -> PFCG_COMPOSTA.
======================================================================
"""

import sys
import os
import re
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Set, Tuple, Any, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from sap_rfc._rfc_common import (
    build_connection_params_for, make_read_only_guard, read_table,
    load_project_env, find_project_root
)
from pyrfc import Connection
import openpyxl

from scripts.reconciliar_matrizes_com_prd import obter_departamentos_processados, extrair_user_id

CAMINHO_EXCEL_PADRAO = str(PROJECT_ROOT / "S4H_Perfis de autorização_v1.xlsx")


def obter_funcoes_excluidas_definicoes(caminho_excel: str) -> Set[str]:
    """Lê a sheet DEFINIÇÕES e extrai o conjunto único de roles que devem ser excluídas de Funções Individuais."""
    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    if "DEFINIÇÕES" not in wb.sheetnames:
        wb.close()
        return set()

    ws = wb["DEFINIÇÕES"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    funcoes_excluidas = set()
    if len(rows) > 1:
        for r in rows[1:]:
            for cell in r[1:]:  # Ignorar coluna DEPARTAMENTO (índice 0)
                if cell:
                    items = str(cell).split(",")
                    for it in items:
                        val = it.strip().upper()
                        if val:
                            funcoes_excluidas.add(val)

    return funcoes_excluidas


def carregar_mapeamento_tcode_para_funcoes_proposta(caminho_excel: str) -> Tuple[Dict[str, Set[str]], List[Dict[str, Any]]]:
    """
    Lê a sheet 'Proposta' do Excel e monta a relação 1:N TCODE -> Set de Funções Individuais.
    Também retorna lista de avisos para TCODEs sem mapeamento de função.
    """
    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    if "Proposta" not in wb.sheetnames:
        wb.close()
        return {}, []

    ws = wb["Proposta"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    tcode_map = {}
    current_role = None

    for r in rows[1:]:
        if not r:
            continue
        func = str(r[0]).strip().upper() if r[0] else ""
        desc = str(r[1]).strip() if len(r) > 1 and r[1] else ""

        if func.startswith("Z_") and len(func) >= 4:
            current_role = func
        elif func and current_role and "TRANSACAO NAO EXISTE" not in desc.upper():
            for t in re.split(r"[;, \s\n]+", func):
                tc = t.strip().upper()
                if tc:
                    if tc not in tcode_map:
                        tcode_map[tc] = set()
                    tcode_map[tc].add(current_role)

    return tcode_map, []


def executar_reconciliacao_integrada(caminho_excel: str, simular: bool = True, only_prd: bool = False) -> Dict[str, Any]:
    print(f"\n==============================================================================")
    print(f"  EXECUÇÃO DE RECONCILIAÇÃO INTEGRADA (FASE 1 -> FASE 2 -> FASE 3A -> FASE 3B -> FASE 4)")
    print(f"  Ficheiro: {caminho_excel}")
    print(f"  Modo PRD-Only: {'SIM' if only_prd else 'NÃO'}")
    print(f"  Modo Execução: {'SIMULAÇÃO (Somente Leitura)' if simular else 'EXECUÇÃO REAL'}")
    print(f"==============================================================================\n")

    # 0. Obter conjunto de exclusão da sheet DEFINIÇÕES e Mapeamento 1:N TCODE -> Roles da sheet Proposta
    funcoes_excluidas_definicoes = obter_funcoes_excluidas_definicoes(caminho_excel)
    tcode_to_roles_map, _ = carregar_mapeamento_tcode_para_funcoes_proposta(caminho_excel)

    print(f"[0/5] Pré-carregamento:")
    print(f"      - Roles excluídas em DEFINIÇÕES ({len(funcoes_excluidas_definicoes)})")
    print(f"      - Mapeamento 1:N TCODE -> Roles em 'Proposta' ({len(tcode_to_roles_map)} TCODEs mapeadas)\n")

    # FASE 1: USERS ATIVOS (CONTROLO -> Sheet Departamental -> User SAP)
    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    deps = obter_departamentos_processados(caminho_excel, incluir_pendencias=True, incluir_pendentes=True)
    print(f"[FASE 1] Departamentos elegíveis na sheet CONTROLO ({len(deps)}): {', '.join(deps)}")

    users_excel_info = []
    user_ids_set = set()
    users_pendentes_classificacao = []

    for d in deps:
        if d in wb.sheetnames:
            ws = wb[d]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) >= 2:
                header = rows[1]
                for col_idx in range(2, len(header)):
                    raw_val = str(header[col_idx] or "").strip()
                    uid = extrair_user_id(header[col_idx])
                    nome = raw_val.replace(uid, "").strip().replace("\n", " ") if uid else raw_val.replace("\n", " ")

                    if not uid:
                        # Classificar o motivo da ausência de User SAP
                        estado = "PENDENTE_CRIACAO_USER_SAP" if "NEW" in raw_val.upper() else "USER_SEM_ID_SAP"
                        users_pendentes_classificacao.append({
                            "dep": d, "uid": None, "nome": nome, "raw_cabecalho": raw_val,
                            "coluna": col_idx + 1, "estado": estado,
                            "motivo": "Utilizador rascunho sem ID SAP cadastrado na matriz departamental"
                        })
                    else:
                        users_excel_info.append({"dep": d, "uid": uid, "nome": nome, "coluna": col_idx + 1, "raw_cabecalho": raw_val})
                        user_ids_set.add(uid)

    # Mapeamento para a Proposta Ativa
    ws_prop = wb["Proposta Ativa"]
    rows_prop = list(ws_prop.iter_rows(values_only=True))

    proposta_ativa_map: Dict[str, int] = {}
    user_funcs_excel: Dict[str, Set[str]] = {}
    user_composite_excel: Dict[str, str] = {}

    for r_idx, r in enumerate(rows_prop[1:], start=2):
        if r[0]:
            uid_prop = str(r[0]).strip().upper()
            proposta_ativa_map[uid_prop] = r_idx

    users_mapeados_fase1 = []
    erros_mapeamento = []
    for item in users_excel_info:
        uid = item["uid"]
        if uid in proposta_ativa_map:
            linha_prop = proposta_ativa_map[uid]
            item["linha_proposta_ativa"] = linha_prop
            item["estado"] = "PROCESSAVEL"
            r = rows_prop[linha_prop - 1]
            comp_role = str(r[7]).strip().upper() if len(r) > 7 and r[7] else ""
            user_composite_excel[uid] = comp_role
            funcs = set()
            for cell in r[9:]:
                if cell is not None:
                    val = str(cell).strip().upper()
                    if val and val.startswith("Z") and not val.startswith("Z_BR_") and len(val) >= 4:
                        funcs.add(val)
            user_funcs_excel[uid] = funcs
            users_mapeados_fase1.append(item)
        else:
            item["linha_proposta_ativa"] = None
            item["estado"] = "USER_NAO_MAPEADO_PROPOSTA_ATIVA"
            item["motivo"] = f"Utilizador '{uid}' não possui linha correspondente na sheet Proposta Ativa"
            erros_mapeamento.append(item)
            users_pendentes_classificacao.append(item)

    print(f"      - População congelada: {len(users_excel_info)} Users ({len(users_mapeados_fase1)} mapeados na Proposta Ativa, {len(erros_mapeamento)} erros de mapeamento)\n")

    user_ids_list = sorted(list(user_ids_set))
    user_singles_prd: Dict[str, Set[str]] = {uid: set() for uid in user_ids_list}

    try:
        load_project_env()
        params = build_connection_params_for("PRD")
        guard = make_read_only_guard()
        conn = Connection(**params)
        today_str = datetime.now().strftime("%Y%m%d")

        chunk_size = 30
        for i in range(0, len(user_ids_list), chunk_size):
            chunk = user_ids_list[i : i + chunk_size]
            options = []
            for idx, u in enumerate(chunk):
                prefix = "OR " if idx > 0 else ""
                options.append({"TEXT": f"{prefix}UNAME = '{u}'"})

            rows_users = read_table(
                conn, guard, table_name="AGR_USERS",
                fields=["UNAME", "AGR_NAME", "FROM_DAT", "TO_DAT", "COL_FLAG"],
                options=options, rowcount=0
            )
            for r in rows_users:
                if len(r) >= 5:
                    uname, agr_name, f_dat, t_dat, col_flag = [str(x or "").strip() for x in r]
                    uname = uname.upper()
                    agr_name = agr_name.upper()
                    if uname in user_singles_prd and agr_name:
                        if (not f_dat or f_dat <= today_str) and (not t_dat or t_dat >= today_str):
                            if col_flag != "X" and not agr_name.startswith("Z_BR_"):
                                user_singles_prd[uname].add(agr_name)

        conn.close()
        print("      - Conexão SAP PRD efetuada com sucesso.")
    except Exception as e_rfc:
        print(f"      - [AVISO/VPN] Conexão SAP PRD não disponível offline ({e_rfc}). Simulação prossegue com dados locais.")

    # Cálculo da Fase 2
    fase2_adicoes_por_user = {}
    for item in users_mapeados_fase1:
        uid = item["uid"]
        funcs_ex = user_funcs_excel.get(uid, set())
        funcs_prd_diretas = user_singles_prd.get(uid, set())

        funcs_elegiveis = funcs_prd_diretas - funcoes_excluidas_definicoes
        a_adicionar = funcs_elegiveis - funcs_ex
        fase2_adicoes_por_user[uid] = a_adicionar

    total_adicionar_fase2 = sum(len(v) for v in fase2_adicoes_por_user.values())
    print(f"      - Concluída. Funções a adicionar pela Fase 2: {total_adicionar_fase2}\n")

    # FASE 3A: X DA MATRIZ -> PROPOSTA ATIVA (Usando strict população da Fase 1)
    print(f"[FASE 3A] Leitura dos X das matrizes departamentais e mapeamento 1:N -> Proposta Ativa...")
    fase3a_adicoes_por_user = {}
    tcodes_sem_mapeamento_list = []

    total_tcodes_com_x = 0
    total_casos_1n = 0

    for item in users_mapeados_fase1:
        dep = item["dep"]
        uid = item["uid"]
        col_idx = item["coluna"] - 1  # 0-based

        # Estado da Proposta Ativa já enriquecido pela Fase 2
        funcs_atuais_pos_fase2 = set(user_funcs_excel.get(uid, set())).union(fase2_adicoes_por_user.get(uid, set()))

        funcs_necessarias_matriz = set()
        if dep in wb.sheetnames:
            ws_dep = wb[dep]
            rows_dep = list(ws_dep.iter_rows(values_only=True))

            for r_idx_dep, row_dep in enumerate(rows_dep[2:], start=3):
                if len(row_dep) > col_idx and row_dep[col_idx]:
                    flag = str(row_dep[col_idx]).strip().upper()
                    if flag in ("X", "1", "SIM", "YES", "S"):
                        tcode = str(row_dep[0] or "").strip().upper()
                        if tcode:
                            total_tcodes_com_x += 1
                            roles_mapeadas = tcode_to_roles_map.get(tcode, set())
                            if not roles_mapeadas:
                                tcodes_sem_mapeamento_list.append({
                                    "dep": dep, "uid": uid, "nome": item["nome"],
                                    "tcode": tcode, "linha": r_idx_dep
                                })
                            else:
                                if len(roles_mapeadas) > 1:
                                    total_casos_1n += 1
                                funcs_necessarias_matriz.update(roles_mapeadas)

        # Adicionar apenas o que falta no estado pós-Fase 2
        a_adicionar_3a = funcs_necessarias_matriz - funcs_atuais_pos_fase2
        fase3a_adicoes_por_user[uid] = a_adicionar_3a

    total_adicionar_fase3a = sum(len(v) for v in fase3a_adicoes_por_user.values())
    print(f"      - Concluída. Total TCODEs com X: {total_tcodes_com_x}")
    print(f"      - TCODEs sem mapeamento em 'Proposta': {len(tcodes_sem_mapeamento_list)}")
    print(f"      - Casos 1:N TCODE -> Roles: {total_casos_1n}")
    print(f"      - Funções a adicionar pela Fase 3A: {total_adicionar_fase3a}\n")

    # FASE 3B: PROPOSTA ATIVA -> PFCG_COMPOSTA (Calculada sobre estado final)
    print(f"[FASE 3B] Alinhamento da PFCG_COMPOSTA por União Aditiva dos utilizadores...")

    # 1. Agrupar utilizadores e calcular união aditiva de funções por Composite
    composta_users: Dict[str, List[str]] = {}
    composta_user_funcs: Dict[str, Dict[str, Set[str]]] = {}
    composta_roles_esperadas: Dict[str, Set[str]] = {}

    for item in users_mapeados_fase1:
        uid = item["uid"]
        comp = user_composite_excel.get(uid, "")
        if comp and comp not in ("NAN", "NONE", "-", ""):
            funcs_finais_user = set(user_funcs_excel.get(uid, set())) \
                .union(fase2_adicoes_por_user.get(uid, set())) \
                .union(fase3a_adicoes_por_user.get(uid, set()))

            if comp not in composta_users:
                composta_users[comp] = []
                composta_user_funcs[comp] = {}
                composta_roles_esperadas[comp] = set()

            composta_users[comp].append(uid)
            composta_user_funcs[comp][uid] = funcs_finais_user
            composta_roles_esperadas[comp].update(funcs_finais_user)

    # 2. Ler estado atual da sheet PFCG_COMPOSTA
    ws_comp = wb["PFCG_COMPOSTA"]
    rows_comp = list(ws_comp.iter_rows(values_only=True))

    pfcg_composta_atual: Dict[str, Set[str]] = {}
    pfcg_composta_texts: Dict[str, str] = {}
    maior_id_existente = 0

    if len(rows_comp) > 1:
        for r in rows_comp[1:]:
            if not r:
                continue
            r_id = r[0]
            if r_id is not None:
                try:
                    r_id_num = int(r_id)
                    if r_id_num > maior_id_existente:
                        maior_id_existente = r_id_num
                except (ValueError, TypeError):
                    pass

            c_name = str(r[1]).strip().upper() if len(r) > 1 and r[1] else ""
            c_text = str(r[2]).strip() if len(r) > 2 and r[2] else ""
            s_name = str(r[3]).strip().upper() if len(r) > 3 and r[3] else ""

            if c_name:
                if c_name not in pfcg_composta_atual:
                    pfcg_composta_atual[c_name] = set()
                if s_name:
                    pfcg_composta_atual[c_name].add(s_name)

                if c_text and c_name not in pfcg_composta_texts:
                    pfcg_composta_texts[c_name] = c_text

    # 3. Processar deltas e gerar novas linhas para cada Composite
    detalhes_fase3b = []
    novas_linhas_pfcg = []
    proximo_id = maior_id_existente + 1

    total_relacoes_user_funcao = 0
    total_relacoes_comp_single_calc = 0
    total_ja_existente = 0
    total_novas_adicionar = 0
    total_duplicadas_evitadas = 0
    total_roles_nao_requeridas = 0
    total_composites_sem_text = 0
    total_composites_sem_linha_previa = 0

    for comp in sorted(list(composta_roles_esperadas.keys())):
        users_grupo = composta_users[comp]
        funcs_by_u = composta_user_funcs[comp]
        calc_set = composta_roles_esperadas[comp]
        atual_set = pfcg_composta_atual.get(comp, set())
        text_comp = pfcg_composta_texts.get(comp, "")

        for u in users_grupo:
            total_relacoes_user_funcao += len(funcs_by_u[u])

        total_relacoes_comp_single_calc += len(calc_set)

        if comp not in pfcg_composta_atual:
            total_composites_sem_linha_previa += 1

        if not text_comp:
            total_composites_sem_text += 1

        roles_a_adicionar = sorted(list(calc_set - atual_set))
        roles_nao_requeridas = sorted(list(atual_set - calc_set))
        ja_existentes = calc_set.intersection(atual_set)

        total_ja_existente += len(ja_existentes)
        total_novas_adicionar += len(roles_a_adicionar)
        total_duplicadas_evitadas += len(ja_existentes)
        total_roles_nao_requeridas += len(roles_nao_requeridas)

        # Gerar novas linhas propostas (somente se TEXT existir)
        for role_add in roles_a_adicionar:
            if text_comp:
                novas_linhas_pfcg.append({
                    "ID": proximo_id,
                    "AGR_NAME_COMPOSTA": comp,
                    "TEXT": text_comp,
                    "AGR_NAME": role_add
                })
                proximo_id += 1

        detalhes_fase3b.append({
            "composite": comp,
            "text": text_comp if text_comp else "TEXT_COMPOSITE_NAO_ENCONTRADO",
            "users_do_grupo": users_grupo,
            "funcoes_por_user": funcs_by_u,
            "composite_calculada": sorted(list(calc_set)),
            "composite_atual": sorted(list(atual_set)),
            "roles_a_adicionar": roles_a_adicionar,
            "roles_existentes_nao_requeridas": roles_nao_requeridas
        })

    wb.close()

    print(f"      - Concluída. Total Composites calculadas via União: {len(composta_roles_esperadas)}")
    print(f"      - Novas atribuições a adicionar na PFCG_COMPOSTA: {total_novas_adicionar}\n")

    res_fases = {
        "deps_processados": deps,
        "users_mapeados": users_mapeados_fase1,
        "users_pendentes": users_pendentes_classificacao,
        "erros_mapeamento": erros_mapeamento,
        "total_adicionar_fase2": total_adicionar_fase2,
        "total_tcodes_com_x": total_tcodes_com_x,
        "tcodes_sem_mapeamento": tcodes_sem_mapeamento_list,
        "total_casos_1n": total_casos_1n,
        "total_adicionar_fase3a": total_adicionar_fase3a,
        "total_compostas_fase3b": len(composta_roles_esperadas),
        "fase3b_detalhes": detalhes_fase3b,
        "fase3b_novas_linhas": novas_linhas_pfcg,
        "fase3b_resumo": {
            "total_composites_processadas": len(composta_roles_esperadas),
            "total_users_incluidos": len(set(u["uid"] for u in users_mapeados_fase1 if user_composite_excel.get(u["uid"]))),
            "total_relacoes_user_funcao": total_relacoes_user_funcao,
            "total_relacoes_comp_single_calc": total_relacoes_comp_single_calc,
            "total_ja_existente": total_ja_existente,
            "total_novas_adicionar": total_novas_adicionar,
            "total_duplicadas_evitadas": total_duplicadas_evitadas,
            "total_roles_nao_requeridas": total_roles_nao_requeridas,
            "total_composites_sem_text": total_composites_sem_text,
            "total_composites_sem_linha_previa": total_composites_sem_linha_previa
        }
    }

    # FASE 4: DISPARO AUTOMÁTICO DE SINCRONIZAÇÃO PFCG_COMPOSTA (DRY-RUN POR PADRÃO)
    print(f"\n[FASE 4] Iniciando verificação/sincronização automática de linhas pendentes em PFCG_COMPOSTA...")
    from sap_rfc.pfcg_composta_sync_service import sincronizar_pfcg_composta
    fase4_res = sincronizar_pfcg_composta(caminho_excel=caminho_excel, dry_run=simular, only_prd=only_prd)
    res_fases["fase4_resultado"] = fase4_res
    print(f"      - Concluída Fase 4. Status: {fase4_res.get('status')} | Total Pendências: {fase4_res.get('total_linhas_pendentes', 0)}\n")

    return res_fases


def main():
    parser = argparse.ArgumentParser(description="Executar Simulação Integrada de Reconciliação (Fases 1 a 4).")
    parser.add_argument("--excel", type=str, default=CAMINHO_EXCEL_PADRAO, help="Caminho do Excel mestre.")
    parser.add_argument("--real", action="store_true", help="Executar modo REAL no Excel (por padrão é simulação dry-run).")
    parser.add_argument("--only-prd", action="store_true", help="Executar Fase 4 exclusivamente contra SAP PRD (bypassa QAD).")
    args = parser.parse_args()

    simular = not args.real
    executar_reconciliacao_integrada(args.excel, simular=simular, only_prd=args.only_prd)


if __name__ == "__main__":
    main()

