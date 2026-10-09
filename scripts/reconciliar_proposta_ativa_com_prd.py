# -*- coding: utf-8 -*-
"""
scripts/reconciliar_proposta_ativa_com_prd.py
======================================================================
Script temporário para comparar as Funções Individuais (Single Roles)
já existentes na 'Proposta Ativa' do Excel mestre com as Single Roles
atribuídas diretamente aos utilizadores no SAP PRD.

REGRAS:
1. População de utilizadores extraída 100% dos departamentos PROCESSADOS do Excel.
2. FUNCOES_ATUAIS_EXCEL = Single Roles da linha do utilizador em 'Proposta Ativa'.
3. FUNCOES_DIRETAS_PRD = Single Roles diretas ativas no SAP PRD (AGR_USERS com COL_FLAG != 'X').
4. FUNCOES_A_ADICIONAR = FUNCOES_DIRETAS_PRD - FUNCOES_ATUAIS_EXCEL.
5. FUNCOES_EXCEL_NAO_CONFIRMADAS_PRD = FUNCOES_ATUAIS_EXCEL - FUNCOES_DIRETAS_PRD.
6. Estritamente READ-ONLY (sem alterações no Excel nem no SAP).
======================================================================
"""

import sys
import os
import re
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Set, Tuple, Any

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


def executar_reconciliacao_proposta_ativa(caminho_excel: str) -> Dict[str, Any]:
    print(f"\n==============================================================================")
    print(f"  RECONCILIAÇÃO DE FUNÇÕES INDIVIDUAIS: PROPOSTA ATIVA vs SAP PRD")
    print(f"  Ficheiro: {caminho_excel}")
    print(f"  Modo: SIMULAÇÃO (Somente Leitura)")
    print(f"==============================================================================\n")

    # 0. Obter conjunto de exclusão da sheet DEFINIÇÕES
    funcoes_excluidas_definicoes = obter_funcoes_excluidas_definicoes(caminho_excel)
    print(f"[0/5] Roles encontradas em DEFINIÇÕES para exclusão ({len(funcoes_excluidas_definicoes)}): {', '.join(sorted(list(funcoes_excluidas_definicoes)))}\n")

    # 1. Obter departamentos PROCESSADOS da sheet CONTROLO
    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    deps = obter_departamentos_processados(caminho_excel)
    print(f"[1/5] Departamentos PROCESSADOS na sheet CONTROLO ({len(deps)}): {', '.join(deps)}")

    # 2. Ler a população de User SAP EXCLUSIVAMENTE nas sheets departamentais correspondentes
    users_excel_info = []
    user_ids_set = set()

    print(f"\n[2/5] Mapeamento de Utilizadores (CONTROLO -> Sheet Departamental):")
    print("-" * 100)
    print(f"{'DEPARTAMENTO':<28} | {'USER SAP':<10} | {'NOME COMPLETO':<30} | {'COLUNA SHEET DEP':<18}")
    print("-" * 100)

    for d in deps:
        if d in wb.sheetnames:
            ws = wb[d]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) >= 2:
                header = rows[1]
                for col_idx in range(2, len(header)):
                    uid = extrair_user_id(header[col_idx])
                    if uid:
                        nome = str(header[col_idx] or "").replace(uid, "").strip().replace("\n", " ")
                        users_excel_info.append({"dep": d, "uid": uid, "nome": nome, "coluna": col_idx + 1})
                        user_ids_set.add(uid)
                        print(f"{d:<28} | {uid:<10} | {nome:<30} | Coluna {col_idx + 1:<11}")

    print("-" * 100)
    print(f"Total de utilizadores extraídos das sheets departamentais: {len(users_excel_info)} ({len(user_ids_set)} User IDs únicos)\n")

    # 3. Mapear a linha correspondente de cada User SAP na sheet 'Proposta Ativa'
    ws_prop = wb["Proposta Ativa"]
    rows_prop = list(ws_prop.iter_rows(values_only=True))

    proposta_ativa_map: Dict[str, int] = {}
    user_funcs_excel: Dict[str, Set[str]] = {}
    user_composite_excel: Dict[str, str] = {}

    for r_idx, r in enumerate(rows_prop[1:], start=2):
        if r[0]:
            uid_prop = str(r[0]).strip().upper()
            proposta_ativa_map[uid_prop] = r_idx

    print(f"[3/5] Mapeamento para a sheet Proposta Ativa:")
    print("-" * 100)
    print(f"{'DEPARTAMENTO':<28} | {'USER SAP':<10} | {'NOME COMPLETO':<30} | {'LINHA PROPOSTA ATIVA'}")
    print("-" * 100)

    erros_mapeamento = []
    for item in users_excel_info:
        uid = item["uid"]
        dep = item["dep"]
        nome = item["nome"]
        if uid in proposta_ativa_map:
            linha_prop = proposta_ativa_map[uid]
            item["linha_proposta_ativa"] = linha_prop
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
            print(f"{dep:<28} | {uid:<10} | {nome:<30} | Linha {linha_prop}")
        else:
            item["linha_proposta_ativa"] = None
            erros_mapeamento.append(item)
            print(f"{dep:<28} | {uid:<10} | {nome:<30} | ERRO_DE_MAPEAMENTO (Ausente na Proposta Ativa)")

    print("-" * 100)

    if erros_mapeamento:
        print(f"\n[!] ALERTA: {len(erros_mapeamento)} utilizadores da sheet departamental não foram encontrados na Proposta Ativa.")
        for err in erros_mapeamento:
            print(f"    - {err['dep']} | {err['uid']} ({err['nome']}) -> ERRO_DE_MAPEAMENTO")

    wb.close()

    # 4. Consultar SAP PRD (Read-Only) para os User SAP válidos
    print(f"\n[4/5] A consultar Single Roles diretas no SAP PRD para os {len(user_ids_set)} User SAP...")
    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(("AGR_USERS", "AGR_AGRS", "AGR_1251", "AGR_TCODES", "AGR_DEFINE", "USR02"))

    conn = Connection(**params)
    today_str = datetime.now().strftime("%Y%m%d")
    user_ids_list = sorted(list(user_ids_set))

    user_singles_prd: Dict[str, Set[str]] = {uid: set() for uid in user_ids_list}

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
    print("      Consulta SAP PRD concluída com sucesso.\n")

    # 5. Comparação por Utilizador
    print(f"[5/5] A comparar Proposta Ativa vs SAP PRD com filtro da sheet DEFINIÇÕES...")
    print("=" * 165)
    print(f"{'DEPARTAMENTO':<25} | {'USER':<9} | {'LINHA PA':<8} | {'PRD DIRETA':<10} | {'EXCLUÍDAS DEF':<13} | {'ELEGÍVEIS':<9} | {'EXCEL':<5} | {'A ADICIONAR':<12} | {'NÃO CONFIRMADAS'}")
    print("-" * 165)

    from collections import Counter
    c_roles_add = Counter()
    relatorio_utilizadores = []

    total_singles_prd_diretas = 0
    total_excluidas_definicoes = 0
    total_elegiveis_após_definicoes = 0
    total_singles_excel = 0
    total_adicionar = 0
    total_nao_confirmadas = 0
    users_com_diferencas = 0
    users_sem_diferencas = 0

    for item in users_excel_info:
        dep = item["dep"]
        uid = item["uid"]
        nome = item["nome"]
        linha_pa = item.get("linha_proposta_ativa")

        if not linha_pa:
            print(f"{dep:<25} | {uid:<9} | {'N/A':<8} | {'-':<10} | {'-':<13} | {'-':<9} | {'-':<5} | {'-':<12} | ERRO_DE_MAPEAMENTO")
            continue

        comp_ex = user_composite_excel.get(uid, "")
        funcs_ex = user_funcs_excel.get(uid, set())
        funcs_prd_diretas = user_singles_prd.get(uid, set())

        # Aplicação da regra de exclusão da sheet DEFINIÇÕES
        funcs_excluidas = funcs_prd_diretas.intersection(funcoes_excluidas_definicoes)
        funcs_elegiveis = funcs_prd_diretas - funcoes_excluidas_definicoes

        a_adicionar = sorted(list(funcs_elegiveis - funcs_ex))
        nao_confirmadas = sorted(list(funcs_ex - funcs_prd_diretas))

        total_singles_prd_diretas += len(funcs_prd_diretas)
        total_excluidas_definicoes += len(funcs_excluidas)
        total_elegiveis_após_definicoes += len(funcs_elegiveis)
        total_singles_excel += len(funcs_ex)
        total_adicionar += len(a_adicionar)
        total_nao_confirmadas += len(nao_confirmadas)

        if len(a_adicionar) > 0 or len(nao_confirmadas) > 0:
            users_com_diferencas += 1
        else:
            users_sem_diferencas += 1

        for r_add in a_adicionar:
            c_roles_add[r_add] += 1

        relatorio_utilizadores.append({
            "departamento": dep,
            "user_id": uid,
            "nome": nome,
            "linha_proposta_ativa": linha_pa,
            "composite_role_excel": comp_ex,
            "total_prd_diretas": len(funcs_prd_diretas),
            "total_excluidas_def": len(funcs_excluidas),
            "total_elegiveis": len(funcs_elegiveis),
            "total_excel": len(funcs_ex),
            "a_adicionar": a_adicionar,
            "excluidas_por_definicoes": sorted(list(funcs_excluidas)),
            "nao_confirmadas": nao_confirmadas,
        })

        print(f"{dep:<25} | {uid:<9} | Linha {linha_pa:<2} | {len(funcs_prd_diretas):<10} | {len(funcs_excluidas):<13} | {len(funcs_elegiveis):<9} | {len(funcs_ex):<5} | {len(a_adicionar):<12} | {len(nao_confirmadas)}")
        if funcs_excluidas:
            print(f"   [*] EXCLUÍDAS POR DEFINIÇÕES ({len(funcs_excluidas)}): {', '.join(sorted(list(funcs_excluidas)))}")
        if a_adicionar:
            print(f"   [+] FUNÇÕES A ADICIONAR ({len(a_adicionar)}): {', '.join(a_adicionar)}")
        if nao_confirmadas:
            print(f"   [-] NÃO CONFIRMADAS EM PRD ({len(nao_confirmadas)}): {', '.join(nao_confirmadas)}")

    print("=" * 165 + "\n")

    # Resumo Geral e Validação Matemática
    print("==============================================================================")
    print("  RESUMO GERAL COM FILTRO DEFINIÇÕES (PROPOSTA ATIVA vs SAP PRD)")
    print("==============================================================================")
    print(f"  Departamentos PROCESSADOS:                         {len(deps)} ({', '.join(deps)})")
    print(f"  Total de Utilizadores Mapeados Analisados:          {len(relatorio_utilizadores)}")
    print(f"  Total de Single Roles Diretas em PRD:               {total_singles_prd_diretas}")
    print(f"  Total Excluídas pela sheet DEFINIÇÕES:             {total_excluidas_definicoes}")
    print(f"  Total Elegíveis após Filtro DEFINIÇÕES:            {total_elegiveis_após_definicoes}")
    print(f"  Total de Funções Individuais Atuais na Proposta:   {total_singles_excel}")
    print(f"  Total de Funções Individuais a ADICIONAR:           {total_adicionar}")
    print(f"  Total de Funções Excel NÃO CONFIRMADAS em PRD:      {total_nao_confirmadas}")
    print("------------------------------------------------------------------------------")
    print(f"  VALIDAÇÃO MATEMÁTICA: {total_singles_prd_diretas} (PRD) == {total_excluidas_definicoes} (Excluídas) + {total_elegiveis_após_definicoes} (Elegíveis)")
    if total_singles_prd_diretas == (total_excluidas_definicoes + total_elegiveis_após_definicoes):
        print("  => [OK] IGUALDADE CONIRMADA MATEMATICAMENTE!")
    else:
        print("  => [!] ERRO: FALHA NA EQUAÇÃO MATEMÁTICA DOS TOTAIS!")
    print("------------------------------------------------------------------------------")
    print("  TOP ROLES ELEGÍVEIS MAIS ADICIONADAS:")
    for r_name, count in c_roles_add.most_common(15):
        print(f"    - {r_name:<40}: {count} utilizadores")
    print("==============================================================================\n")

    return {
        "relatorio_utilizadores": relatorio_utilizadores,
        "total_singles_prd_diretas": total_singles_prd_diretas,
        "total_excluidas_definicoes": total_excluidas_definicoes,
        "total_elegiveis_após_definicoes": total_elegiveis_após_definicoes,
        "total_adicionar": total_adicionar,
        "total_nao_confirmadas": total_nao_confirmadas,
        "roles_mais_adicionadas": c_roles_add.most_common(15)
    }


def main():
    parser = argparse.ArgumentParser(description="Reconciliar Proposta Ativa com Single Roles do SAP PRD.")
    parser.add_argument("--excel", type=str, default=CAMINHO_EXCEL_PADRAO, help="Caminho do Excel mestre.")
    parser.add_argument("--simular", action="store_true", default=True, help="Modo simulação.")
    args = parser.parse_args()

    executar_reconciliacao_proposta_ativa(args.excel)


if __name__ == "__main__":
    main()

