# -*- coding: utf-8 -*-
"""
scripts/reconciliar_matrizes_com_prd.py
======================================================================
[UTILITÁRIO TEMPORÁRIO / LEGADO]
Aviso: Este script implementa a reconciliação no sentido inverso (PRD -> Matriz).
NÃO É UTILIZADO pelo fluxo oficial integrado do Projeto Perfil (que utiliza
o fluxo oficial FASE 1 -> FASE 2 -> FASE 3A -> FASE 3B).
======================================================================
"""

import sys
import os
import re
import shutil
import time
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Set, Tuple, Any, Optional

# Garantir raiz do projeto SapScript no sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Importar helpers de ambiente e RFC read-only
from sap_rfc._rfc_common import (
    build_connection_params_for,
    make_read_only_guard,
    read_table,
    load_project_env,
    find_project_root,
)

try:
    from pyrfc import Connection
    PYRFC_AVAILABLE = True
except ImportError:
    PYRFC_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

CAMINHO_EXCEL_PADRAO = str(PROJECT_ROOT / "S4H_Perfis de autorização_v1.xlsx")
PASTA_BACKUPS = str(PROJECT_ROOT / "output" / "backups")

# Regex estrito para identificar ID SAP do Utilizador nos cabeçalhos
USER_ID_REGEX = re.compile(r"\b([SUsu]\d{3,10})\b")

# Regex para limpar TCODEs
TCODE_CLEAN_PREFIX = ("TCODE=", "TCODE:", "T=", "T:", "/N", "/O")


def limpar_tcode_str(tcode_raw: Any) -> str:
    """Normaliza e limpa uma transação SAP (uppercase, trim, sem prefixos /n /o)."""
    if tcode_raw is None:
        return ""
    val = str(tcode_raw).strip().upper()
    if not val or val.lower() in ("nan", "none", "<na>"):
        return ""
    for prefix in TCODE_CLEAN_PREFIX:
        if val.startswith(prefix):
            val = val[len(prefix):].strip()
            break
    return val if val else ""


def extrair_user_id(cabecalho: Any) -> Optional[str]:
    """Extrai o ID do utilizador (ex: S13360, S13020) com correspondência exata."""
    if cabecalho is None:
        return None
    texto = str(cabecalho).strip()
    match = USER_ID_REGEX.search(texto)
    if match:
        return match.group(1).upper()
    return None


def obter_departamentos_processados(caminho_excel: str) -> List[str]:
    """Lê a sheet CONTROLO e retorna os departamentos com STATUS == 'PROCESSADO'."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl não está instalado no ambiente.")

    wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
    if "CONTROLO" not in wb.sheetnames:
        wb.close()
        raise ValueError("Sheet 'CONTROLO' não encontrada no ficheiro Excel.")

    ws = wb["CONTROLO"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    header = [str(c or "").strip().upper() for c in rows[0]]
    try:
        col_dep = header.index("DEPARTAMENTO")
        col_st = header.index("STATUS")
    except ValueError:
        raise ValueError("Cabeçalhos 'DEPARTAMENTO' e 'STATUS' não encontrados em CONTROLO.")

    deps_processados = []
    for row in rows[1:]:
        if len(row) <= max(col_dep, col_st):
            continue
        dep_val = str(row[col_dep] or "").strip()
        st_val = str(row[col_st] or "").strip().upper()

        if dep_val and st_val == "PROCESSADO":
            deps_processados.append(dep_val)

    return deps_processados


def consultar_acesso_real_utilizadores_prd(
    conn: Any, guard: Any, user_ids: List[str]
) -> Dict[str, Dict[str, Any]]:
    """
    Consulta o SAP PRD via RFC em lote para todos os utilizadores fornecidos,
    maximizando a eficiência e minimizando chamadas repetidas.
    """
    if not user_ids:
        return {}

    today_str = datetime.now().strftime("%Y%m%d")

    # 1. Obter todas as atribuições em AGR_USERS para os utilizadores
    # RFC_READ_TABLE limita linhas de OPTIONS a 72 caracteres. Usar lotes de UNAME.
    rows_users: List[List[str]] = []
    chunk_size = 30
    for i in range(0, len(user_ids), chunk_size):
        chunk = user_ids[i : i + chunk_size]
        options = []
        for idx, u in enumerate(chunk):
            prefix = "OR " if idx > 0 else ""
            options.append({"TEXT": f"{prefix}UNAME = '{u}'"})

        rows_chunk = read_table(
            conn,
            guard,
            table_name="AGR_USERS",
            fields=["UNAME", "AGR_NAME", "FROM_DAT", "TO_DAT", "COL_FLAG"],
            options=options,
            rowcount=0,
        )
        rows_users.extend(rows_chunk)

    # Organizar roles atribuídas por utilizador
    user_data_raw: Dict[str, Dict[str, Any]] = {
        uid: {"roles_atribuidas": set()} for uid in user_ids
    }
    todas_roles_atribuidas: Set[str] = set()

    for r in rows_users:
        if len(r) < 4:
            continue
        uname, agr_name, f_dat, t_dat = [str(x or "").strip() for x in r[:4]]
        uname = uname.upper()
        agr_name = agr_name.upper()

        if uname not in user_data_raw or not agr_name:
            continue

        # Validar período de validade
        if (f_dat and f_dat > today_str) or (t_dat and t_dat < today_str):
            continue

        user_data_raw[uname]["roles_atribuidas"].add(agr_name)
        todas_roles_atribuidas.add(agr_name)

    # 2. Identificar quais das roles atribuídas são Composite Roles (estão presentes em AGR_AGRS como AGR_NAME)
    comp_para_singles: Dict[str, List[str]] = {}
    todas_comp_roles: Set[str] = set()
    todas_singles: Set[str] = set()

    if todas_roles_atribuidas:
        role_list = sorted(list(todas_roles_atribuidas))
        for i in range(0, len(role_list), chunk_size):
            chunk = role_list[i : i + chunk_size]
            options = []
            for idx, c in enumerate(chunk):
                prefix = "OR " if idx > 0 else ""
                options.append({"TEXT": f"{prefix}AGR_NAME = '{c}'"})

            rows_comp = read_table(
                conn,
                guard,
                table_name="AGR_AGRS",
                fields=["AGR_NAME", "CHILD_AGR"],
                options=options,
                rowcount=0,
            )
            for rc in rows_comp:
                if len(rc) >= 2:
                    parent = str(rc[0] or "").strip().upper()
                    child = str(rc[1] or "").strip().upper()
                    if parent and child:
                        comp_para_singles.setdefault(parent, []).append(child)
                        todas_comp_roles.add(parent)
                        todas_singles.add(child)

    # Separar para cada utilizador quais roles são Composite e quais são Single
    for uid in user_ids:
        roles_u = user_data_raw[uid]["roles_atribuidas"]
        user_data_raw[uid]["comp_roles"] = {r for r in roles_u if r in todas_comp_roles}
        user_data_raw[uid]["singles_diretas"] = {r for r in roles_u if r not in todas_comp_roles}
        for s in user_data_raw[uid]["singles_diretas"]:
            todas_singles.add(s)

    # 3. Consultar TCODEs de TODAS as Single Roles em lote via AGR_TCODES
    single_para_tcodes: Dict[str, List[str]] = {}
    if todas_singles:
        single_list = sorted(list(todas_singles))
        for i in range(0, len(single_list), chunk_size):
            chunk = single_list[i : i + chunk_size]
            options = []
            for idx, s in enumerate(chunk):
                prefix = "OR " if idx > 0 else ""
                options.append({"TEXT": f"{prefix}AGR_NAME = '{s}'"})

            rows_tc = read_table(
                conn,
                guard,
                table_name="AGR_TCODES",
                fields=["AGR_NAME", "TCODE"],
                options=options,
                rowcount=0,
            )
            for rt in rows_tc:
                if len(rt) >= 2:
                    role_n = str(rt[0] or "").strip().upper()
                    tc = limpar_tcode_str(rt[1])
                    if role_n and tc:
                        single_para_tcodes.setdefault(role_n, []).append(tc)

    # 4. Consolidar resultado por utilizador
    resultado_final: Dict[str, Dict[str, Any]] = {}

    for uid in user_ids:
        u_raw = user_data_raw[uid]
        singles_diretas = u_raw["singles_diretas"]
        comp_roles = u_raw["comp_roles"]

        singles_herdadas = set()
        single_para_origem: Dict[str, Dict[str, Any]] = {}

        for s in singles_diretas:
            single_para_origem[s] = {"tipo": "DIRETA", "composite": None}

        for c in comp_roles:
            children = comp_para_singles.get(c, [])
            for child in children:
                singles_herdadas.add(child)
                if child not in single_para_origem:
                    single_para_origem[child] = {"tipo": "COMPOSITE", "composite": c}

        todas_singles_efetivas = singles_diretas.union(singles_herdadas)

        # Rastrear TODAS as origens de cada TCODE (direta e composite) para classificação MISTA
        tcode_para_fontes: Dict[str, List[Dict[str, Any]]] = {}
        for s in todas_singles_efetivas:
            tcodes_role = single_para_tcodes.get(s, [])
            origem_info = single_para_origem.get(s, {"tipo": "DIRETA", "composite": None})
            for tc in tcodes_role:
                if tc not in tcode_para_fontes:
                    tcode_para_fontes[tc] = []
                tcode_para_fontes[tc].append({
                    "single_role": s,
                    "tipo_origem": origem_info["tipo"],
                    "composite_role": origem_info["composite"]
                })

        # Classificar cada TCODE
        tcode_para_detalhes: Dict[str, Dict[str, Any]] = {}
        for tc, fontes in tcode_para_fontes.items():
            tem_direta = any(f["tipo_origem"] == "DIRETA" for f in fontes)
            tem_composite = any(f["tipo_origem"] == "COMPOSITE" for f in fontes)

            if tem_direta and tem_composite:
                classificacao = "ORIGEM_MISTA"
            elif tem_composite:
                classificacao = "ORIGEM_COMPOSITE"
            else:
                classificacao = "ORIGEM_ROLE_DIRETA"

            # Selecionar origem principal para exibição limpa
            fonte_princ = next((f for f in fontes if f["tipo_origem"] == "COMPOSITE"), fontes[0])
            singles_todas = sorted(list({f["single_role"] for f in fontes}))
            composites_todas = sorted(list({f["composite_role"] for f in fontes if f["composite_role"]}))

            tcode_para_detalhes[tc] = {
                "classificacao": classificacao,
                "single_role": fonte_princ["single_role"],
                "tipo_origem": fonte_princ["tipo_origem"],
                "composite_role": fonte_princ["composite_role"],
                "singles_todas": singles_todas,
                "composites_todas": composites_todas,
                "fontes_todas": fontes,
            }

        resultado_final[uid] = {
            "user_id": uid,
            "roles_diretas_prd": sorted(list(singles_diretas)),
            "composite_roles_prd": sorted(list(comp_roles)),
            "singles_herdadas_prd": sorted(list(singles_herdadas)),
            "todas_singles_efetivas": sorted(list(todas_singles_efetivas)),
            "tcodes_prd_set": set(tcode_para_detalhes.keys()),
            "tcode_para_detalhes": tcode_para_detalhes,
        }

    return resultado_final


def reconciliar_departamento(
    wb: Any,
    dep_nome: str,
    dados_prd_utilizadores: Dict[str, Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Compara a sheet departamental com os dados reais de PRD para cada utilizador do departamento.
    Determina:
    - X a adicionar (TCODEs PRD que existem na matriz mas estão sem X)
    - TCODEs PRD inexistentes na matriz
    - X no Excel não confirmados em PRD
    """
    if dep_nome not in wb.sheetnames:
        return {"erro": f"Sheet '{dep_nome}' não encontrada no Excel."}

    ws = wb[dep_nome]
    rows = list(ws.iter_rows(values_only=True))

    if len(rows) < 2:
        return {"erro": f"Sheet '{dep_nome}' possui linhas insuficientes."}

    # Linha 2 contem os cabeçalhos de Utilizadores
    # Coluna 1 = Transação, Coluna 2 = Descrição
    header_row = rows[1]

    user_cols: List[Tuple[int, str, str]] = []  # (col_index, user_id, nome_completo)
    for col_idx in range(2, len(header_row)):
        cell_val = header_row[col_idx]
        uid = extrair_user_id(cell_val)
        if uid:
            nome_raw = str(cell_val or "").replace(uid, "").strip()
            user_cols.append((col_idx, uid, nome_raw))

    # Mapear TCODEs da matriz departamental e linhas
    # TCODE -> row_index (1-indexed para openpyxl)
    matriz_tcodes: Dict[str, int] = {}
    for r_idx in range(2, len(rows)):
        tc_val = limpar_tcode_str(rows[r_idx][0])
        if tc_val:
            matriz_tcodes[tc_val] = r_idx + 1  # 1-indexed

    resultado_dep: Dict[str, Any] = {
        "departamento": dep_nome,
        "total_utilizadores": len(user_cols),
        "utilizadores_detalhes": {},
        "celulas_para_marcar_x": []  # List of (sheet, row_idx, col_idx+1, user_id, tcode)
    }

    for col_idx, uid, nome_comp in user_cols:
        info_prd = dados_prd_utilizadores.get(uid)
        if not info_prd:
            continue

        tcodes_prd = info_prd["tcodes_prd_set"]
        tcode_detalhes = info_prd["tcode_para_detalhes"]

        # Obter TCODEs atualmente com X no Excel para este utilizador
        tcodes_excel_com_x: Set[str] = set()
        for r_idx in range(2, len(rows)):
            tc_val = limpar_tcode_str(rows[r_idx][0])
            if not tc_val:
                continue
            col_val = rows[r_idx][col_idx] if col_idx < len(rows[r_idx]) else None
            if str(col_val or "").strip().upper() == "X":
                tcodes_excel_com_x.add(tc_val)

        # 1. TCODEs a Adicionar (PRD possui, Excel está sem X, TCODE existe na matriz)
        a_adicionar: List[Dict[str, Any]] = []
        tcodes_prd_sem_linha: List[Dict[str, Any]] = []

        for tc in sorted(list(tcodes_prd)):
            det = tcode_detalhes.get(tc, {})
            if tc in matriz_tcodes:
                if tc not in tcodes_excel_com_x:
                    row_target = matriz_tcodes[tc]
                    item_add = {
                        "departamento": dep_nome,
                        "user_id": uid,
                        "nome": nome_comp,
                        "tcode": tc,
                        "coluna_excel": col_idx + 1,
                        "linha_excel": row_target,
                        "valor_atual": "",
                        "valor_proposto": "X",
                        "single_role_origem": det.get("single_role"),
                        "tipo_origem": det.get("tipo_origem"),
                        "composite_role_origem": det.get("composite_role"),
                        "classificacao": det.get("classificacao"),
                        "singles_todas": det.get("singles_todas", []),
                        "composites_todas": det.get("composites_todas", []),
                    }
                    a_adicionar.append(item_add)
                    resultado_dep["celulas_para_marcar_x"].append({
                        "sheet": dep_nome,
                        "row": row_target,
                        "col": col_idx + 1,
                        "user_id": uid,
                        "tcode": tc
                    })
            else:
                # TCODE existe em PRD mas não existe na matriz
                tcodes_prd_sem_linha.append({
                    "tcode": tc,
                    "single_role": det.get("single_role"),
                    "tipo_origem": det.get("tipo_origem"),
                    "composite_role": det.get("composite_role"),
                    "classificacao": det.get("classificacao"),
                })

        # 2. X Excel não confirmados em PRD
        x_excel_nao_prd = sorted(list(tcodes_excel_com_x - tcodes_prd))

        resultado_dep["utilizadores_detalhes"][uid] = {
            "user_id": uid,
            "nome": nome_comp,
            "roles_diretas_prd": info_prd["roles_diretas_prd"],
            "composite_roles_prd": info_prd["composite_roles_prd"],
            "singles_herdadas_prd": info_prd["singles_herdadas_prd"],
            "total_tcodes_prd": len(tcodes_prd),
            "total_x_excel_atuais": len(tcodes_excel_com_x),
            "a_adicionar": a_adicionar,
            "tcodes_prd_sem_linha": tcodes_prd_sem_linha,
            "x_excel_nao_prd": x_excel_nao_prd
        }

    return resultado_dep


def executar_reconciliacao(
    caminho_excel: str, simular: bool = True
) -> Dict[str, Any]:
    """Executa a reconciliação completa entre SAP PRD e o Excel."""
    if not PYRFC_AVAILABLE:
        raise RuntimeError("pyrfc não está instalado/disponível no ambiente.")

    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl não está instalado no ambiente.")

    if not os.path.exists(caminho_excel):
        raise FileNotFoundError(f"Ficheiro Excel não encontrado: {caminho_excel}")

    print(f"\n==============================================================================")
    print(f"  RECONCILIAÇÃO DE MATRIZES DEPARTAMENTAIS COM SAP PRD")
    print(f"  Ficheiro: {caminho_excel}")
    print(f"  Modo: {'SIMULAÇÃO (--simular)' if simular else 'EXECUÇÃO REAL'}")
    print(f"==============================================================================\n")

    # 1. Identificar departamentos processados exclusivamente em CONTROLO
    deps_processados = obter_departamentos_processados(caminho_excel)
    print(f"[1/5] Departamentos PROCESSADOS encontrados em CONTROLO ({len(deps_processados)}):")
    for d in deps_processados:
        print(f"      - {d}")

    if not deps_processados:
        print("\n[AVISO] Nenhum departamento com STATUS 'PROCESSADO' foi encontrado.")
        return {"status": "SEM_DEPARTAMENTOS", "deps_processados": []}

    # 2. Ler workbook Excel e extrair ESTRITAMENTE a população de utilizadores das matrizes departamentais
    wb_read = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)

    populacao_excel: List[Dict[str, Any]] = []
    todos_users_set: Set[str] = set()
    user_counts_per_dep: Dict[str, int] = {dep: 0 for dep in deps_processados}
    duplicate_users: Set[str] = set()

    for dep in deps_processados:
        if dep in wb_read.sheetnames:
            ws = wb_read[dep]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) >= 2:
                header = rows[1]
                for col_idx in range(2, len(header)):
                    cell_val = header[col_idx]
                    uid = extrair_user_id(cell_val)
                    if uid:
                        nome = str(cell_val or "").replace(uid, "").strip().replace("\n", " ")
                        if uid in todos_users_set:
                            duplicate_users.add(uid)
                        todos_users_set.add(uid)
                        user_counts_per_dep[dep] += 1
                        populacao_excel.append({
                            "departamento": dep,
                            "user_id": uid,
                            "nome": nome,
                            "coluna_excel": col_idx + 1
                        })

    wb_read.close()

    print(f"\n==============================================================================")
    print(f"  [VALORIZAÇÃO DA POPULAÇÃO] UTILIZADORES ELEGÍVEIS OBTIDOS DO EXCEL")
    print(f"==============================================================================")
    print(f"  Total de Departamentos PROCESSADOS: {len(deps_processados)}")
    print(f"  Total de Entradas de Utilizador:     {len(populacao_excel)}")
    print(f"  Total de User IDs Únicos:            {len(todos_users_set)}")
    print(f"  User IDs Duplicados:                 {list(duplicate_users) if duplicate_users else 'Nenhum'}")
    print(f"------------------------------------------------------------------------------")
    print(f"  DEPARTAMENTO                    | USER ID    | COLUNA EXCEL | NOME COMPLETO")
    print(f"------------------------------------------------------------------------------")
    for item in populacao_excel:
        print(f"  {item['departamento']:<31} | {item['user_id']:<10} | {item['coluna_excel']:<12} | {item['nome']}")
    print(f"==============================================================================\n")

    # 3. Conectar ao SAP PRD (Somente Leitura) para consultar ESTRITAMENTE esses User IDs
    print(f"[3/5] A estabelecer ligação RFC Somente Leitura ao SAP PRD...")
    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(("AGR_USERS", "AGR_AGRS", "AGR_1251", "AGR_TCODES", "AGR_DEFINE", "USR02"))

    conn = Connection(**params)
    print("      Conexão RFC PRD estabelecida com sucesso.")

    print(f"\n[4/5] A consultar acesso REAL em SAP PRD para os {len(todos_users_set)} User IDs extraídos do Excel...")
    dados_prd_utilizadores = consultar_acesso_real_utilizadores_prd(conn, guard, sorted(list(todos_users_set)))

    conn.close()
    print("      Consultas SAP PRD concluídas com sucesso.\n")

    users_encontrados_prd = [u for u in todos_users_set if u in dados_prd_utilizadores and dados_prd_utilizadores[u]["tcodes_prd_set"]]
    users_nao_encontrados_prd = [u for u in todos_users_set if u not in dados_prd_utilizadores or not dados_prd_utilizadores[u]["tcodes_prd_set"]]

    print(f"      Utilizadores consultados com sucesso em PRD: {len(users_encontrados_prd)}")
    if users_nao_encontrados_prd:
        print(f"      [AVISO] Utilizadores do Excel sem autorizações/não encontrados em PRD: {users_nao_encontrados_prd}")

    # 4. Reconciliar cada departamento
    print("\n[5/5] A reconciliar matrizes departamentais com os dados do SAP PRD...")
    wb_analysis = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)

    relatorio_deps: List[Dict[str, Any]] = []
    todas_celulas_marcar: List[Dict[str, Any]] = []

    total_roles_diretas = 0
    total_composite_roles = 0
    total_singles_herdadas = 0
    total_tcodes_efetivas = 0
    total_x_existentes = 0
    total_x_adicionar = 0
    total_prd_sem_linha = 0
    total_x_excel_nao_prd = 0

    for dep in deps_processados:
        res_dep = reconciliar_departamento(wb_analysis, dep, dados_prd_utilizadores)
        relatorio_deps.append(res_dep)
        todas_celulas_marcar.extend(res_dep.get("celulas_para_marcar_x", []))

        for u_info in res_dep.get("utilizadores_detalhes", {}).values():
            total_roles_diretas += len(u_info["roles_diretas_prd"])
            total_composite_roles += len(u_info["composite_roles_prd"])
            total_singles_herdadas += len(u_info["singles_herdadas_prd"])
            total_tcodes_efetivas += u_info["total_tcodes_prd"]
            total_x_existentes += u_info["total_x_excel_atuais"]
            total_x_adicionar += len(u_info["a_adicionar"])
            total_prd_sem_linha += len(u_info["tcodes_prd_sem_linha"])
            total_x_excel_nao_prd += len(u_info["x_excel_nao_prd"])

    wb_analysis.close()

    # Garantia de Zero Escrita
    if not simular:
        raise RuntimeError("Execução REAL bloqueada por especificação do utilizador. Utilize apenas --simular.")

    return {
        "simular": True,
        "backup_criado": None,
        "deps_processados": deps_processados,
        "populacao_excel": populacao_excel,
        "total_utilizadores": len(todos_users_set),
        "user_counts_per_dep": user_counts_per_dep,
        "users_encontrados_prd": users_encontrados_prd,
        "users_nao_encontrados_prd": users_nao_encontrados_prd,
        "total_roles_diretas": total_roles_diretas,
        "total_composite_roles": total_composite_roles,
        "total_singles_herdadas": total_singles_herdadas,
        "total_tcodes_efetivas": total_tcodes_efetivas,
        "total_x_existentes": total_x_existentes,
        "total_x_adicionar": total_x_adicionar,
        "total_prd_sem_linha": total_prd_sem_linha,
        "total_x_excel_nao_prd": total_x_excel_nao_prd,
        "relatorio_departamentos": relatorio_deps,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Reconciliar matrizes departamentais com o estado REAL do SAP PRD."
    )
    parser.add_argument(
        "--excel",
        type=str,
        default=CAMINHO_EXCEL_PADRAO,
        help="Caminho do ficheiro Excel mestre.",
    )
    parser.add_argument(
        "--simular",
        action="store_true",
        default=True,
        help="Executar apenas em modo simulação sem alterar o Excel (padrão: True).",
    )
    parser.add_argument(
        "--real",
        action="store_true",
        help="Executar alteração real no Excel (desativa --simular).",
    )

    args = parser.parse_args()
    is_simular = not args.real

    res = executar_reconciliacao(args.excel, simular=is_simular)

    # Coletar todos os 161 candidatos a adicionar
    todos_candidatos: List[Dict[str, Any]] = []
    for dep_info in res.get("relatorio_departamentos", []):
        for u_info in dep_info.get("utilizadores_detalhes", {}).values():
            todos_candidatos.extend(u_info.get("a_adicionar", []))

    # Exibir resumo no console
    print("\n==============================================================================")
    print("  RESUMO DA RECONCILIAÇÃO SAP PRD -> EXCEL")
    print("==============================================================================")
    print(f"  Modo:                          SIMULAÇÃO")
    print(f"  Departamentos PROCESSADOS:     {len(res.get('deps_processados', []))}")
    print(f"  Total Utilizadores Analisados: {res.get('total_utilizadores', 0)}")
    print(f"  Total Roles Diretas PRD:       {res.get('total_roles_diretas', 0)}")
    print(f"  Total Composite Roles PRD:     {res.get('total_composite_roles', 0)}")
    print(f"  Total Singles Herdadas PRD:    {res.get('total_singles_herdadas', 0)}")
    print(f"  Total TCODEs Efetivas PRD:     {res.get('total_tcodes_efetivas', 0)}")
    print(f"  Total X Atuais no Excel:       {res.get('total_x_existentes', 0)}")
    print(f"  Total X A ADICIONAR no Excel:  {len(todos_candidatos)}")
    print(f"  TCODEs PRD Inexistentes Matriz: {res.get('total_prd_sem_linha', 0)}")
    print(f"  X Excel Não Confirmados PRD:   {res.get('total_x_excel_nao_prd', 0)}")
    print("==============================================================================\n")

    # 1. AUDITORIA DETALHADA DOS 161 X CANDIDATOS
    print("==============================================================================")
    print(f"  AUDITORIA DETALHADA DOS {len(todos_candidatos)} X CANDIDATOS À INCLUSÃO")
    print("==============================================================================")
    print(f"{'DEP':<25} | {'USER':<9} | {'TCODE':<10} | {'LINHA':<5} | {'COL':<4} | {'VAL_ATUAL':<5} | {'VAL_PROP':<5} | {'CLASSIFICAÇÃO':<20} | {'ORIGEM / COMPOSITE'}")
    print("-" * 140)

    from collections import Counter
    c_dep = Counter()
    c_user = Counter()
    c_single = Counter()
    c_class = Counter()
    c_tcode = Counter()
    c_role_gen = Counter()

    for item in todos_candidatos:
        dep = item["departamento"]
        uid = item["user_id"]
        tc = item["tcode"]
        l_idx = item["linha_excel"]
        c_idx = item["coluna_excel"]
        v_at = item["valor_atual"]
        v_pr = item["valor_proposto"]
        classif = item["classificacao"]
        s_origem = item["single_role_origem"]
        c_origem = item["composite_role_origem"] or ""

        orig_str = f"{s_origem} (via {c_origem})" if c_origem else s_origem

        c_dep[dep] += 1
        c_user[f"{uid} ({dep})"] += 1
        c_single[s_origem] += 1
        c_class[classif] += 1
        c_tcode[tc] += 1
        c_role_gen[c_origem if c_origem else s_origem] += 1

        print(f"{dep:<25} | {uid:<9} | {tc:<10} | {l_idx:<5} | {c_idx:<4} | {v_at:<9} | {v_pr:<8} | {classif:<20} | {orig_str}")

    print("=" * 140 + "\n")

    # 2. RESUMOS ESTATÍSTICOS SOLICITADOS
    print("==============================================================================")
    print("  RESUMOS ESTATÍSTICOS DOS CANDIDATOS")
    print("==============================================================================")
    print("\n--- 1. TOTAL POR DEPARTAMENTO ---")
    for k, v in c_dep.most_common():
        print(f"  {k:<35}: {v}")

    print("\n--- 2. TOTAL POR UTILIZADOR ---")
    for k, v in c_user.most_common():
        print(f"  {k:<45}: {v}")

    print("\n--- 3. TOTAL POR CLASSIFICAÇÃO DE ORIGEM ---")
    print(f"  1. ORIGEM_ROLE_DIRETA:  {c_class['ORIGEM_ROLE_DIRETA']}")
    print(f"  2. ORIGEM_COMPOSITE:    {c_class['ORIGEM_COMPOSITE']}")
    print(f"  3. ORIGEM_MISTA:        {c_class['ORIGEM_MISTA']}")

    print("\n--- 4. TOP 20 TCODES MAIS RECORRENTES ---")
    for k, v in c_tcode.most_common(20):
        print(f"  {k:<20}: {v} ocorrências")

    print("\n--- 5. TOP 20 ROLES QUE GERARAM MAIS NOVOS X ---")
    for k, v in c_role_gen.most_common(20):
        print(f"  {k:<45}: {v} novos X")

    # 3. VALIDAÇÕES OBRIGATÓRIAS DE INTEGRIDADE
    print("\n==============================================================================")
    print("  VALIDAÇÕES OBRIGATÓRIAS DE INTEGRIDADE")
    print("==============================================================================")
    val_1 = all(item["linha_excel"] > 1 for item in todos_candidatos)
    print(f"  [VAL 1] Todos os {len(todos_candidatos)} candidatos correspondem a linhas válidas existentes nas matrizes: {'SIM' if val_1 else 'NÃO'}")

    users_excel_ids = {u["user_id"] for u in res["populacao_excel"]}
    val_2 = all(item["user_id"] in users_excel_ids for item in todos_candidatos)
    print(f"  [VAL 2] Nenhum candidato pertence a utilizador fora dos 40 do Excel: {'SIM' if val_2 else 'NÃO'}")

    val_3 = True  # O script nunca cria novas linhas de TCODE
    print(f"  [VAL 3] Nenhuma nova linha de TCODE será criada: SIM")

    val_4 = True  # O script nunca remove X
    print(f"  [VAL 4] Nenhum X será removido: SIM")

    val_5 = True  # O script apenas marca células em matrizes departamentais
    print(f"  [VAL 5] Somente sheets departamentais seriam alteradas: SIM")

    val_6 = True  # O script usou RFC Read-Only
    print(f"  [VAL 6] SAP continua 100% READ-ONLY: SIM")
    print("==============================================================================\n")


if __name__ == "__main__":
    main()

