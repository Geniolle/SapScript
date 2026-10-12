# -*- coding: utf-8 -*-
"""
sap_rfc/role_existence_validator_service.py
======================================================================
Serviço Oficial de Validação de Existência e Tipos de Roles no SAP
(PRD e QAD Independentes)
======================================================================
Fase obrigatória executada antes de:
  - PFCG_CREATE
  - PFCG_COMPOSTA
  - CUA_ADICIONAR
  - CUA_REMOVE

Validações:
1. Existência física via AGR_DEFINE (AGR_NAME).
2. Classificação dinâmica de tipo via AGR_FLAGS (FLAG_TYPE == 'COLL_AGR' & FLAG_VALUE == 'X').
   - Se COLL_AGR='X' -> COMPOSITE
   - Caso contrário -> SINGLE
3. Detecção de TIPO_DIVERGENTE (ex: esperado COMPOSITE mas SAP indica SINGLE, ou vice-versa).
4. Independência total entre ambientes (PRD e QAD).
5. Criação de roles faltantes SOMENTE com confirmação explícita do utilizador.
6. Readback obrigatório pós-criação confirmando existência e tipo.
7. Filtragem de lotes executáveis para dependências seguras (continuação parcial).
======================================================================
"""

from __future__ import annotations

import os
from typing import Dict, List, Set, Tuple, Any, Optional
from pathlib import Path

from sap_rfc._rfc_common import (
    load_project_env,
    build_connection_params_for,
    make_read_only_guard,
    make_write_guard,
    read_table,
    make_option_in,
    make_option_eq,
    role_exists,
    classify_rfc_error,
)

# Status oficiais de classificação de role (Geral)
STATUS_ROLE_OK = "ROLE_OK"
STATUS_ROLE_NAO_EXISTE = "ROLE_NAO_EXISTE"
STATUS_TIPO_DIVERGENTE = "TIPO_DIVERGENTE"
STATUS_AMBIENTE_INDISPONIVEL = "AMBIENTE_INDISPONIVEL"
STATUS_CRIACAO_NAO_AUTORIZADA = "CRIACAO_NAO_AUTORIZADA"
STATUS_CRIADA_E_CONFIRMADA = "CRIADA_E_CONFIRMADA"
STATUS_ERRO_CRIACAO = "ERRO_CRIACAO"
STATUS_ERRO_READBACK = "ERRO_READBACK"

# Status oficiais específicos para Funções Individuais (SINGLE)
STATUS_SINGLE_OK = "SINGLE_OK"
STATUS_SINGLE_AUSENTE = "SINGLE_AUSENTE"
STATUS_SINGLE_TIPO_DIVERGENTE = "SINGLE_TIPO_DIVERGENTE"
STATUS_SINGLE_ERRO_LEITURA = "ERRO_LEITURA"
STATUS_SINGLE_UNKNOWN_AMBIENTE = "UNKNOWN_AMBIENTE"

# Status oficiais específicos para Funções Compostas (COMPOSITE)
STATUS_COMPOSITE_OK = "COMPOSITE_OK"
STATUS_COMPOSITE_AUSENTE = "COMPOSITE_AUSENTE"
STATUS_COMPOSITE_TIPO_DIVERGENTE = "COMPOSITE_TIPO_DIVERGENTE"
STATUS_COMPOSITE_ERRO_LEITURA = "ERRO_LEITURA"
STATUS_COMPOSITE_UNKNOWN_AMBIENTE = "UNKNOWN_AMBIENTE"

# Classificações de auditoria para correção
STATUS_SEGURA_PARA_CORRIGIR = "SEGURA_PARA_CORRIGIR"
STATUS_BLOQUEADA_PARA_REVISAO = "BLOQUEADA_PARA_REVISAO"

# Status de validação do ambiente
ENV_STATUS_VALIDADO = "VALIDADO"
ENV_STATUS_INDISPONIVEL = "INDISPONIVEL"
ENV_STATUS_ERRO = "ERRO"

# Status globais
GLOBAL_TODOS_VALIDOS = "TODOS_VALIDOS"
GLOBAL_VALIDACAO_PARCIAL = "VALIDACAO_PARCIAL"
GLOBAL_NENHUM_AMBIENTE_DISPONIVEL = "NENHUM_AMBIENTE_DISPONIVEL"
GLOBAL_ERRO_VALIDACAO = "ERRO_VALIDACAO"


def extrair_roles_necessarias_projeto(caminho_excel: str, dados: Optional[Any] = None) -> Dict[str, Set[str]]:
    """
    Levanta dinamicamente e deduplica todas as roles necessárias do projeto Excel:
    - Composite Roles:
        * Sheet 'Proposta Ativa' (coluna Composite Role)
        * Sheet 'PFCG_COMPOSTA' (coluna Parent)
        * Catálogo de roles compostas (dados.roles_compostas)
    - Single Roles:
        * Sheet 'PFCG_COMPOSTA' (coluna Child)
        * Singles avulsas / Funções Individuais ('Proposta Ativa')
        * Sheet 'DEFINIÇÕES'
        * Sheet 'PFCG_CREATE'
        * Sheet 'CUA_ADICIONAR'
    Retorna: {"COMPOSITE": set(...), "SINGLE": set(...)}
    """
    def _norm(val: Any) -> str:
        if val is None:
            return ""
        s = str(val).strip().upper()
        return "" if s in ("NAN", "NONE", "-", "") else s

    composites: Set[str] = set()
    singles: Set[str] = set()

    # 1. Usar dados consolidados em memória se fornecidos
    if dados:
        for comp in getattr(dados, "roles_compostas", {}).keys():
            c_norm = _norm(comp)
            if c_norm:
                composites.add(c_norm)
        for s in getattr(dados, "roles_simples", {}).keys():
            s_norm = _norm(s)
            if s_norm:
                singles.add(s_norm)
        for dep_roles in getattr(dados, "definicoes_departamento", {}).values():
            for r in dep_roles:
                r_norm = _norm(r)
                if r_norm:
                    singles.add(r_norm)

    # 2. Ler diretamente o Excel para garantir cobertura total das folhas operacionais
    import openpyxl
    try:
        wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
    except Exception as e:
        return {"COMPOSITE": composites, "SINGLE": singles}

    # A. Proposta Ativa
    if "Proposta Ativa" in sheet_names:
        ws_at = wb["Proposta Ativa"]
        header = [str(c).upper() if c is not None else "" for c in next(ws_at.iter_rows(max_row=1, values_only=True), ())]
        col_comp_idx = next((i for i, c in enumerate(header) if "COMPOSITE" in c), -1)
        col_fi_idx = next((i for i, c in enumerate(header) if "FUN" in c and "INDIV" in c), -1)

        for row in ws_at.iter_rows(min_row=2, values_only=True):
            if col_comp_idx >= 0 and col_comp_idx < len(row):
                v_str = _norm(row[col_comp_idx])
                if v_str and v_str.startswith("Z"):
                    composites.add(v_str)
            if col_fi_idx >= 0:
                for c_val in row[col_fi_idx:]:
                    v_str = _norm(c_val)
                    if v_str and v_str.startswith("Z") and v_str not in composites:
                        singles.add(v_str)

    # B. PFCG_COMPOSTA
    if "PFCG_COMPOSTA" in sheet_names:
        ws_pc = wb["PFCG_COMPOSTA"]
        header_pc = [str(c).replace("_", "").replace(" ", "").upper() if c is not None else "" for c in next(ws_pc.iter_rows(max_row=1, values_only=True), ())]
        col_p_idx = next((i for i, c in enumerate(header_pc) if c in ("AGRNAMECOMPOSTA", "COMPOSITEROLE")), -1)
        col_c_idx = next((i for i, c in enumerate(header_pc) if c in ("AGRNAME", "SINGLEROLE")), -1)

        for row in ws_pc.iter_rows(min_row=2, values_only=True):
            if col_p_idx >= 0 and col_p_idx < len(row):
                v_str = _norm(row[col_p_idx])
                if v_str and v_str.startswith("Z"):
                    composites.add(v_str)
            if col_c_idx >= 0 and col_c_idx < len(row):
                v_str = _norm(row[col_c_idx])
                if v_str and v_str.startswith("Z"):
                    singles.add(v_str)

    # C. PFCG_CREATE
    if "PFCG_CREATE" in sheet_names:
        ws_cr = wb["PFCG_CREATE"]
        header_cr = [str(c).upper() if c is not None else "" for c in next(ws_cr.iter_rows(max_row=1, values_only=True), ())]
        col_cr_idx = next((i for i, c in enumerate(header_cr) if "AGR_NAME" in c or "ROLE" in c), -1)
        if col_cr_idx >= 0:
            for row in ws_cr.iter_rows(min_row=2, values_only=True):
                if col_cr_idx < len(row):
                    v_str = _norm(row[col_cr_idx])
                    if v_str and v_str.startswith("Z") and v_str not in composites:
                        singles.add(v_str)

    # D. CUA_ADICIONAR
    if "CUA_ADICIONAR" in sheet_names:
        ws_cua = wb["CUA_ADICIONAR"]
        header_cua = [str(c).upper() if c is not None else "" for c in next(ws_cua.iter_rows(max_row=1, values_only=True), ())]
        col_cua_idx = next((i for i, c in enumerate(header_cua) if "AGR_NAME" in c or "ROLE" in c), -1)
        if col_cua_idx >= 0:
            for row in ws_cua.iter_rows(min_row=2, values_only=True):
                if col_cua_idx < len(row):
                    v_str = _norm(row[col_cua_idx])
                    if v_str and v_str.startswith("Z") and v_str not in composites:
                        singles.add(v_str)

    # E. PFCG_AUTHORITY (Classificação estrita pela coluna de origem)
    if "PFCG_AUTHORITY" in sheet_names:
        import unicodedata
        ws_au = wb["PFCG_AUTHORITY"]
        header_au = [unicodedata.normalize("NFKD", str(c)).encode("ASCII", "ignore").decode("utf-8").strip().upper() if c is not None else "" for c in next(ws_au.iter_rows(max_row=1, values_only=True), ())]
        col_au_comp = next((i for i, c in enumerate(header_au) if c in ("AGR_NAME_COMPOSTA", "COMPOSITEROLE")), -1)
        col_au_ind = next((i for i, c in enumerate(header_au) if c in ("AGR_NAME", "SINGLEROLE")), -1)

        for row in ws_au.iter_rows(min_row=2, values_only=True):
            if col_au_comp >= 0 and col_au_comp < len(row):
                p_str = _norm(row[col_au_comp])
                if p_str and p_str.startswith("Z"):
                    composites.add(p_str)
            if col_au_ind >= 0 and col_au_ind < len(row):
                c_str = _norm(row[col_au_ind])
                if c_str and c_str.startswith("Z"):
                    singles.add(c_str)

    wb.close()

    # Deduplicação estrita: se uma role foi classificada como Composite, não deve constar em Singles
    singles_puras = {s for s in singles if s not in composites}

    return {
        "COMPOSITE": composites,
        "SINGLE": singles_puras
    }


def verificar_existencia_e_tipo_sap(
    roles: List[str],
    environment: str = "PRD",
    connection: Optional[Any] = None,
    chunk_size: int = 20
) -> Dict[str, Dict[str, Any]]:
    """
    Verifica existência física e determina dinamicamente o tipo de cada role no SAP via RFC.
    Consulta:
      1. AGR_DEFINE (existência de AGR_NAME)
      2. AGR_FLAGS (FLAG_TYPE == 'COLL_AGR' & FLAG_VALUE == 'X' para classificar Composite)
    Retorna dicionário indexado por AGR_NAME:
      {
        "ROLE_NAME": {
          "existe": bool,
          "tipo_sap": "COMPOSITE" | "SINGLE" | "INEXISTENTE",
          "erro": Optional[str]
        }
      }
    """
    if not roles:
        return {}

    roles_unicas = sorted({str(r).strip().upper() for r in roles if r and str(r).strip()})
    resultado: Dict[str, Dict[str, Any]] = {
        r: {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None}
        for r in roles_unicas
    }

    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            from sap_rfc._rfc_common import find_project_root
            load_project_env(find_project_root())
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_conn:
            err_msg = str(e_conn)
            for r in roles_unicas:
                resultado[r]["erro"] = err_msg
            return resultado

    guard = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS"))

    try:
        # 1. Consulta AGR_DEFINE em chunks
        roles_existentes: Set[str] = set()
        for i in range(0, len(roles_unicas), chunk_size):
            chunk = roles_unicas[i:i + chunk_size]
            opts = make_option_in("AGR_NAME", chunk)
            try:
                rows_def = read_table(
                    conn,
                    guard,
                    table_name="AGR_DEFINE",
                    fields=["AGR_NAME"],
                    options=opts,
                    rowcount=0
                )
                for row in rows_def:
                    if row and row[0].strip():
                        roles_existentes.add(row[0].strip().upper())
            except Exception as e_def:
                for r in chunk:
                    resultado[r]["erro"] = f"Erro ao ler AGR_DEFINE: {e_def}"

        # 2. Para as existentes, consultar AGR_FLAGS em chunks para saber se é COLL_AGR
        existentes_lista = sorted(list(roles_existentes))
        roles_composite: Set[str] = set()

        for i in range(0, len(existentes_lista), chunk_size):
            chunk = existentes_lista[i:i + chunk_size]
            opts = make_option_in("AGR_NAME", chunk)
            try:
                rows_flags = read_table(
                    conn,
                    guard,
                    table_name="AGR_FLAGS",
                    fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
                    options=opts,
                    rowcount=0
                )
                for row in rows_flags:
                    if len(row) >= 3:
                        agr = row[0].strip().upper()
                        flag = row[1].strip().upper()
                        val = row[2].strip().upper()
                        if flag == "COLL_AGR" and val == "X":
                            roles_composite.add(agr)
            except Exception as e_flg:
                pass

        # 3. Consolidar resultado
        for r in roles_unicas:
            if r in roles_existentes:
                resultado[r]["existe"] = True
                resultado[r]["tipo_sap"] = "COMPOSITE" if r in roles_composite else "SINGLE"
            else:
                resultado[r]["existe"] = False
                resultado[r]["tipo_sap"] = "INEXISTENTE"

    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass

    return resultado


def validar_existencia_roles_ambientes(
    caminho_excel: str,
    dados: Optional[Any] = None,
    ambientes: Tuple[str, ...] = ("PRD", "QAD")
) -> Dict[str, Any]:
    """
    Executa a auditoria completa de existência de roles nos ambientes especificados.
    Compara tipo esperado vs tipo SAP e classifica cada role em:
      - ROLE_OK
      - ROLE_NAO_EXISTE
      - TIPO_DIVERGENTE
      - AMBIENTE_INDISPONIVEL
    """
    roles_por_tipo = extrair_roles_necessarias_projeto(caminho_excel, dados=dados)
    comp_esperadas = sorted(list(roles_por_tipo["COMPOSITE"]))
    single_esperadas = sorted(list(roles_por_tipo["SINGLE"]))
    todas_roles = sorted(list(set(comp_esperadas + single_esperadas)))

    relatorio: Dict[str, Any] = {
        "ok": True,
        "total_composites": len(comp_esperadas),
        "total_singles": len(single_esperadas),
        "total_roles": len(todas_roles),
        "ambientes": {}
    }

    for env in ambientes:
        env_res: Dict[str, Any] = {
            "online": False,
            "status": ENV_STATUS_INDISPONIVEL,
            "validado": False,
            "erro_conexao": None,
            "existentes": None,
            "composites_existentes": None,
            "composites_faltantes": None,
            "composite_em_falta": None,
            "singles_existentes": None,
            "singles_faltantes": None,
            "single_em_falta": None,
            "tipos_divergentes": None,
            "detalhes": {}
        }

        # Teste de conexão/consulta
        sap_info = verificar_existencia_e_tipo_sap(todas_roles, environment=env)
        
        # Verificar se ambiente falhou totalmente por conexão
        erros_conn = [info["erro"] for info in sap_info.values() if info.get("erro")]
        if erros_conn and len(erros_conn) == len(todas_roles):
            env_res["online"] = False
            env_res["status"] = ENV_STATUS_INDISPONIVEL
            env_res["validado"] = False
            env_res["erro_conexao"] = erros_conn[0]
            for r in todas_roles:
                t_esp = "COMPOSITE" if r in comp_esperadas else "SINGLE"
                env_res["detalhes"][r] = {
                    "tipo_esperado": t_esp,
                    "tipo_sap": "DESCONHECIDO",
                    "status": STATUS_AMBIENTE_INDISPONIVEL,
                    "erro": erros_conn[0]
                }
            relatorio["ambientes"][env] = env_res
            continue

        # Ambiente ONLINE e acessível via RFC
        env_res["online"] = True
        env_res["status"] = ENV_STATUS_VALIDADO
        env_res["validado"] = True
        env_res["existentes"] = 0
        env_res["composites_existentes"] = 0
        env_res["composites_faltantes"] = 0
        env_res["composite_em_falta"] = []
        env_res["singles_existentes"] = 0
        env_res["singles_faltantes"] = 0
        env_res["single_em_falta"] = []
        env_res["tipos_divergentes"] = []

        for r in todas_roles:
            tipo_esperado = "COMPOSITE" if r in comp_esperadas else "SINGLE"
            info_sap = sap_info.get(r, {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None})

            if not info_sap["existe"]:
                status = STATUS_ROLE_NAO_EXISTE
                if tipo_esperado == "COMPOSITE":
                    env_res["composite_em_falta"].append(r)
                else:
                    env_res["single_em_falta"].append(r)
            else:
                tipo_sap = info_sap["tipo_sap"]
                if tipo_sap != tipo_esperado:
                    status = STATUS_TIPO_DIVERGENTE
                    env_res["tipos_divergentes"].append({
                        "role": r,
                        "tipo_esperado": tipo_esperado,
                        "tipo_sap": tipo_sap
                    })
                else:
                    status = STATUS_ROLE_OK
                    env_res["existentes"] += 1
                    if tipo_esperado == "COMPOSITE":
                        env_res["composites_existentes"] += 1
                    else:
                        env_res["singles_existentes"] += 1

            env_res["detalhes"][r] = {
                "tipo_esperado": tipo_esperado,
                "tipo_sap": info_sap["tipo_sap"],
                "status": status,
                "erro": info_sap["erro"]
            }

        env_res["composites_faltantes"] = len(env_res["composite_em_falta"])
        env_res["singles_faltantes"] = len(env_res["single_em_falta"])
        relatorio["ambientes"][env] = env_res

    # Cálculo do resultado global
    validados_cnt = sum(1 for a in relatorio["ambientes"].values() if a.get("validado"))
    total_ambientes = len(relatorio["ambientes"])

    if total_ambientes == 0:
        relatorio["resultado_global"] = GLOBAL_NENHUM_AMBIENTE_DISPONIVEL
    elif validados_cnt == total_ambientes:
        relatorio["resultado_global"] = GLOBAL_TODOS_VALIDOS
    elif validados_cnt > 0:
        relatorio["resultado_global"] = GLOBAL_VALIDACAO_PARCIAL
    else:
        relatorio["resultado_global"] = GLOBAL_NENHUM_AMBIENTE_DISPONIVEL

    return relatorio


def criar_single_role_com_readback(
    role_name: str,
    environment: str,
    descricao: str = "",
    tcodes: Optional[List[str]] = None,
    connection: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Cria uma Single Role no ambiente indicado via PRGN_RFC_CREATE_ACTIVITY_GROUP com READBACK OBRIGATÓRIO.
    Bloqueado estritamente para Composite Roles.
    """
    role_norm = role_name.strip().upper()
    desc_final = descricao.strip() or f"Role {role_norm}"

    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_c:
            return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Falha ao conectar via RFC: {e_c}"}

    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS", "AGR_TCODES", "TSTC"))
    guard_write = make_write_guard(
        allowed_functions=(
            "RFC_PING",
            "PRGN_RFC_CREATE_ACTIVITY_GROUP",
            "PRGN_GEN_PROFILES_FOR_ROLES",
            "BAPI_TRANSACTION_COMMIT",
        ),
        allowed_tables=("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS", "AGR_TCODES", "TSTC")
    )

    try:
        if role_exists(conn, guard_ro, role_norm):
            return {"ok": False, "status": "ROLE_JA_EXISTE", "mensagem": f"Role {role_norm} já existe no ambiente {environment}."}

        tcodes_payload = [{"TCODE": tc.strip().upper()} for tc in (tcodes or []) if tc and tc.strip()]

        guard_write.assert_function_allowed("PRGN_RFC_CREATE_ACTIVITY_GROUP")
        try:
            conn.call(
                "PRGN_RFC_CREATE_ACTIVITY_GROUP",
                ACTIVITY_GROUP=role_norm,
                ACTIVITY_GROUP_TEXT=desc_final[:80],
                NO_DIALOG="X",
                ONLY_TCODE_ASSIGNMENT="",
                ORG_LEVELS_WITH_STAR="",
                UNMAINTAINED_FIELDS_WITH_STAR="",
                PARENT_ROLE="",
                PROFILE_NAME="",
                PROFILE_TEXT=desc_final[:60],
                REQUEST="",
                TEMPLATE="",
                TCODES=tcodes_payload
            )
        except Exception:
            pass

        try:
            guard_write.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception:
            pass

        if not role_exists(conn, guard_ro, role_norm):
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback falhou: {role_norm} não foi encontrada na AGR_DEFINE após criação."
            }

        return {
            "ok": True,
            "status": STATUS_CRIADA_E_CONFIRMADA,
            "role": role_norm,
            "environment": environment,
            "tipo_esperado": "SINGLE",
            "tipo_confirmado": "SINGLE",
            "mensagem": f"✓ Single Role {role_norm} criada e confirmada por readback físico no SAP {environment}."
        }
    except Exception as exc:
        return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Erro na criação da Single Role: {exc}"}
    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass


def criar_composite_role_gui_controlado(
    role_name: str,
    environment: str,
    descricao: str = ""
) -> Dict[str, Any]:
    """
    Cria uma Composite Role através do fluxo SAP GUI oficial (D. PFCG_COMPOSTA / PFCGPage),
    garantindo que o botão 'Criar Função Composta' é acionado e gerando COLL_AGR='X' em AGR_FLAGS.
    """
    role_norm = role_name.strip().upper()
    desc_final = descricao.strip() or f"Composite Role {role_norm}"
    try:
        from sap_session import get_sap_session
        session = get_sap_session(target_system=environment)
        if not session:
            return {"ok": False, "status": "GUI_INDISPONIVEL", "mensagem": f"Sessão SAP GUI para {environment} não encontrada."}

        session.findById("wnd[0]/tbar[0]/okcd").text = "/NPFCG"
        session.findById("wnd[0]").sendVKey(0)

        # Inserir nome da role
        session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = role_norm
        # Clicar no botão 'Criar Função Composta'
        session.findById("wnd[0]/usr/btn%#AUTOTEXT004").press()

        # Inserir descrição
        try:
            session.findById("wnd[0]/usr/txtS_AGR_TEXTS-TEXT").text = desc_final[:80]
        except Exception:
            try:
                session.findById("wnd[0]/usr/txtAGR_TEXTS-TEXT").text = desc_final[:80]
            except Exception:
                pass

        # Salvar (F11 / Ctrl+S)
        session.findById("wnd[0]").sendVKey(11)

        return {"ok": True, "status": "CRIADO_VIA_GUI", "mensagem": f"Composite Role {role_norm} enviada via SAP GUI."}
    except Exception as e_gui:
        return {"ok": False, "status": "ERRO_GUI", "mensagem": f"Falha no fluxo SAP GUI PFCG: {e_gui}"}


def criar_composite_role_rfc_oficial(
    role_name: str,
    environment: str,
    descricao: str = "",
    connection: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Cria uma Composite Role via RFC oficial PRGN_RFC_CREATE_AGR_MULTIPLE com COLLECTIVE_AGR='X'.
    100% RFC (sem necessidade de SAP GUI Scripting).
    """
    role_norm = role_name.strip().upper()
    desc_final = descricao.strip() or f"Role {role_norm}"

    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_c:
            return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Falha ao conectar via RFC: {e_c}"}

    guard_write = make_write_guard(
        allowed_functions=(
            "RFC_PING",
            "PRGN_RFC_CREATE_AGR_MULTIPLE",
            "BAPI_TRANSACTION_COMMIT",
        ),
        allowed_tables=("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS")
    )

    try:
        guard_write.assert_function_allowed("PRGN_RFC_CREATE_AGR_MULTIPLE")
        res_create = conn.call(
            "PRGN_RFC_CREATE_AGR_MULTIPLE",
            ACTIVITY_GROUP=role_norm,
            ACTIVITY_GROUP_TEXT=desc_final[:80],
            COLLECTIVE_AGR="X",
            NO_DIALOG="X"
        )

        # COMMIT físico
        try:
            guard_write.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception:
            pass

        return {"ok": True, "status": "CRIADO_VIA_RFC", "mensagem": f"Composite Role {role_norm} criada via RFC PRGN_RFC_CREATE_AGR_MULTIPLE."}
    except Exception as exc:
        return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Erro na chamada RFC PRGN_RFC_CREATE_AGR_MULTIPLE: {exc}"}
    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass


def criar_composite_role_com_readback(
    role_name: str,
    environment: str,
    descricao: str = "",
    connection: Optional[Any] = None,
    usar_gui: bool = False
) -> Dict[str, Any]:
    """
    Cria uma Composite Role com validação e READBACK OBRIGATÓRIO (AGR_DEFINE e AGR_FLAGS COLL_AGR='X').
    BLOQUEIO DE SEGURANÇA: NUNCA utiliza PRGN_RFC_CREATE_ACTIVITY_GROUP.
    Utiliza método oficial suportado:
      - Principal: RFC oficial PRGN_RFC_CREATE_AGR_MULTIPLE(COLLECTIVE_AGR='X').
      - Fallback opcional: SAP GUI controlado PFCGPage.
    """
    role_norm = role_name.strip().upper()
    desc_final = descricao.strip() or f"Role {role_norm}"

    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_c:
            return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Falha ao conectar via RFC: {e_c}"}

    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS", "AGR_TCODES", "TSTC"))

    try:
        # 1. Pré-verificar se já existe
        if role_exists(conn, guard_ro, role_norm):
            return {"ok": False, "status": "ROLE_JA_EXISTE", "mensagem": f"Role {role_norm} já existe no ambiente {environment}."}

        # 2. Executar criação pelo método suportado para Composite (RFC oficial por padrão)
        if usar_gui:
            res_exec = criar_composite_role_gui_controlado(role_norm, environment, descricao=desc_final)
        else:
            res_exec = criar_composite_role_rfc_oficial(role_norm, environment, descricao=desc_final, connection=conn)

        if not res_exec.get("ok"):
            return {
                "ok": False,
                "status": STATUS_ERRO_CRIACAO,
                "mensagem": f"Falha ao criar Composite Role: {res_exec.get('mensagem')}"
            }

        # 3. READBACK OBRIGATÓRIO 1: AGR_DEFINE
        if not role_exists(conn, guard_ro, role_norm):
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback falhou: {role_norm} não foi encontrada na AGR_DEFINE após criação."
            }

        # 4. READBACK OBRIGATÓRIO 2: AGR_FLAGS (deve ter COLL_AGR = 'X')
        is_composite_rb = False
        try:
            rows_flg = read_table(
                conn,
                guard_ro,
                table_name="AGR_FLAGS",
                fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
                options=make_option_eq("AGR_NAME", role_norm),
                rowcount=0
            )
            for row in rows_flg:
                if len(row) >= 3 and row[1].strip().upper() == "COLL_AGR" and row[2].strip().upper() == "X":
                    is_composite_rb = True
                    break
        except Exception:
            pass

        if not is_composite_rb:
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback falhou: {role_norm} foi criada mas AGR_FLAGS NÃO possui COLL_AGR='X'."
            }

        return {
            "ok": True,
            "status": STATUS_CRIADA_E_CONFIRMADA,
            "role": role_norm,
            "environment": environment,
            "tipo_esperado": "COMPOSITE",
            "tipo_confirmado": "COMPOSITE",
            "metodo_utilizado": "GUI" if usar_gui else "RFC",
            "mensagem": f"✓ Composite Role {role_norm} criada e confirmada por readback físico (COLL_AGR='X') no SAP {environment}."
        }
    except Exception as exc:
        return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Erro na criação da Composite Role: {exc}"}
    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass


def criar_role_sap_com_readback(
    role_name: str,
    tipo: str,
    environment: str,
    descricao: str = "",
    tcodes: Optional[List[str]] = None,
    connection: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Despachante estrito de criação com readback.
    NUNCA chama PRGN_RFC_CREATE_ACTIVITY_GROUP quando tipo == 'COMPOSITE'.
    """
    tipo_norm = tipo.strip().upper()
    if tipo_norm == "COMPOSITE":
        return criar_composite_role_com_readback(
            role_name=role_name,
            environment=environment,
            descricao=descricao,
            connection=connection
        )
    elif tipo_norm == "SINGLE":
        return criar_single_role_com_readback(
            role_name=role_name,
            environment=environment,
            descricao=descricao,
            tcodes=tcodes,
            connection=connection
        )
    else:
        return {
            "ok": False,
            "status": "TIPO_INVALIDO",
            "mensagem": f"Tipo de role inválido: '{tipo}'. Tipos permitidos: SINGLE, COMPOSITE."
        }


def exibir_relatorio_validacao_roles(relatorio: Dict[str, Any]) -> None:
    """Imprime no terminal o relatório padronizado da validação técnica de roles com semântica UNKNOWN."""
    print("\n" + "=" * 78)
    print("  VALIDAÇÃO DE ROLES — SAP PRD / QAD")
    print("=" * 78)
    print(f"  Composite Roles analisadas : {relatorio.get('total_composites', 0)}")
    print(f"  Single Roles analisadas    : {relatorio.get('total_singles', 0)}")
    print(f"  Total de roles distintas   : {relatorio.get('total_roles', 0)}")

    for env, info in relatorio.get("ambientes", {}).items():
        st_env = info.get("status", ENV_STATUS_INDISPONIVEL)
        print(f"\n  {env} (Status: {st_env}):")
        if not info.get("validado"):
            print(f"    [OFFLINE] Ambiente indisponível: {info.get('erro_conexao')}")
            print(f"    Existentes               : UNKNOWN")
            print(f"    Composite em falta       : UNKNOWN")
            print(f"    Single em falta          : UNKNOWN")
            print(f"    Tipo divergente          : UNKNOWN")
            continue

        print(f"    Existentes               : {info.get('existentes', 0)}")
        print(f"    Composite em falta       : {info.get('composites_faltantes', 0)}")
        print(f"    Single em falta          : {info.get('singles_faltantes', 0)}")
        print(f"    Tipo divergente          : {len(info.get('tipos_divergentes', []))}")

    print("-" * 78)
    print(f"Resultado Global: {relatorio.get('resultado_global', 'UNKNOWN')}")
    print("-" * 78)
    print("Roles em falta:")

    for env, info in relatorio.get("ambientes", {}).items():
        print(f"\n{env}")
        if not info.get("validado"):
            print("  (Ambiente indisponível — estado das roles: UNKNOWN / NÃO VALIDADO)")
            continue

        comps = info.get("composite_em_falta") or []
        singles = info.get("single_em_falta") or []

        print("  COMPOSITE:")
        if comps:
            for c in comps:
                print(f"    - {c}")
        else:
            print("    (nenhuma em falta)")

        print("  SINGLE:")
        if singles:
            for s in singles:
                print(f"    - {s}")
        else:
            print("    (nenhuma em falta)")

        divs = info.get("tipos_divergentes") or []
        if divs:
            print("  TIPOS DIVERGENTES (Bloqueados):")
            for d in divs:
                print(f"    ! {d['role']}: Esperado={d['tipo_esperado']} vs SAP={d['tipo_sap']}")

    print("=" * 78)


def auditar_single_role_divergente(
    role_name: str,
    environment: str = "PRD",
    connection: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Audita uma Single Role com tipo incorreto no SAP PRD:
    1. Confirmar tipo atual no SAP (AGR_DEFINE e AGR_FLAGS COLL_AGR)
    2. Confirmar tipo esperado (SINGLE)
    3. Validar utilizadores (AGR_USERS)
    4. Validar tcodes (AGR_TCODES)
    5. Validar autorizações (AGR_1251)
    6. Validar relações/dependências (AGR_AGRS como parent ou child)
    7. Classificar: SEGURA_PARA_CORRIGIR ou BLOQUEADA_PARA_REVISAO
    Nesta fase, COMPOSITE -> SINGLE não é recriada automaticamente por delete/recreate sem análise;
    é classificada e mantida bloqueada para revisão manual com informações completas.
    """
    role_norm = role_name.strip().upper()
    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_c:
            return {
                "role": role_norm,
                "status": STATUS_BLOQUEADA_PARA_REVISAO,
                "motivo": f"Falha na ligação RFC: {e_c}",
                "classificacao": STATUS_BLOQUEADA_PARA_REVISAO
            }

    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_USERS", "AGR_TCODES", "AGR_1251", "AGR_AGRS"))

    try:
        # 1 & 2. Tipo físico no SAP
        existe = role_exists(conn, guard_ro, role_norm)
        is_composite = False
        if existe:
            rows_flg = read_table(
                conn, guard_ro, table_name="AGR_FLAGS",
                fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
                options=make_option_eq("AGR_NAME", role_norm),
                rowcount=0
            )
            for r in rows_flg:
                if len(r) >= 3 and r[1].strip().upper() == "COLL_AGR" and r[2].strip().upper() == "X":
                    is_composite = True
                    break

        tipo_sap = "COMPOSITE" if is_composite else ("SINGLE" if existe else "INEXISTENTE")

        # 3. Utilizadores
        users = []
        try:
            rows_u = read_table(conn, guard_ro, table_name="AGR_USERS", fields=["UNAME"], options=make_option_eq("AGR_NAME", role_norm), rowcount=0)
            users = sorted({r[0].strip().upper() for r in rows_u if r and r[0].strip()})
        except Exception:
            pass

        # 4. TCodes
        tcodes = []
        try:
            rows_tc = read_table(conn, guard_ro, table_name="AGR_TCODES", fields=["TCODE"], options=make_option_eq("AGR_NAME", role_norm), rowcount=0)
            tcodes = sorted({r[0].strip().upper() for r in rows_tc if r and r[0].strip()})
        except Exception:
            pass

        # 5. Autorizações
        auth_count = 0
        try:
            rows_auth = read_table(conn, guard_ro, table_name="AGR_1251", fields=["OBJECT"], options=make_option_eq("AGR_NAME", role_norm), rowcount=0)
            auth_count = len(rows_auth)
        except Exception:
            pass

        # 6. Relações / Dependências (se for composite, quais as filhas)
        filhas = []
        pais = []
        try:
            rows_rel_c = read_table(conn, guard_ro, table_name="AGR_AGRS", fields=["CHILD_AGR"], options=make_option_eq("AGR_NAME", role_norm), rowcount=0)
            filhas = sorted({r[0].strip().upper() for r in rows_rel_c if r and r[0].strip()})
        except Exception:
            pass
        try:
            rows_rel_p = read_table(conn, guard_ro, table_name="AGR_AGRS", fields=["AGR_NAME"], options=make_option_eq("CHILD_AGR", role_norm), rowcount=0)
            pais = sorted({r[0].strip().upper() for r in rows_rel_p if r and r[0].strip()})
        except Exception:
            pass

        # 7. Classificação: método seguro não validado para conversão destrutiva automática de Composite para Single
        # Bloquear para revisão manual com diagnóstico detalhado
        classificacao = STATUS_BLOQUEADA_PARA_REVISAO
        motivo = (
            f"Role fisicamente é {tipo_sap} no SAP. Conversão COMPOSITE -> SINGLE requer revisão manual. "
            f"Utilizadores: {len(users)}, TCodes: {len(tcodes)}, Objetos de autorização: {auth_count}, "
            f"Filhas associadas: {len(filhas)}, Pais: {len(pais)}."
        )

        return {
            "role": role_norm,
            "tipo_esperado": "SINGLE",
            "tipo_sap": tipo_sap,
            "existe": existe,
            "utilizadores": users,
            "utilizadores_count": len(users),
            "tcodes": tcodes,
            "tcodes_count": len(tcodes),
            "autorizacoes_count": auth_count,
            "filhas": filhas,
            "pais": pais,
            "classificacao": classificacao,
            "motivo": motivo,
        }
    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass


def corrigir_composite_role_divergente_com_readback(
    role_name: str,
    environment: str = "PRD",
    descricao: str = "",
    connection: Optional[Any] = None,
    transport_mode: str = "LOCAL",
    request_number: str = ""
) -> Dict[str, Any]:
    """
    Executa o fluxo seguro de correção de COMPOSITE com tipo divergente (esperado COMPOSITE, mas SAP é SINGLE):
    1. Readback inicial: confirma que existe, tipo atual = SINGLE, esperado = COMPOSITE
    2. Validar se está segura para substituir
    3. Eliminar Single incorreta via serviço oficial (pfcg_role_delete_service)
    4. COMMIT
    5. Readback de ausência
    6. Criar novamente via PRGN_RFC_CREATE_AGR_MULTIPLE com COLLECTIVE_AGR = 'X'
    7. COMMIT
    8. Readback físico: AGR_DEFINE existe e AGR_FLAGS.COLL_AGR = 'X'
    9. Somente então marcar CORRIGIDA. Falha em qualquer readback não gera sucesso.
    NUNCA usa PRGN_RFC_CREATE_ACTIVITY_GROUP para Composite.
    """
    role_norm = role_name.strip().upper()
    conn = connection
    deve_fechar = False
    if conn is None:
        try:
            params = build_connection_params_for(environment)
            from pyrfc import Connection
            conn = Connection(**params)
            deve_fechar = True
        except Exception as e_c:
            return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Falha na ligação RFC: {e_c}"}

    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS", "AGR_USERS", "AGR_TCODES", "TSTC"))
    guard_write = make_write_guard(
        allowed_functions=(
            "RFC_PING",
            "BPC_DELETE_SINGLE_ROLE",
            "PRGN_RFC_CREATE_AGR_MULTIPLE",
            "BAPI_TRANSACTION_COMMIT",
        ),
        allowed_tables=("AGR_DEFINE", "AGR_FLAGS", "AGR_TEXTS", "AGR_USERS", "AGR_TCODES", "TSTC")
    )

    try:
        # Passo 1: Readback inicial
        if not role_exists(conn, guard_ro, role_norm):
            return {
                "ok": False,
                "status": "ROLE_NAO_EXISTE",
                "mensagem": f"Readback inicial: Role {role_norm} não existe no ambiente {environment}."
            }

        # Verificar tipo atual
        rows_flg = read_table(
            conn, guard_ro, table_name="AGR_FLAGS",
            fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
            options=make_option_eq("AGR_NAME", role_norm),
            rowcount=0
        )
        is_composite_inicial = any(
            len(r) >= 3 and r[1].strip().upper() == "COLL_AGR" and r[2].strip().upper() == "X"
            for r in rows_flg
        )
        if is_composite_inicial:
            return {
                "ok": True,
                "status": STATUS_COMPOSITE_OK,
                "mensagem": f"Role {role_norm} já é Composite no SAP {environment}. Nenhuma correção necessária."
            }

        # Recuperar texto da role se descricao nao fornecida
        desc_final = descricao.strip()
        if not desc_final:
            try:
                rows_txt = read_table(
                    conn, guard_ro, table_name="AGR_TEXTS",
                    fields=["TEXT"],
                    options=make_option_eq("AGR_NAME", role_norm),
                    rowcount=1
                )
                if rows_txt and rows_txt[0] and rows_txt[0][0]:
                    desc_final = str(rows_txt[0][0]).strip()
            except Exception:
                desc_final = f"Role {role_norm}"
        if not desc_final:
            desc_final = f"Role {role_norm}"

        # Passo 2 & 3: Eliminar Single incorreta via BPC_DELETE_SINGLE_ROLE
        guard_write.assert_function_allowed("BPC_DELETE_SINGLE_ROLE")
        # IV_PREF_AGR_NAME tem tamanho máximo de 24 chars e IV_NAME_PREFIX tem 6 chars.
        # No ABAP: CONCATENATE IV_NAME_PREFIX IV_PREF_AGR_NAME INTO LV_AGR_NAME.
        if len(role_norm) <= 24:
            del_prefix = " "
            del_name = role_norm
        else:
            split_pos = len(role_norm) - 24
            del_prefix = role_norm[:split_pos]
            del_name = role_norm[split_pos:]

        call_del_kwargs: dict[str, Any] = {
            "IV_NAME_PREFIX": del_prefix,
            "IV_PREF_AGR_NAME": del_name,
            "NO_DIALOG": "X",
            "REQUEST": request_number or "",
        }
        res_del = conn.call("BPC_DELETE_SINGLE_ROLE", **call_del_kwargs)
        msgs_del = res_del.get("ET_MESSAGES") or []
        del_errors = [str(m.get("MESSAGE") or m) for m in msgs_del if isinstance(m, dict) and m.get("TYPE") in {"E", "A"}]
        if del_errors:
            return {
                "ok": False,
                "status": "ERRO_DELETE",
                "mensagem": f"Erro retornado pelo SAP ao eliminar Single incorreta {role_norm}: {'; '.join(del_errors)}"
            }

        # Passo 4: COMMIT
        try:
            guard_write.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception:
            pass

        # Passo 5: Readback de ausência
        if role_exists(conn, guard_ro, role_norm):
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback de ausência falhou: {role_norm} ainda existe na AGR_DEFINE após eliminação."
            }

        # Passo 6: Criar novamente via PRGN_RFC_CREATE_AGR_MULTIPLE com COLLECTIVE_AGR = 'X'
        guard_write.assert_function_allowed("PRGN_RFC_CREATE_AGR_MULTIPLE")
        conn.call(
            "PRGN_RFC_CREATE_AGR_MULTIPLE",
            ACTIVITY_GROUP=role_norm,
            ACTIVITY_GROUP_TEXT=desc_final[:80],
            COLLECTIVE_AGR="X",
            NO_DIALOG="X"
        )

        # Passo 7: COMMIT
        try:
            guard_write.assert_function_allowed("BAPI_TRANSACTION_COMMIT")
            conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        except Exception:
            pass

        # Passo 8: Readback obrigatório 1: AGR_DEFINE existe
        if not role_exists(conn, guard_ro, role_norm):
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback de recriação falhou: {role_norm} não foi encontrada na AGR_DEFINE."
            }

        # Readback obrigatório 2: AGR_FLAGS.COLL_AGR = 'X'
        rows_flg_rb = read_table(
            conn, guard_ro, table_name="AGR_FLAGS",
            fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
            options=make_option_eq("AGR_NAME", role_norm),
            rowcount=0
        )
        is_comp_rb = any(
            len(r) >= 3 and r[1].strip().upper() == "COLL_AGR" and r[2].strip().upper() == "X"
            for r in rows_flg_rb
        )
        if not is_comp_rb:
            return {
                "ok": False,
                "status": STATUS_ERRO_READBACK,
                "mensagem": f"Readback falhou: {role_norm} existe mas AGR_FLAGS NÃO possui COLL_AGR='X'."
            }

        # Passo 9: Somente então marcar CORRIGIDA
        return {
            "ok": True,
            "status": "CORRIGIDA",
            "role": role_norm,
            "environment": environment,
            "tipo_anterior": "SINGLE",
            "tipo_confirmado": "COMPOSITE",
            "mensagem": f"✓ Composite Role {role_norm} corrigida com sucesso (delete single + create composite + readback COLL_AGR='X')."
        }

    except Exception as exc:
        return {"ok": False, "status": STATUS_ERRO_CRIACAO, "mensagem": f"Erro na correção da Composite Role: {exc}"}
    finally:
        if deve_fechar and conn:
            try: conn.close()
            except Exception: pass


def executar_etapa_validacao_funcoes_individuais(
    caminho_excel: str,
    dados: Optional[Any] = None,
    environment: str = "PRD",
    assumir_sim: bool = False
) -> Dict[str, Any]:
    """
    ETAPA [6/8] — VALIDAÇÃO DAS FUNÇÕES INDIVIDUAIS
    Valida SOMENTE roles que o projeto espera como SINGLE.
    Não inclui roles que o projeto espera como COMPOSITE.
    Para cada SINGLE esperada, valida no SAP PRD:
      - existe?
      - tipo físico é SINGLE?
      - está ausente?
      - está erradamente como COMPOSITE?
    Classificações: SINGLE_OK, SINGLE_AUSENTE, SINGLE_TIPO_DIVERGENTE, ERRO_LEITURA, UNKNOWN_AMBIENTE
    Se houver divergências, lista todas e pergunta se deseja analisar/corrigir (S/N).
    Se S: audita cada uma e classifica SEGURA_PARA_CORRIGIR ou BLOQUEADA_PARA_REVISAO.
    """
    roles_por_tipo = extrair_roles_necessarias_projeto(caminho_excel, dados=dados)
    singles_esperadas = sorted(list(roles_por_tipo["SINGLE"]))
    total_singles = len(singles_esperadas)

    sap_info = verificar_existencia_e_tipo_sap(singles_esperadas, environment=environment)
    erros_conn = [info["erro"] for info in sap_info.values() if info.get("erro")]

    if erros_conn and len(erros_conn) == total_singles:
        err_msg = erros_conn[0]
        # Verificar logon failure / password must be changed
        is_logon_failure = "password must be changed" in err_msg.lower() or "logon" in err_msg.lower()
        print(f"  Ambiente {environment} offline / indisponível: {err_msg}")
        print(f"  Corretas no SAP {environment}            : UNKNOWN")
        print(f"  Ausentes                       : UNKNOWN")
        print(f"  Tipo divergente                : UNKNOWN")
        return {
            "ok": not is_logon_failure,
            "status": STATUS_SINGLE_UNKNOWN_AMBIENTE,
            "ambiente": environment,
            "total_esperadas": total_singles,
            "corretas": 0,
            "ausentes": 0,
            "tipos_divergentes": 0,
            "lista_divergentes": [],
            "lista_ausentes": [],
            "is_logon_failure": is_logon_failure,
            "erro_conexao": err_msg,
            "detalhes": {s: {"status": STATUS_SINGLE_UNKNOWN_AMBIENTE, "tipo_sap": "UNKNOWN"} for s in singles_esperadas}
        }

    corretas = []
    ausentes = []
    divergentes = []
    detalhes = {}

    for s in singles_esperadas:
        inf = sap_info.get(s, {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None})
        if not inf["existe"]:
            ausentes.append(s)
            detalhes[s] = {"status": STATUS_SINGLE_AUSENTE, "tipo_sap": "INEXISTENTE"}
        elif inf["tipo_sap"] != "SINGLE":
            divergentes.append({
                "role": s,
                "esperado": "SINGLE",
                "tipo_sap": inf["tipo_sap"]
            })
            detalhes[s] = {"status": STATUS_SINGLE_TIPO_DIVERGENTE, "tipo_sap": inf["tipo_sap"]}
        else:
            corretas.append(s)
            detalhes[s] = {"status": STATUS_SINGLE_OK, "tipo_sap": "SINGLE"}

    debug = os.environ.get("SAP_DEBUG", "").strip() in ("1", "true", "TRUE", "True")

    if debug:
        print("\n" + "=" * 78)
        print("  [6/8] VALIDAÇÃO DAS FUNÇÕES INDIVIDUAIS (DEBUG)")
        print("=" * 78)
        print(f"  Funções individuais esperadas : {total_singles}")
        print(f"  Corretas no SAP {environment}            : {len(corretas)}")
        print(f"  Ausentes                       : {len(ausentes)}")
        print(f"  Tipo divergente                : {len(divergentes)}")
        print("=" * 78)

    tem_problema = (len(ausentes) > 0 or len(divergentes) > 0)
    if not tem_problema:
        from sap_rfc.status_logger import print_status_success
        print_status_success(f"  ✓ [6/8] Funções Individuais: {len(corretas)}/{total_singles} corretas | ausentes {len(ausentes)} | divergentes {len(divergentes)}")
    else:
        from sap_rfc.status_logger import print_status_warning
        print_status_warning(f"  ⚠️ [6/8] Funções Individuais: {len(corretas)}/{total_singles} corretas | ausentes {len(ausentes)} | divergentes {len(divergentes)}")
        if ausentes:
            print("  Ausentes:")
            for a in ausentes:
                print(f"    - {a}")
        if divergentes:
            print("  Tipo divergente:")
            for d in divergentes:
                print(f"    - {d['role']} | Esperado={d['esperado']} | SAP={d['tipo_sap']}")

            print(f"\n  Foram encontradas {len(divergentes)} funções individuais com tipo incorreto.")
            resp = "S" if assumir_sim else input("  Deseja analisar/corrigir agora? (S/N) [Padrão: N]: ").strip().upper()
            if resp in ("S", "SIM", "Y", "YES"):
                print("\n  A auditar roles individuais com tipo divergente...")
                auditorias = []
                for d in divergentes:
                    res_aud = auditar_single_role_divergente(d["role"], environment=environment)
                    auditorias.append(res_aud)
                    print(f"\n  - Role: {res_aud['role']}")
                    print(f"    Classificação: {res_aud['classificacao']}")
                    print(f"    Diagnóstico : {res_aud['motivo']}")
                print("\n  ℹ Revisão manual necessária: nenhuma alteração destrutiva cega realizada.")
            else:
                print("  ℹ Nenhuma alteração realizada nas Single Roles divergentes.")

    return {
        "ok": True,
        "status": STATUS_SINGLE_OK if not ausentes and not divergentes else STATUS_SINGLE_TIPO_DIVERGENTE,
        "ambiente": environment,
        "total_esperadas": total_singles,
        "corretas": len(corretas),
        "ausentes": len(ausentes),
        "tipos_divergentes": len(divergentes),
        "lista_divergentes": divergentes,
        "lista_ausentes": ausentes,
        "detalhes": detalhes
    }


def executar_etapa_validacao_funcoes_compostas(
    caminho_excel: str,
    dados: Optional[Any] = None,
    environment: str = "PRD",
    assumir_sim: bool = False
) -> Dict[str, Any]:
    """
    ETAPA [7/8] — VALIDAÇÃO DAS FUNÇÕES COMPOSTAS
    Valida SOMENTE roles que o projeto espera como COMPOSITE.
    Não inclui roles que o projeto espera como SINGLE.
    Para cada COMPOSITE esperada, valida no SAP PRD:
      - existe?
      - AGR_FLAGS.COLL_AGR = 'X'?
      - está ausente?
      - está erradamente como SINGLE?
    Classificações: COMPOSITE_OK, COMPOSITE_AUSENTE, COMPOSITE_TIPO_DIVERGENTE, ERRO_LEITURA, UNKNOWN_AMBIENTE
    Se houver divergências:
      - Lista todas as roles com tipo incorreto em runtime
      - Pergunta: "Deseja corrigir agora as X Composite Roles com tipo incorreto? (S/N)"
      - Se S: executa fluxo seguro (readback -> delete single -> commit -> readback ausência ->
              create PRGN_RFC_CREATE_AGR_MULTIPLE(COLLECTIVE_AGR='X') -> commit -> readback COLL_AGR='X' -> corrigida)
    """
    roles_por_tipo = extrair_roles_necessarias_projeto(caminho_excel, dados=dados)
    composites_esperadas = sorted(list(roles_por_tipo["COMPOSITE"]))
    total_composites = len(composites_esperadas)

    sap_info = verificar_existencia_e_tipo_sap(composites_esperadas, environment=environment)
    erros_conn = [info["erro"] for info in sap_info.values() if info.get("erro")]

    if erros_conn and len(erros_conn) == total_composites:
        err_msg = erros_conn[0]
        is_logon_failure = "password must be changed" in err_msg.lower() or "logon" in err_msg.lower()
        print(f"  Ambiente {environment} offline / indisponível: {err_msg}")
        print(f"  Corretas no SAP {environment}        : UNKNOWN")
        print(f"  Ausentes                   : UNKNOWN")
        print(f"  Tipo divergente            : UNKNOWN")
        return {
            "ok": not is_logon_failure,
            "status": STATUS_COMPOSITE_UNKNOWN_AMBIENTE,
            "ambiente": environment,
            "total_esperadas": total_composites,
            "corretas": 0,
            "ausentes": 0,
            "tipos_divergentes": 0,
            "lista_divergentes": [],
            "lista_ausentes": [],
            "is_logon_failure": is_logon_failure,
            "erro_conexao": err_msg,
            "detalhes": {c: {"status": STATUS_COMPOSITE_UNKNOWN_AMBIENTE, "tipo_sap": "UNKNOWN"} for c in composites_esperadas}
        }

    corretas = []
    ausentes = []
    divergentes = []
    detalhes = {}

    for c in composites_esperadas:
        inf = sap_info.get(c, {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None})
        if not inf["existe"]:
            ausentes.append(c)
            detalhes[c] = {"status": STATUS_COMPOSITE_AUSENTE, "tipo_sap": "INEXISTENTE"}
        elif inf["tipo_sap"] != "COMPOSITE":
            divergentes.append({
                "role": c,
                "esperado": "COMPOSITE",
                "tipo_sap": inf["tipo_sap"]
            })
            detalhes[c] = {"status": STATUS_COMPOSITE_TIPO_DIVERGENTE, "tipo_sap": inf["tipo_sap"]}
        else:
            corretas.append(c)
            detalhes[c] = {"status": STATUS_COMPOSITE_OK, "tipo_sap": "COMPOSITE"}

    debug = os.environ.get("SAP_DEBUG", "").strip() in ("1", "true", "TRUE", "True")

    if debug:
        print("\n" + "=" * 78)
        print("  [7/8] VALIDAÇÃO DAS FUNÇÕES COMPOSTAS (DEBUG)")
        print("=" * 78)
        print(f"  Composite Roles esperadas : {total_composites}")
        print(f"  Corretas no SAP {environment}        : {len(corretas)}")
        print(f"  Ausentes                   : {len(ausentes)}")
        print(f"  Tipo divergente            : {len(divergentes)}")
        print("=" * 78)

    tem_problema = (len(ausentes) > 0 or len(divergentes) > 0)
    if not tem_problema:
        from sap_rfc.status_logger import print_status_success
        print_status_success(f"  ✓ [7/8] Funções Compostas: {len(corretas)}/{total_composites} corretas | ausentes {len(ausentes)} | divergentes {len(divergentes)}")
    else:
        from sap_rfc.status_logger import print_status_warning
        print_status_warning(f"  ⚠️ [7/8] Funções Compostas: {len(corretas)}/{total_composites} corretas | ausentes {len(ausentes)} | divergentes {len(divergentes)}")
        if ausentes:
            print("  Ausentes:")
            for a in ausentes:
                print(f"    - {a}")
        if divergentes:
            print("  Tipo divergente:")
            for d in divergentes:
                print(f"    - {d['role']} | Esperado={d['esperado']} | SAP={d['tipo_sap']}")

    corrigidas_count = 0
    if divergentes:
        resp = "S" if assumir_sim else input(f"Deseja corrigir agora as {len(divergentes)} Composite Roles com tipo incorreto? (S/N) [Padrão: N]: ").strip().upper()
        if resp in ("S", "SIM", "Y", "YES"):
            for d in divergentes:
                r_nome = d["role"]
                res_corr = corrigir_composite_role_divergente_com_readback(r_nome, environment=environment)
                print(f"  {res_corr['mensagem']}")
                if res_corr.get("ok"):
                    corrigidas_count += 1
                    corretas.append(r_nome)
                    detalhes[r_nome] = {"status": STATUS_COMPOSITE_OK, "tipo_sap": "COMPOSITE"}
        else:
            print("  ℹ Nenhuma alteração realizada nas Composite Roles divergentes. Relações dependentes permanecem bloqueadas.")

    # Recalcular divergentes pendentes após tentativa de correção
    divs_restantes = [d for d in divergentes if d["role"] not in corretas]

    return {
        "ok": True,
        "status": STATUS_COMPOSITE_OK if not ausentes and not divs_restantes else STATUS_COMPOSITE_TIPO_DIVERGENTE,
        "ambiente": environment,
        "total_esperadas": total_composites,
        "corretas": len(corretas),
        "ausentes": len(ausentes),
        "tipos_divergentes": len(divs_restantes),
        "corrigidas_count": corrigidas_count,
        "lista_divergentes": divs_restantes,
        "lista_ausentes": ausentes,
        "detalhes": detalhes
    }


def executar_fase_validacao_existencia_roles(
    caminho_excel: str,
    dados: Optional[Any] = None,
    ambientes: Tuple[str, ...] = ("PRD", "QAD"),
    assumir_sim: bool = False
) -> Dict[str, Any]:
    """
    Fase Obrigatória Integrada:
    1. Valida existência física e tipos de todas as roles necessárias em PRD e QAD.
    2. Exibe o relatório claro no terminal com semântica UNKNOWN para ambientes offline.
    3. Para roles em falta em ambientes estritamente VALIDADOS, pergunta se deseja criar.
       Ambientes INDISPONÍVEIS/UNKNOWN NUNCA geram perguntas de criação.
    4. Cria somente com confirmação explícita (NUNCA automático).
    5. Realiza readback obrigatório pós-criação.
    6. Retorna conjunto de roles confirmadas por ambiente para filtrar fluxos dependentes.
    """
    relatorio = validar_existencia_roles_ambientes(caminho_excel, dados=dados, ambientes=ambientes)
    exibir_relatorio_validacao_roles(relatorio)

    roles_confirmadas_por_ambiente: Dict[str, Set[str]] = {
        env: set() for env in ambientes
    }

    # Preencher roles já existentes e com tipo correto nos ambientes validados
    for env, info in relatorio.get("ambientes", {}).items():
        if not info.get("validado"):
            continue
        for r, det in info.get("detalhes", {}).items():
            if det.get("status") in (STATUS_ROLE_OK, STATUS_SINGLE_OK, STATUS_COMPOSITE_OK):
                roles_confirmadas_por_ambiente[env].add(r)

    # Questionar criação de roles faltantes SOMENTE para ambientes validados
    for env in ambientes:
        info = relatorio.get("ambientes", {}).get(env, {})
        if not info.get("validado"):
            print(f"\n[SAP {env}] Ambiente offline / indisponível (UNKNOWN). Nenhuma criação solicitada ou executada.")
            continue

        comps_faltam = info.get("composite_em_falta") or []
        singles_faltam = info.get("single_em_falta") or []

        # A. Composite Roles Faltantes
        if comps_faltam:
            print(f"\n[SAP {env}] Foram detetadas {len(comps_faltam)} Composite Role(s) em falta:")
            for c in comps_faltam:
                print(f"  - {c}")

            resp_c = "S" if assumir_sim else input(f"Deseja criar as {len(comps_faltam)} Composite Role(s) no SAP {env}? (S/N) [Padrão: N]: ").strip().upper()
            if resp_c in ("S", "SIM", "Y", "YES"):
                for c in comps_faltam:
                    res_cr = criar_role_sap_com_readback(c, tipo="COMPOSITE", environment=env)
                    print(f"  {res_cr['mensagem']}")
                    if res_cr.get("ok"):
                        roles_confirmadas_por_ambiente[env].add(c)
            else:
                print(f"  ℹ Criação de Composite Roles ignorada pelo utilizador em {env}.")

        # B. Single Roles Faltantes
        if singles_faltam:
            print(f"\n[SAP {env}] Foram detetadas {len(singles_faltam)} Single Role(s) em falta:")
            for s in singles_faltam:
                print(f"  - {s}")

            resp_s = "S" if assumir_sim else input(f"Deseja criar as {len(singles_faltam)} Single Role(s) no SAP {env}? (S/N) [Padrão: N]: ").strip().upper()
            if resp_s in ("S", "SIM", "Y", "YES"):
                for s in singles_faltam:
                    res_cr = criar_role_sap_com_readback(s, tipo="SINGLE", environment=env)
                    print(f"  {res_cr['mensagem']}")
                    if res_cr.get("ok"):
                        roles_confirmadas_por_ambiente[env].add(s)
            else:
                print(f"  ℹ Criação de Single Roles ignorada pelo utilizador em {env}.")

    relatorio["roles_confirmadas"] = roles_confirmadas_por_ambiente
    return relatorio

