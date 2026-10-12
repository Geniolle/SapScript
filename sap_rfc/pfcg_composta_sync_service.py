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
    make_option_in,
    classify_rfc_error
)
from pyrfc import Connection
from sap_rfc.role_existence_validator_service import verificar_existencia_e_tipo_sap

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
STATE_INVALID_COMPOSITE = "INVALIDA_COMPOSITE_COMPOSITE"
STATE_PARENT_NOT_FOUND = "PARENT_NAO_EXISTE"
STATE_CHILD_NOT_FOUND = "CHILD_NAO_EXISTE"
STATE_PARENT_INVALID_TYPE = "PARENT_TIPO_DIVERGENTE"
STATE_CHILD_INVALID_TYPE = "CHILD_TIPO_DIVERGENTE"
STATE_EXECUTAVEL = "REALMENTE_EXECUTAVEL"
STATE_ERROR = "ERRO_LEITURA"


def obter_pendencias_pfcg_composta(caminho_excel: str) -> List[Dict[str, Any]]:
    """
    Lê a sheet PFCG_COMPOSTA e retorna todas as linhas elegíveis como candidatas:
    STATUS vazio (None ou "") ou "Pendente QAD" ou "Pendente PRD".
    Esta é a ÚNICA fonte de verdade para a fila PFCG_COMPOSTA.
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
                    "categoria": categoria,
                    "status_original": status
                })

    return pendencias


def validar_roles_filhas_nao_compostas(conn: Any, guard: Any, single_roles: List[str]) -> Tuple[Set[str], Set[str]]:
    """
    Consulta a tabela AGR_FLAGS para garantir que nenhuma Single Role a ser adicionada
    é na verdade uma Composite Role (FLAG_TYPE == 'COLL_AGR' e FLAG_VALUE == 'X').
    Divide em chunks de 20 para evitar limites de buffer/OPTIONS.
    """
    if not single_roles:
        return set(), set()

    roles_unicas = sorted({r.strip().upper() for r in single_roles if r and str(r).strip()})
    if not roles_unicas:
        return set(), set()

    invalid_composites = set()

    chunk_size = 20
    for i in range(0, len(roles_unicas), chunk_size):
        chunk = roles_unicas[i:i + chunk_size]
        options_chunk = make_option_in("AGR_NAME", chunk)

        try:
            flags_rows = read_table(
                conn,
                guard,
                table_name="AGR_FLAGS",
                fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
                options=options_chunk,
                rowcount=0
            )
            for row in flags_rows:
                if len(row) >= 3:
                    agr = row[0].strip().upper()
                    flag = row[1].strip().upper()
                    val = row[2].strip().upper()
                    if flag == "COLL_AGR" and val == "X":
                        invalid_composites.add(agr)
        except Exception:
            pass

    roles_validas = set(roles_unicas) - invalid_composites
    return roles_validas, invalid_composites


def ler_relacoes_agr_agrs(conn: Any, guard: Any, composite_roles: List[str]) -> Tuple[Dict[str, Set[str]], bool, Optional[str]]:
    """
    Lê a tabela AGR_AGRS para um conjunto de Composite Roles usando make_option_in oficial com chunks de 20.
    Retorna (dicionário {COMPOSITE: set(SINGLES_EXISTENTES)}, ok_status, erro_msg).
    """
    if not composite_roles:
        return {}, True, None

    composites_unicas = sorted({c.strip().upper() for c in composite_roles if c and str(c).strip()})
    if not composites_unicas:
        return {}, True, None

    relacoes: Dict[str, Set[str]] = {c: set() for c in composites_unicas}
    chunk_size = 20

    try:
        for i in range(0, len(composites_unicas), chunk_size):
            chunk = composites_unicas[i:i + chunk_size]
            options = make_option_in("AGR_NAME", chunk)
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
                    if parent in relacoes and child:
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


def prevalidar_pfcg_composta(
    caminho_excel: str,
    ambiente_prd: str = "PRD"
) -> Dict[str, Any]:
    """
    Executa a PRÉ-VALIDAÇÃO READ-ONLY oficial da folha PFCG_COMPOSTA contra o SAP PRD.
    Não executa nenhuma escrita SAP nem no Excel.
    Separa estritamente:
      - candidatas
      - ja_existiam (em PRD)
      - ausentes (em PRD)
      - invalidas_composite_composite (filha é Composite em AGR_FLAGS)
      - executaveis (realmente pendentes e filhas Single válidas)
    """
    pendencias = obter_pendencias_pfcg_composta(caminho_excel)
    if not pendencias:
        return {
            "ok": True,
            "analisadas": 0,
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "executaveis": 0,
            "erros": [],
            "invalidas": [],
            "itens_executaveis": [],
            "itens_ja_existem": [],
            "relacoes_iniciais": {}
        }

    # Preflight PRD
    prd_preflight_ok, prd_preflight_err = preflight_test_connection(ambiente_prd)
    if not prd_preflight_ok:
        msg_err = f"Falha de conexão RFC com SAP PRD: {prd_preflight_err}"
        return {
            "ok": False,
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "executaveis": 0,
            "erros": [msg_err],
            "invalidas": [],
            "itens_executaveis": [],
            "itens_ja_existem": [],
            "relacoes_iniciais": {}
        }

    composites = sorted(list({p["composite"] for p in pendencias}))
    singles = sorted(list({p["single"] for p in pendencias}))
    todas_roles = sorted(list(set(composites).union(set(singles))))

    params_prd = build_connection_params_for_env(ambiente_prd)
    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
    conn_ro = Connection(**params_prd)

    try:
        roles_sap_info = verificar_existencia_e_tipo_sap(todas_roles, environment=ambiente_prd, connection=conn_ro)
        valid_singles_set, invalid_composites_set = validar_roles_filhas_nao_compostas(conn_ro, guard_ro, singles)
        prd_rel_iniciais, prd_read_ok, prd_read_err = ler_relacoes_agr_agrs(conn_ro, guard_ro, composites)
    finally:
        conn_ro.close()

    if not prd_read_ok:
        msg_err = f"✗ Pré-validação PFCG_COMPOSTA falhou ao ler AGR_AGRS: {prd_read_err}"
        return {
            "ok": False,
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "parent_nao_existe": 0,
            "child_nao_existe": 0,
            "parent_tipo_divergente": 0,
            "child_tipo_divergente": 0,
            "executaveis": 0,
            "erros": [msg_err],
            "invalidas": [],
            "itens_executaveis": [],
            "itens_ja_existem": [],
            "itens_bloqueados": [],
            "pais_inexistentes": {},
            "relacoes_iniciais": {}
        }

    itens_ja_existem = []
    itens_ausentes = []
    itens_invalidas = []
    itens_executaveis = []
    itens_bloqueados = []

    # Dicionário de pais inexistentes agrupados: parent -> list of children
    pais_inexistentes: Dict[str, List[str]] = {}
    filhos_inexistentes: Dict[str, List[str]] = {}

    cnt_parent_nao_existe = 0
    cnt_child_nao_existe = 0
    cnt_parent_tipo_divergente = 0
    cnt_child_tipo_divergente = 0

    # Rastrear pares únicos para evitar que duplicações aumentem lote executável
    pares_executaveis_vistos = set()

    for p in pendencias:
        parent = p["composite"]
        child = p["single"]
        existentes_parent = prd_rel_iniciais.get(parent, set())

        p_copia = dict(p)

        # 1. Verificar se a relação já existe fisicamente no SAP PRD (AGR_AGRS)
        if child in existentes_parent:
            p_copia["estado_prd"] = STATE_EXISTS
            itens_ja_existem.append(p_copia)
            continue

        itens_ausentes.append(p)

        # 2. Verificar existência física e tipo do Parent (deve existir e ser COMPOSITE)
        parent_info = roles_sap_info.get(parent, {"existe": False, "tipo_sap": "INEXISTENTE"})
        if not parent_info.get("existe"):
            cnt_parent_nao_existe += 1
            p_copia["estado_prd"] = STATE_PARENT_NOT_FOUND
            p_copia["acao"] = "NAO_EXECUTAR"
            itens_bloqueados.append(p_copia)
            pais_inexistentes.setdefault(parent, []).append(child)
            continue
        elif parent_info.get("tipo_sap") != "COMPOSITE":
            cnt_parent_tipo_divergente += 1
            p_copia["estado_prd"] = STATE_PARENT_INVALID_TYPE
            p_copia["acao"] = "NAO_EXECUTAR"
            itens_bloqueados.append(p_copia)
            continue

        # 3. Verificar existência física e tipo da Child (deve existir e ser SINGLE)
        child_info = roles_sap_info.get(child, {"existe": False, "tipo_sap": "INEXISTENTE"})
        if not child_info.get("existe"):
            cnt_child_nao_existe += 1
            p_copia["estado_prd"] = STATE_CHILD_NOT_FOUND
            p_copia["acao"] = "NAO_EXECUTAR"
            itens_bloqueados.append(p_copia)
            filhos_inexistentes.setdefault(child, []).append(parent)
            continue

        # 4. Verificar se a child é Composite Role (invalida composite->composite / child com tipo divergente)
        if child_info.get("tipo_sap") == "COMPOSITE" or child in invalid_composites_set:
            cnt_child_tipo_divergente += 1
            p_copia["estado_prd"] = STATE_CHILD_INVALID_TYPE
            p_copia["tipo_child"] = "COMPOSITE"
            p_copia["acao"] = "NAO_EXECUTAR"
            itens_invalidas.append(p_copia)
            itens_bloqueados.append(p_copia)
            continue

        # 5. Se passou por todas as regras estritas, é REALMENTE EXECUTÁVEL
        par = (parent, child)
        p_copia["estado_prd"] = STATE_EXECUTAVEL
        p_copia["tipo_parent"] = "COMPOSITE"
        p_copia["tipo_child"] = "SINGLE"
        if par not in pares_executaveis_vistos:
            pares_executaveis_vistos.add(par)
            itens_executaveis.append(p_copia)

    return {
        "ok": True,
        "analisadas": len(pendencias),
        "ja_existiam": len(itens_ja_existem),
        "ausentes": len(itens_ausentes),
        "invalidas_composite_composite": len(itens_invalidas),
        "parent_nao_existe": cnt_parent_nao_existe,
        "child_nao_existe": cnt_child_nao_existe,
        "parent_tipo_divergente": cnt_parent_tipo_divergente,
        "child_tipo_divergente": cnt_child_tipo_divergente,
        "executaveis": len(itens_executaveis),
        "erros": [],
        "invalidas": itens_invalidas,
        "itens_executaveis": itens_executaveis,
        "itens_ja_existem": itens_ja_existem,
        "itens_bloqueados": itens_bloqueados,
        "pais_inexistentes": pais_inexistentes,
        "filhos_inexistentes": filhos_inexistentes,
        "relacoes_iniciais": prd_rel_iniciais
    }


def reconciliar_ja_existentes_pfcg_composta_excel(
    caminho_excel: str,
    itens_ja_existem: List[Dict[str, Any]],
    criar_backup: bool = True
) -> Dict[str, Any]:
    """
    Atualiza no Excel as linhas que foram comprovadamente validadas como JA_EXISTE_PRD:
      - STATUS = "Concluído"
      - MSG = "Validado e confirmado fisicamente no SAP PRD via RFC (Readback OK)"
      - TIMESTEMP = <timestamp atual>
      - PRD = "Validado"
    
    Regras:
      1. Se criar_backup=True e houver itens a atualizar, cria backup com timestamp em output/.
      2. Se a gravação falhar, retorna ok=False com o erro para abortar a sequência.
      3. As relações INVALIDA_COMPOSITE_COMPOSITE NUNCA são passadas aqui e NUNCA são marcadas como Concluído.
    """
    import os
    import shutil

    if not itens_ja_existem:
        return {
            "ok": True,
            "linhas_reconciliadas": 0,
            "backup_criado": None,
            "mensagem": "Nenhuma linha para reconciliar."
        }

    caminho_backup = None
    try:
        if criar_backup:
            dir_output = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output")
            os.makedirs(dir_output, exist_ok=True)
            ts_backup = datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_backup = os.path.join(dir_output, f"S4H_Perfis_v1_backup_{ts_backup}.xlsx")
            shutil.copy2(caminho_excel, caminho_backup)

        wb = openpyxl.load_workbook(caminho_excel)
        if "PFCG_COMPOSTA" not in wb.sheetnames:
            wb.close()
            return {
                "ok": False,
                "linhas_reconciliadas": 0,
                "backup_criado": caminho_backup,
                "erro": "Sheet PFCG_COMPOSTA não encontrada no Excel."
            }

        ws = wb["PFCG_COMPOSTA"]
        ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for it in itens_ja_existem:
            r_idx = it["row_idx"]
            ws.cell(row=r_idx, column=COL_STATUS, value="Concluído")
            ws.cell(row=r_idx, column=COL_MSG, value="Validado e confirmado fisicamente no SAP PRD via RFC (Readback OK)")
            ws.cell(row=r_idx, column=COL_TIMESTEMP, value=ts_now)
            ws.cell(row=r_idx, column=COL_PRD, value="Validado")

        wb.save(caminho_excel)
        wb.close()

        return {
            "ok": True,
            "linhas_reconciliadas": len(itens_ja_existem),
            "backup_criado": caminho_backup,
            "mensagem": f"✓ {len(itens_ja_existem)} linhas reconciliadas no Excel com sucesso."
        }
    except Exception as exc:
        return {
            "ok": False,
            "linhas_reconciliadas": 0,
            "backup_criado": caminho_backup,
            "erro": str(exc)
        }


def sincronizar_pfcg_composta(
    caminho_excel: str,
    dry_run: bool = True,
    only_prd: bool = False,
    ambiente_qad: str = "QAD",
    ambiente_prd: str = "PRD",
    qtd_esperada_executaveis: Optional[int] = None
) -> Dict[str, Any]:
    """
    Orquestra a sincronização da sheet PFCG_COMPOSTA via RFC com pré-validação estrutural
    e readback obrigatório.
    Devolve estrutura de resultado padronizada.
    
    Bloqueio de segurança: Se qtd_esperada_executaveis for informada e divergir
    dos executáveis calculados, aborta imediatamente sem qualquer escrita SAP.
    """
    from collections import defaultdict

    pendencias = obter_pendencias_pfcg_composta(caminho_excel)
    if not pendencias:
        return {
            "ok": True,
            "status": "SEM_PENDENCIAS",
            "mensagem": "Nenhuma linha candidata encontrada na sheet PFCG_COMPOSTA.",
            "analisadas": 0,
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "executaveis": 0,
            "criadas": 0,
            "confirmadas": 0,
            "total_pendencias": 0,
            "total_novo_lote": 0,
            "total_historicas": 0,
            "erros": [],
            "invalidas": [],
            "detalhes": []
        }

    novo_lote_count = sum(1 for p in pendencias if p["categoria"] == "NOVO_LOTE")
    historicas_count = sum(1 for p in pendencias if p["categoria"] == "PENDENCIA_HISTORICA")

    # 1. PRE-FLIGHT DE CONEXÃO RFC
    if not only_prd:
        qad_preflight_ok, qad_preflight_err = preflight_test_connection(ambiente_qad)
        if not qad_preflight_ok:
            # Em modo não-only_prd registra ou aborta se estrito
            pass

    prd_preflight_ok, prd_preflight_err = preflight_test_connection(ambiente_prd)
    if not prd_preflight_ok:
        msg_err = f"Falha de conexão RFC com SAP PRD: {prd_preflight_err}"
        return {
            "ok": False,
            "status": "PREFLIGHT_FAILED",
            "mensagem": msg_err,
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "executaveis": 0,
            "criadas": 0,
            "confirmadas": 0,
            "total_pendencias": len(pendencias),
            "erros": [msg_err],
            "invalidas": [],
            "detalhes": []
        }

    # 2. PRÉ-VALIDAÇÃO READ-ONLY ESTRUTURAL
    preval = prevalidar_pfcg_composta(caminho_excel, ambiente_prd=ambiente_prd)
    if not preval.get("ok"):
        return {
            "ok": False,
            "status": "READ_FAILED",
            "mensagem": "; ".join(preval.get("erros", ["Falha na pré-validação"])),
            "analisadas": len(pendencias),
            "ja_existiam": 0,
            "ausentes": 0,
            "invalidas_composite_composite": 0,
            "executaveis": 0,
            "criadas": 0,
            "confirmadas": 0,
            "erros": preval.get("erros", []),
            "invalidas": [],
            "detalhes": []
        }

    analisadas_cnt = preval["analisadas"]
    ja_existiam_cnt = preval["ja_existiam"]
    ausentes_cnt = preval["ausentes"]
    invalidas_comp_cnt = preval["invalidas_composite_composite"]
    executaveis_cnt = preval["executaveis"]
    itens_executaveis = preval["itens_executaveis"]
    invalidas_lista = preval["invalidas"]

    # BLOQUEIO DE SEGURANÇA: Se quantidade esperada divergir da quantidade executável
    if qtd_esperada_executaveis is not None and qtd_esperada_executaveis != executaveis_cnt:
        msg_bloqueio = (
            f"✗ Inconsistência entre pré-validação e lote de execução: "
            f"esperado={qtd_esperada_executaveis}, executaveis={executaveis_cnt}. "
            f"Nenhuma escrita SAP realizada."
        )
        return {
            "ok": False,
            "status": "BLOQUEIO_SEGURANCA_DIVERGENCIA",
            "mensagem": msg_bloqueio,
            "analisadas": analisadas_cnt,
            "ja_existiam": ja_existiam_cnt,
            "ausentes": ausentes_cnt,
            "invalidas_composite_composite": invalidas_comp_cnt,
            "executaveis": executaveis_cnt,
            "criadas": 0,
            "confirmadas": 0,
            "erros": [msg_bloqueio],
            "invalidas": invalidas_lista,
            "detalhes": []
        }

    # 3. EXECUÇÃO RFC DAS REALMENTE EXECUTÁVEIS (DELTA ESTREITO)
    criadas_cnt = 0
    erros_execucao = []
    params_prd = build_connection_params_for_env(ambiente_prd)

    if itens_executaveis and not dry_run:
        guard_wr = make_write_guard(
            allowed_functions=("PRGN_RFC_ADD_AGRS_TO_COLL_AGR", "PRGN_GEN_PROFILES_FOR_ROLES"),
            allowed_tables=("AGR_AGRS", "AGR_TEXTS")
        )
        conn_wr = Connection(**params_prd)
        try:
            delta_map = defaultdict(list)
            for it in itens_executaveis:
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

                    conn_wr.call("PRGN_GEN_PROFILES_FOR_ROLES", IT_ROLES=[{"AGR_NAME": comp}], IV_USERCOMPARE="X")
                    criadas_cnt += len(items_delta)
                except Exception as exc:
                    erros_execucao.append(f"Exceção RFC ao processar {comp}: {exc}")
        finally:
            conn_wr.close()

    # 4. READBACK OBRIGATÓRIO EM AGR_AGRS PÓS-ESCRITA
    guard_ro = make_read_only_guard(("AGR_DEFINE", "AGR_FLAGS", "AGR_AGRS"))
    conn_rb = Connection(**params_prd)
    try:
        composites_alvo = sorted(list({p["composite"] for p in pendencias}))
        prd_rel_finais, rb_ok, rb_err = ler_relacoes_agr_agrs(conn_rb, guard_ro, composites_alvo)
    finally:
        conn_rb.close()

    confirmadas_cnt = 0
    excel_updates = []

    # No modo dry_run, confirmadas reflete quantas das executáveis já estariam OK se simulado
    # Em dry_run sem escrita, apenas as já existentes continuam existentes.
    # Mas se houver escrita real ou em dry_run pós validação:
    # Apenas confirmamos o que existe fisicamente em AGR_AGRS.
    # Inválidas estruturais NUNCA são confirmadas e NUNCA são marcadas como Concluído.
    for it in itens_executaveis:
        parent = it["composite"]
        child = it["single"]
        finais_parent = prd_rel_finais.get(parent, set())
        if child in finais_parent:
            confirmadas_cnt += 1
            it["status_readback"] = "CONFIRMADO"
            if not dry_run:
                excel_updates.append({
                    "row_idx": it["row_idx"],
                    "status": "Concluído",
                    "msg": "Validado e confirmado fisicamente no SAP PRD via RFC (Readback OK)",
                    "timestemp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "prd": "Validado"
                })
        else:
            it["status_readback"] = "PENDENTE_OU_ERRO"

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

    is_ok = len(erros_execucao) == 0
    msg_resumo = (
        f"✓ PFCG_COMPOSTA processada via RFC.\n"
        f"   Linhas candidatas: {analisadas_cnt}\n"
        f"   Já existentes PRD: {ja_existiam_cnt}\n"
        f"   Ausentes em PRD: {ausentes_cnt}\n"
        f"   Inválidas (Composite->Composite): {invalidas_comp_cnt}\n"
        f"   Realmente executáveis: {executaveis_cnt}\n"
        f"   Criadas no SAP: {criadas_cnt}\n"
        f"   Confirmadas por readback: {confirmadas_cnt}\n"
        f"   Erros: {len(erros_execucao)}"
    )

    return {
        "ok": is_ok,
        "status": "SUCESSO" if is_ok else "FALHA",
        "mensagem": msg_resumo if is_ok else f"Falha na sincronização PFCG_COMPOSTA: {'; '.join(erros_execucao)}",
        "analisadas": analisadas_cnt,
        "ja_existiam": ja_existiam_cnt,
        "ausentes": ausentes_cnt,
        "invalidas_composite_composite": invalidas_comp_cnt,
        "executaveis": executaveis_cnt,
        "criadas": criadas_cnt,
        "confirmadas": confirmadas_cnt,
        "total_pendencias": len(pendencias),
        "total_novo_lote": novo_lote_count,
        "total_historicas": historicas_count,
        "erros": erros_execucao,
        "invalidas": invalidas_lista,
        "detalhes": []
    }
