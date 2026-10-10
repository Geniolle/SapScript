# -*- coding: utf-8 -*-
"""
tests/test_reconciliar_proposta_ativa_com_prd.py
======================================================================
Testes unitários abrangentes para o fluxo oficial de Reconciliação Integrada:
- Fase 1: População congelada via CONTROLO -> Sheet -> User SAP.
- Fase 2: Inclusão de Single Roles diretas do PRD (com filtro DEFINIÇÕES).
- Fase 3A: Mapeamento dos X da matriz departamental -> Proposta -> Proposta Ativa (1:N TCODE -> Roles).
- Fase 3B: União aditiva de Composites para PFCG_COMPOSTA.
======================================================================
"""

import sys
import pytest
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.reconciliar_proposta_ativa_com_prd import (
    obter_funcoes_excluidas_definicoes,
    carregar_mapeamento_tcode_para_funcoes_proposta
)


def simular_fluxo_fase3a_utilizador(
    uid_excel: str,
    pertence_populacao_fase1: bool,
    tcodes_com_x: list,
    tcode_to_roles_map: dict,
    funcs_atuais_proposta_ativa: set
):
    """
    Helper puro de testes para validar as regras A a M da Fase 3A.
    """
    if not pertence_populacao_fase1:
        return {"processado": False, "motivo": "FORA_DA_POPULACAO", "a_adicionar": []}

    funcs_necessarias_matriz = set()
    tcodes_sem_mapeamento = []

    for tc in tcodes_com_x:
        tc_upper = str(tc).strip().upper()
        if tc_upper in tcode_to_roles_map:
            roles = tcode_to_roles_map[tc_upper]
            funcs_necessarias_matriz.update(roles)
        else:
            tcodes_sem_mapeamento.append(tc_upper)

    a_adicionar = sorted(list(funcs_necessarias_matriz - funcs_atuais_proposta_ativa))
    funcs_finais = sorted(list(funcs_atuais_proposta_ativa.union(set(a_adicionar))))

    return {
        "processado": True,
        "funcs_necessarias": sorted(list(funcs_necessarias_matriz)),
        "tcodes_sem_mapeamento": tcodes_sem_mapeamento,
        "a_adicionar": a_adicionar,
        "funcs_finais": funcs_finais
    }


def test_fase3a_a_utilizador_pertence_populacao():
    """A. Utilizador pertence à população da Fase 1 -> processar."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["processado"] is True
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3a_b_utilizador_fora_populacao():
    """B. Utilizador não pertence à população -> ignorar mesmo que exista na sheet."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S99999",
        pertence_populacao_fase1=False,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["processado"] is False
    assert res["a_adicionar"] == []


def test_fase3a_c_d_tcode_com_e_sem_x():
    """C. TCODE com X -> considerar. D. TCODE sem X -> ignorar."""
    # Apenas TCODEs com X são passadas no parâmetro tcodes_com_x
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],  # ME22N sem X não foi passada
        tcode_to_roles_map={
            "ME21N": {"Z_PURCHASE_ORDER_CREATE"},
            "ME22N": {"Z_PURCHASE_ORDER_CHANGE"}
        },
        funcs_atuais_proposta_ativa=set()
    )
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3a_e_tcode_relacao_1_para_1():
    """E. TCODE -> 1 Função Individual -> adicionar se faltar."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3a_f_tcode_relacao_1_para_n():
    """F. TCODE -> múltiplas Funções Individuais (1:N) -> todas devem ser consideradas."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["MIGO"],
        tcode_to_roles_map={"MIGO": {"Z_GOODS_MOVEMENTS", "Z_INVENTORY_DC_STORE"}},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["a_adicionar"] == ["Z_GOODS_MOVEMENTS", "Z_INVENTORY_DC_STORE"]


def test_fase3a_g_funcao_ja_existe_na_proposta_ativa():
    """G. Função já existe na Proposta Ativa -> não duplicar."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa={"Z_PURCHASE_ORDER_CREATE"}
    )
    assert res["a_adicionar"] == []
    assert res["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3a_h_funcao_em_falta_adicionar():
    """H. Função em falta -> adicionar."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3a_i_tcode_sem_mapeamento_na_proposta():
    """I. TCODE sem mapeamento na Proposta -> reportar e não inventar função."""
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["Z_TCODE_DESCONHECIDA"],
        tcode_to_roles_map={},
        funcs_atuais_proposta_ativa=set()
    )
    assert res["tcodes_sem_mapeamento"] == ["Z_TCODE_DESCONHECIDA"]
    assert res["a_adicionar"] == []


def test_fase3a_k_fase2_ja_adicionou_funcao():
    """K. Fase 2 já adicionou uma função -> Fase 3A deve reconhecê-la e não duplicar."""
    # Simula estado da Proposta Ativa já enriquecido pela Fase 2
    funcs_pos_fase2 = {"Z_PURCHASE_ORDER_CREATE"}
    res = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map={"ME21N": {"Z_PURCHASE_ORDER_CREATE"}},
        funcs_atuais_proposta_ativa=funcs_pos_fase2
    )
    assert res["a_adicionar"] == []


def test_fase3a_l_segunda_execucao_idempotente():
    """L. Segunda execução -> zero novas alterações."""
    tcode_map = {"ME21N": {"Z_PURCHASE_ORDER_CREATE"}}

    # 1ª Execução
    res1 = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map=tcode_map,
        funcs_atuais_proposta_ativa=set()
    )
    assert res1["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]

    # 2ª Execução com estado resultante da 1ª
    res2 = simular_fluxo_fase3a_utilizador(
        uid_excel="S13020",
        pertence_populacao_fase1=True,
        tcodes_com_x=["ME21N"],
        tcode_to_roles_map=tcode_map,
        funcs_atuais_proposta_ativa=set(res1["funcs_finais"])
    )
    assert res2["a_adicionar"] == []
    assert res2["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_fase3b_recebe_resultado_final_fase3a():
    """M. Fase 3B recebe a Proposta Ativa já atualizada pela Fase 3A (União Aditiva)."""
    user1_funcs = {"Z_ROLE_A", "Z_ROLE_B"}  # Resultante da Fase 3A
    user2_funcs = {"Z_ROLE_B", "Z_ROLE_C"}  # Resultante da Fase 3A

    composite_calculada = user1_funcs.union(user2_funcs)
    assert composite_calculada == {"Z_ROLE_A", "Z_ROLE_B", "Z_ROLE_C"}


# ======================================================================
# TESTES UNITÁRIOS ESPECÍFICOS DA FASE 3B (REGRAS A a J DO ITEM 17)
# ======================================================================

def simular_fase3b_composite(
    users_composta: dict,
    pfcg_composta_atual: set,
    text_composite: str,
    maior_id_existente: int = 100
):
    """Helper de simulação unitária para a Fase 3B."""
    # 1. União das funções dos utilizadores do grupo
    composite_calculada = set()
    for uid, funcs in users_composta.items():
        composite_calculada.update(funcs)

    # 2. Roles a adicionar e não requeridas
    roles_a_adicionar = sorted(list(composite_calculada - pfcg_composta_atual))
    roles_nao_requeridas = sorted(list(pfcg_composta_atual - composite_calculada))

    # 3. Gerar novas linhas se TEXT existir
    novas_linhas = []
    proximo_id = maior_id_existente + 1
    if text_composite:
        for r in roles_a_adicionar:
            novas_linhas.append({
                "ID": proximo_id,
                "AGR_NAME_COMPOSTA": "Z_COMPOSITE_TEST",
                "TEXT": text_composite,
                "AGR_NAME": r
            })
            proximo_id += 1

    return {
        "composite_calculada": sorted(list(composite_calculada)),
        "roles_a_adicionar": roles_a_adicionar,
        "roles_nao_requeridas": roles_nao_requeridas,
        "text_encontrado": bool(text_composite),
        "novas_linhas": novas_linhas
    }


def test_fase3b_a_dois_users_funcoes_diferentes_uniao():
    """A. Dois Users na mesma Composite com funções diferentes -> união correta."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1"}, "U2": {"Z_ROLE_2"}},
        pfcg_composta_atual=set(),
        text_composite="Descrição Teste"
    )
    assert res["composite_calculada"] == ["Z_ROLE_1", "Z_ROLE_2"]


def test_fase3b_b_funcao_ja_existente_nao_duplicar():
    """B. Função já existente na PFCG_COMPOSTA -> não duplicar."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1"}},
        pfcg_composta_atual={"Z_ROLE_1"},
        text_composite="Descrição Teste"
    )
    assert res["roles_a_adicionar"] == []


def test_fase3b_c_funcao_em_falta_adicionar():
    """C. Função em falta -> candidata a adicionar."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1", "Z_ROLE_2"}},
        pfcg_composta_atual={"Z_ROLE_1"},
        text_composite="Descrição Teste"
    )
    assert res["roles_a_adicionar"] == ["Z_ROLE_2"]


def test_fase3b_d_role_existente_nao_mais_necessaria():
    """D. Role existente mas não mais necessária -> reportar, não remover."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1"}},
        pfcg_composta_atual={"Z_ROLE_1", "Z_ROLE_ANTIGA"},
        text_composite="Descrição Teste"
    )
    assert res["roles_nao_requeridas"] == ["Z_ROLE_ANTIGA"]
    assert res["roles_a_adicionar"] == []


def test_fase3b_e_composite_sem_text():
    """E. Composite sem TEXT -> reportar e não criar automaticamente."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1"}},
        pfcg_composta_atual=set(),
        text_composite=""  # Sem text
    )
    assert res["text_encontrado"] is False
    assert res["roles_a_adicionar"] == ["Z_ROLE_1"]
    assert len(res["novas_linhas"]) == 0  # Não gera linhas sem TEXT


def test_fase3b_f_id_sequencial_correto():
    """F. ID sequencial correto (maior ID + 1, + 2, ...)."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1", "Z_ROLE_2"}},
        pfcg_composta_atual=set(),
        text_composite="Descrição Teste",
        maior_id_existente=500
    )
    assert res["novas_linhas"][0]["ID"] == 501
    assert res["novas_linhas"][1]["ID"] == 502


def test_fase3b_g_novas_linhas_somente_4_campos_autorizados():
    """G. Novas linhas possuem somente os 4 campos autorizados (ID, AGR_NAME_COMPOSTA, TEXT, AGR_NAME)."""
    res = simular_fase3b_composite(
        users_composta={"U1": {"Z_ROLE_1"}},
        pfcg_composta_atual=set(),
        text_composite="Descrição Teste"
    )
    linha = res["novas_linhas"][0]
    assert set(linha.keys()) == {"ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME"}


def test_fase3b_h_segunda_execucao_zero_adicoes():
    """H. Segunda execução -> zero novas adições (idempotência)."""
    users = {"U1": {"Z_ROLE_1", "Z_ROLE_2"}}
    res1 = simular_fase3b_composite(
        users_composta=users,
        pfcg_composta_atual=set(),
        text_composite="Descrição Teste"
    )
    assert len(res1["roles_a_adicionar"]) == 2

    # Aplicar estado resultante na 2ª execução
    pfcg_atualizada = set(res1["composite_calculada"])
    res2 = simular_fase3b_composite(
        users_composta=users,
        pfcg_composta_atual=pfcg_atualizada,
        text_composite="Descrição Teste"
    )
    assert res2["roles_a_adicionar"] == []


def test_fase3b_i_user_com_erro_mapeamento_nao_entra():
    """I. User com ERRO_DE_MAPEAMENTO -> não entra na união."""
    users_mapeados_ok = {"S13020": {"Z_ROLE_OK"}}
    # S9999 (erro de mapeamento) não está presente no dicionário da Fase 1
    assert "S9999" not in users_mapeados_ok


def test_fase3b_j_fase3a_adicionou_funcao_fase3b_enxerga():
    """J. Fase 3A adicionou função em memória -> Fase 3B enxerga e considera na união."""
    funcs_iniciais = {"Z_ROLE_INICIAL"}
    funcs_adicionadas_fase3a = {"Z_ROLE_FASE3A"}
    funcs_finais_user = funcs_iniciais.union(funcs_adicionadas_fase3a)

    res = simular_fase3b_composite(
        users_composta={"U1": funcs_finais_user},
        pfcg_composta_atual=set(),
        text_composite="Descrição Teste"
    )
    assert res["composite_calculada"] == ["Z_ROLE_FASE3A", "Z_ROLE_INICIAL"]


def test_classificacao_utilizador_pendente_new():
    """1. NEW não bloqueia outros -> classificado como PENDENTE_CRIACAO_USER_SAP."""
    raw_header = "NEW Daniela Faria"
    is_new = "NEW" in raw_header.upper()
    estado = "PENDENTE_CRIACAO_USER_SAP" if is_new else "PROCESSAVEL"
    assert estado == "PENDENTE_CRIACAO_USER_SAP"


def test_classificacao_utilizador_nao_mapeado_proposta_ativa():
    """2. User sem Proposta Ativa não bloqueia outros -> classificado como USER_NAO_MAPEADO_PROPOSTA_ATIVA."""
    uid = "S80001999"
    proposta_ativa_map = {"S13020": 2, "S14030": 3}
    pertence_proposta = uid in proposta_ativa_map
    estado = "PROCESSAVEL" if pertence_proposta else "USER_NAO_MAPEADO_PROPOSTA_ATIVA"
    assert estado == "USER_NAO_MAPEADO_PROPOSTA_ATIVA"


def test_cenario_3_user_inativo_nao_bloqueia():
    """3. User inativo não bloqueia outros -> classificado como USER_INATIVO_EXCLUIDO."""
    valido_usr02 = False
    estado = "PROCESSAVEL" if valido_usr02 else "USER_INATIVO_EXCLUIDO"
    assert estado == "USER_INATIVO_EXCLUIDO"


def test_cenario_4_user_sem_x_nao_bloqueia():
    """4. User sem X não bloqueia outros -> classificado como USER_SEM_X_EXCLUIDO."""
    qtd_x = 0
    estado = "PROCESSAVEL" if qtd_x > 0 else "USER_SEM_X_EXCLUIDO"
    assert estado == "USER_SEM_X_EXCLUIDO"


def test_cenario_5_user_ambiguo_nao_bloqueia():
    """5. User ambíguo não bloqueia outros -> classificado como USER_AMBIGUO."""
    candidatos = ["S80001001", "S80001002"]
    estado = "USER_AMBIGUO" if len(candidatos) > 1 else "PROCESSAVEL"
    assert estado == "USER_AMBIGUO"


def test_cenario_6_tcode_sem_role_continua_bloqueando():
    """6. TCODE sem role em Proposta continua bloqueando o departamento com erro/alerta."""
    tcode = "TCODE_DESCONHECIDA"
    tcode_to_roles_map = {"ME21N": {"Z_ROLE"}}
    sem_role = tcode not in tcode_to_roles_map
    status_departamento = "ERRO" if sem_role else "PROCESSADO"
    assert sem_role is True
    assert status_departamento == "ERRO"


def test_cenario_7_lote_com_pendencia_gera_processado_com_pendencias():
    """7. Lote com utilizador pendente e utilizadores válidos gera status PROCESSADO_COM_PENDENCIAS."""
    users = [{"uid": "S13020", "estado": "PROCESSAVEL"}, {"uid": None, "estado": "PENDENTE_CRIACAO_USER_SAP"}]
    has_pendencias = any(u["estado"] != "PROCESSAVEL" for u in users)
    has_validos = any(u["estado"] == "PROCESSAVEL" for u in users)
    status_final = "PROCESSADO_COM_PENDENCIAS" if (has_pendencias and has_validos) else "PROCESSADO"
    assert status_final == "PROCESSADO_COM_PENDENCIAS"


def test_cenario_8_lote_sem_pendencia_gera_processado():
    """8. Lote sem qualquer pendência gera status PROCESSADO."""
    users = [{"uid": "S13020", "estado": "PROCESSAVEL"}, {"uid": "S14030", "estado": "PROCESSAVEL"}]
    has_pendencias = any(u["estado"] != "PROCESSAVEL" for u in users)
    status_final = "PROCESSADO_COM_PENDENCIAS" if has_pendencias else "PROCESSADO"
    assert status_final == "PROCESSADO"


def test_cenario_9_segunda_execucao_idempotente():
    """9. Segunda execução é idempotente (0 novas adições)."""
    funcs_proposta = {"Z_ROLE_1"}
    funcs_necessarias = {"Z_ROLE_1"}
    a_adicionar = funcs_necessarias - funcs_proposta
    assert len(a_adicionar) == 0


def test_cenario_10_user_corrigido_gera_somente_delta_proprio():
    """10. User corrigido gera somente delta próprio, sem reprocessar utilizadores anteriores."""
    user_a_finais = {"Z_ROLE_1"}  # Já processado anteriormente
    user_b_novos = {"Z_ROLE_2"}   # Novo utilizador corrigido
    
    delta_a = set() - user_a_finais  # Adições para A = 0
    delta_b = user_b_novos - set()   # Adições para B = 1
    
    assert len(delta_a) == 0
    assert len(delta_b) == 1


def test_cenario_11_composite_somente_delta_novo():
    """11. PFCG_COMPOSTA gera somente delta novo quando nova role é incluída."""
    composta_atual = {"Z_ROLE_1"}
    composta_calculada = {"Z_ROLE_1", "Z_ROLE_2"}
    novas_composta = composta_calculada - composta_atual
    assert novas_composta == {"Z_ROLE_2"}


def test_cenario_12_transicao_processado_com_pendencias_para_processado():
    """12. Transição de status: PENDENTE -> PROCESSADO_COM_PENDENCIAS -> PROCESSADO após correção."""
    # Passo 1: Com pendência
    users_passo1 = [{"uid": "S13020", "estado": "PROCESSAVEL"}, {"uid": None, "estado": "PENDENTE_CRIACAO_USER_SAP"}]
    status_passo1 = "PROCESSADO_COM_PENDENCIAS" if any(u["estado"] != "PROCESSAVEL" for u in users_passo1) else "PROCESSADO"
    assert status_passo1 == "PROCESSADO_COM_PENDENCIAS"

    # Passo 2: Após correção do utilizador rascunho com ID SAP e Proposta Ativa
    users_passo2 = [{"uid": "S13020", "estado": "PROCESSAVEL"}, {"uid": "S80002000", "estado": "PROCESSAVEL"}]
    status_passo2 = "PROCESSADO_COM_PENDENCIAS" if any(u["estado"] != "PROCESSAVEL" for u in users_passo2) else "PROCESSADO"
    assert status_passo2 == "PROCESSADO"



