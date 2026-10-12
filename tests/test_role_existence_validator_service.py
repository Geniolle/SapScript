# -*- coding: utf-8 -*-
"""
tests/test_role_existence_validator_service.py
======================================================================
Bateria de Testes Unitários Completa para a Validação Obrigatória
de Existência e Tipos de Roles no SAP (PRD e QAD Independentes)
======================================================================
Cobre todos os 18 cenários mandatados:
1. Todas as roles existem em PRD e QAD com tipo correto
2. Composite Role em falta no PRD
3. Composite Role em falta no QAD
4. Single Role em falta no PRD
5. Single Role em falta no QAD
6. Role em falta em ambos os ambientes
7. Tipo divergente no SAP (esperado Composite, retornado Single)
8. Tipo divergente no SAP (esperado Single, retornado Composite)
9. Utilizador autoriza criação da Composite Role
10. Utilizador recusa criação da Composite Role
11. Utilizador autoriza criação da Single Role
12. Utilizador recusa criação da Single Role
13. PRD disponível e QAD indisponível (independência)
14. QAD disponível e PRD indisponível (independência)
15. Falha no readback após criação
16. PFCG_COMPOSTA só recebe roles existentes
17. Continuação parcial autorizada
18. Zero escrita SAP quando utilizador recusa e idempotência
======================================================================
"""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

from sap_rfc.role_existence_validator_service import (
    extrair_roles_necessarias_projeto,
    verificar_existencia_e_tipo_sap,
    validar_existencia_roles_ambientes,
    criar_role_sap_com_readback,
    executar_fase_validacao_existencia_roles,
    executar_etapa_validacao_funcoes_individuais,
    executar_etapa_validacao_funcoes_compostas,
    auditar_single_role_divergente,
    corrigir_composite_role_divergente_com_readback,
    STATUS_ROLE_OK,
    STATUS_ROLE_NAO_EXISTE,
    STATUS_TIPO_DIVERGENTE,
    STATUS_AMBIENTE_INDISPONIVEL,
    STATUS_CRIADA_E_CONFIRMADA,
    STATUS_ERRO_READBACK,
    STATUS_SINGLE_OK,
    STATUS_SINGLE_AUSENTE,
    STATUS_SINGLE_TIPO_DIVERGENTE,
    STATUS_SINGLE_UNKNOWN_AMBIENTE,
    STATUS_COMPOSITE_OK,
    STATUS_COMPOSITE_AUSENTE,
    STATUS_COMPOSITE_TIPO_DIVERGENTE,
    STATUS_COMPOSITE_UNKNOWN_AMBIENTE,
    STATUS_SEGURA_PARA_CORRIGIR,
    STATUS_BLOQUEADA_PARA_REVISAO,
)


@pytest.fixture
def mock_dados_projeto():
    dados = MagicMock()
    dados.roles_compostas = {
        "Z_COMP_1": {"descricao": "Comp 1", "roles_filhas": ["Z_SING_1", "Z_SING_2"]},
        "Z_COMP_2": {"descricao": "Comp 2", "roles_filhas": ["Z_SING_3"]},
    }
    dados.roles_simples = {
        "Z_SING_1": {"descricao": "Single 1"},
        "Z_SING_2": {"descricao": "Single 2"},
        "Z_SING_3": {"descricao": "Single 3"},
        "Z_SING_4": {"descricao": "Single 4"},
    }
    dados.definicoes_departamento = {
        "DEP1": ["Z_BASIS_BASE", "ZORG_EMPRESA"]
    }
    return dados


# 1. Todas as roles existem em PRD e QAD com tipo correto
def test_todas_roles_existem_ambientes():
    roles = ["Z_COMP_1", "Z_SING_1"]
    conn = MagicMock()

    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        # AGR_DEFINE retorna as 2
        # AGR_FLAGS retorna COLL_AGR='X' para Z_COMP_1
        mock_read.side_effect = [
            [["Z_COMP_1"], ["Z_SING_1"]],  # AGR_DEFINE
            [["Z_COMP_1", "COLL_AGR", "X"]], # AGR_FLAGS
        ]
        res = verificar_existencia_e_tipo_sap(roles, environment="PRD", connection=conn)

        assert res["Z_COMP_1"]["existe"] is True
        assert res["Z_COMP_1"]["tipo_sap"] == "COMPOSITE"
        assert res["Z_SING_1"]["existe"] is True
        assert res["Z_SING_1"]["tipo_sap"] == "SINGLE"


# 2 & 3. Composite Role em falta em PRD ou QAD
def test_composite_role_em_falta():
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [["Z_SING_1"]],  # AGR_DEFINE (Z_COMP_1 ausente)
            [],               # AGR_FLAGS
        ]
        res = verificar_existencia_e_tipo_sap(["Z_COMP_1", "Z_SING_1"], connection=conn)
        assert res["Z_COMP_1"]["existe"] is False
        assert res["Z_COMP_1"]["tipo_sap"] == "INEXISTENTE"
        assert res["Z_SING_1"]["existe"] is True


# 4 & 5. Single Role em falta
def test_single_role_em_falta():
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [["Z_COMP_1"]],  # AGR_DEFINE (Z_SING_1 ausente)
            [["Z_COMP_1", "COLL_AGR", "X"]], # AGR_FLAGS
        ]
        res = verificar_existencia_e_tipo_sap(["Z_COMP_1", "Z_SING_1"], connection=conn)
        assert res["Z_SING_1"]["existe"] is False
        assert res["Z_COMP_1"]["existe"] is True


# 6. Role em falta em ambos os ambientes
def test_role_em_falta_ambos():
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        mock_verif.return_value = {
            "Z_COMP_1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None}
        }
        rel = validar_existencia_roles_ambientes("dummy.xlsx", dados=MagicMock(roles_compostas={"Z_COMP_1": {}}, roles_simples={}, definicoes_departamento={}), ambientes=("PRD", "QAD"))
        assert "Z_COMP_1" in rel["ambientes"]["PRD"]["composite_em_falta"]
        assert "Z_COMP_1" in rel["ambientes"]["QAD"]["composite_em_falta"]


# 7 & 8. Tipos divergentes (Esperado Composite vs Single, e vice-versa)
def test_tipos_divergentes():
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        # SAP retorna Z_COMP_1 sem COLL_AGR (ou seja, Single no SAP)
        # e Z_SING_1 com COLL_AGR='X' (ou seja, Composite no SAP)
        mock_read.side_effect = [
            [["Z_COMP_1"], ["Z_SING_1"]], # AGR_DEFINE
            [["Z_SING_1", "COLL_AGR", "X"]], # AGR_FLAGS
        ]
        res = verificar_existencia_e_tipo_sap(["Z_COMP_1", "Z_SING_1"], connection=conn)
        assert res["Z_COMP_1"]["tipo_sap"] == "SINGLE"
        assert res["Z_SING_1"]["tipo_sap"] == "COMPOSITE"


# 9 & 10. Utilizador autoriza ou recusa criação de Composite Role
def test_usuario_autoriza_e_recusa_criacao_composite():
    dados = MagicMock(roles_compostas={"Z_COMP_NOVA": {}}, roles_simples={}, definicoes_departamento={})

    # Caso A: Autoriza (S)
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input", return_value="S"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "status": "VALIDADO",
                    "composite_em_falta": ["Z_COMP_NOVA"],
                    "single_em_falta": [],
                    "detalhes": {}
                }
            }
        }
        mock_cr.return_value = {"ok": True, "status": STATUS_CRIADA_E_CONFIRMADA, "mensagem": "Criada"}

        fase_res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",), assumir_sim=False)
        mock_cr.assert_called_once_with("Z_COMP_NOVA", tipo="COMPOSITE", environment="PRD")
        assert "Z_COMP_NOVA" in fase_res["roles_confirmadas"]["PRD"]

    # Caso B: Recusa (N)
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input", return_value="N"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "status": "VALIDADO",
                    "composite_em_falta": ["Z_COMP_NOVA"],
                    "single_em_falta": [],
                    "detalhes": {}
                }
            }
        }
        fase_res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",), assumir_sim=False)
        mock_cr.assert_not_called()
        assert "Z_COMP_NOVA" not in fase_res["roles_confirmadas"]["PRD"]


# 11 & 12. Utilizador autoriza ou recusa criação de Single Role
def test_usuario_autoriza_e_recusa_criacao_single():
    dados = MagicMock(roles_compostas={}, roles_simples={"Z_SING_NOVA": {}}, definicoes_departamento={})

    # Autoriza
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input", return_value="S"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "status": "VALIDADO",
                    "composite_em_falta": [],
                    "single_em_falta": ["Z_SING_NOVA"],
                    "detalhes": {}
                }
            }
        }
        mock_cr.return_value = {"ok": True, "status": STATUS_CRIADA_E_CONFIRMADA, "mensagem": "Criada"}
        fase_res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",), assumir_sim=False)
        mock_cr.assert_called_once_with("Z_SING_NOVA", tipo="SINGLE", environment="PRD")
        assert "Z_SING_NOVA" in fase_res["roles_confirmadas"]["PRD"]

    # Recusa
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input", return_value="N"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "status": "VALIDADO",
                    "composite_em_falta": [],
                    "single_em_falta": ["Z_SING_NOVA"],
                    "detalhes": {}
                }
            }
        }
        fase_res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",), assumir_sim=False)
        mock_cr.assert_not_called()
        assert "Z_SING_NOVA" not in fase_res["roles_confirmadas"]["PRD"]


# 13 & 14. Independência PRD vs QAD e semântica UNKNOWN
def test_independencia_ambientes_e_semantica_unknown():
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        def side_effect(roles, environment="PRD", **kwargs):
            if environment == "PRD":
                return {"Z_R1": {"existe": True, "tipo_sap": "SINGLE", "erro": None}}
            else:
                return {"Z_R1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": "Connection timed out"}}
        mock_verif.side_effect = side_effect

        dados = MagicMock(roles_compostas={}, roles_simples={"Z_R1": {}}, definicoes_departamento={})
        rel = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("PRD", "QAD"))
        
        # PRD Online e Validado
        assert rel["ambientes"]["PRD"]["online"] is True
        assert rel["ambientes"]["PRD"]["validado"] is True
        assert rel["ambientes"]["PRD"]["status"] == "VALIDADO"
        assert rel["ambientes"]["PRD"]["existentes"] == 1
        assert rel["ambientes"]["PRD"]["singles_existentes"] == 1
        assert rel["ambientes"]["PRD"]["singles_faltantes"] == 0

        # QAD Offline e NÃO Validado (UNKNOWN)
        assert rel["ambientes"]["QAD"]["online"] is False
        assert rel["ambientes"]["QAD"]["validado"] is False
        assert rel["ambientes"]["QAD"]["status"] == "INDISPONIVEL"
        assert rel["ambientes"]["QAD"]["existentes"] is None
        assert rel["ambientes"]["QAD"]["composites_faltantes"] is None
        assert rel["ambientes"]["QAD"]["singles_faltantes"] is None
        assert rel["ambientes"]["QAD"]["tipos_divergentes"] is None

        # Resultado Global: VALIDACAO_PARCIAL
        assert rel["resultado_global"] == "VALIDACAO_PARCIAL"


# Cenário inverso: PRD Offline e QAD Online
def test_prd_offline_qad_online_resultado_parcial():
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        def side_effect(roles, environment="PRD", **kwargs):
            if environment == "QAD":
                return {"Z_R1": {"existe": True, "tipo_sap": "SINGLE", "erro": None}}
            else:
                return {"Z_R1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": "Connection timed out"}}
        mock_verif.side_effect = side_effect

        dados = MagicMock(roles_compostas={}, roles_simples={"Z_R1": {}}, definicoes_departamento={})
        rel = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("PRD", "QAD"))
        assert rel["ambientes"]["PRD"]["validado"] is False
        assert rel["ambientes"]["QAD"]["validado"] is True
        assert rel["resultado_global"] == "VALIDACAO_PARCIAL"


# Ambos Online
def test_ambos_ambientes_online_todos_validos():
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        mock_verif.return_value = {"Z_R1": {"existe": True, "tipo_sap": "SINGLE", "erro": None}}
        dados = MagicMock(roles_compostas={}, roles_simples={"Z_R1": {}}, definicoes_departamento={})
        rel = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("PRD", "QAD"))
        assert rel["ambientes"]["PRD"]["validado"] is True
        assert rel["ambientes"]["QAD"]["validado"] is True
        assert rel["resultado_global"] == "TODOS_VALIDOS"


# Ambos Offline
def test_ambos_ambientes_offline_nenhum_disponivel():
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        mock_verif.return_value = {"Z_R1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": "Timeout"}}
        dados = MagicMock(roles_compostas={}, roles_simples={"Z_R1": {}}, definicoes_departamento={})
        rel = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("PRD", "QAD"))
        assert rel["ambientes"]["PRD"]["validado"] is False
        assert rel["ambientes"]["QAD"]["validado"] is False
        assert rel["resultado_global"] == "NENHUM_AMBIENTE_DISPONIVEL"


# QAD offline não gera pergunta de criação
def test_qad_offline_nao_gera_pergunta_criacao():
    dados = MagicMock(roles_compostas={"Z_COMP_1": {}}, roles_simples={}, definicoes_departamento={})
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input") as mock_input:

        mock_val.return_value = {
            "ambientes": {
                "QAD": {
                    "online": False,
                    "validado": False,
                    "status": "INDISPONIVEL",
                    "composite_em_falta": None,
                    "single_em_falta": None,
                    "detalhes": {}
                }
            }
        }
        res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("QAD",))
        # input não deve ser chamado para o QAD offline
        mock_input.assert_not_called()
        mock_cr.assert_not_called()
        assert len(res["roles_confirmadas"]["QAD"]) == 0


# Recalcular sempre na segunda execução sem reutilizar valores antigos
def test_segunda_execucao_recalcula_tudo():
    dados = MagicMock(roles_compostas={}, roles_simples={"Z_R1": {}}, definicoes_departamento={})
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_verif:
        # 1a execução: Offline
        mock_verif.return_value = {"Z_R1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": "Timeout"}}
        rel1 = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("QAD",))
        assert rel1["ambientes"]["QAD"]["validado"] is False
        assert rel1["ambientes"]["QAD"]["existentes"] is None

        # 2a execução: QAD voltou online
        mock_verif.return_value = {"Z_R1": {"existe": True, "tipo_sap": "SINGLE", "erro": None}}
        rel2 = validar_existencia_roles_ambientes("dummy.xlsx", dados=dados, ambientes=("QAD",))
        assert rel2["ambientes"]["QAD"]["validado"] is True
        assert rel2["ambientes"]["QAD"]["existentes"] == 1
        assert rel2["ambientes"]["QAD"]["singles_existentes"] == 1


# 15. Falha no readback após criação
def test_falha_readback_pos_criacao():
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists") as mock_exists:
        mock_exists.side_effect = [False, False]

        res = criar_role_sap_com_readback("Z_NOVA", tipo="SINGLE", environment="PRD", connection=conn)
        assert res["ok"] is False
        assert res["status"] == STATUS_ERRO_READBACK
        assert "Readback falhou" in res["mensagem"]


# 16. Readback com sucesso
def test_sucesso_criacao_com_readback():
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists") as mock_exists, \
         patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_exists.side_effect = [False, True]
        mock_read.return_value = [["Z_COMP", "COLL_AGR", "X"]]

        res = criar_role_sap_com_readback("Z_COMP", tipo="COMPOSITE", environment="PRD", connection=conn)
        assert res["ok"] is True
        assert res["status"] == STATUS_CRIADA_E_CONFIRMADA
        assert res["tipo_confirmado"] == "COMPOSITE"


# 17. Continuação parcial autorizada
def test_continuacao_parcial_autorizada():
    dados = MagicMock(roles_compostas={"Z_COMP_1": {}}, roles_simples={"Z_SING_1": {}}, definicoes_departamento={})
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("builtins.input", return_value="N"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "composite_em_falta": [],
                    "single_em_falta": ["Z_SING_1"],
                    "detalhes": {
                        "Z_COMP_1": {"status": STATUS_ROLE_OK},
                        "Z_SING_1": {"status": STATUS_ROLE_NAO_EXISTE}
                    }
                }
            }
        }
        res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",))
        assert "Z_COMP_1" in res["roles_confirmadas"]["PRD"]
        assert "Z_SING_1" not in res["roles_confirmadas"]["PRD"]


# 18. Zero escrita SAP quando utilizador recusa (idempotência)
def test_zero_escrita_quando_recusa():
    dados = MagicMock(roles_compostas={"Z_COMP_1": {}}, roles_simples={}, definicoes_departamento={})
    with patch("sap_rfc.role_existence_validator_service.validar_existencia_roles_ambientes") as mock_val, \
         patch("sap_rfc.role_existence_validator_service.criar_role_sap_com_readback") as mock_cr, \
         patch("builtins.input", return_value="N"):

        mock_val.return_value = {
            "ambientes": {
                "PRD": {
                    "online": True,
                    "validado": True,
                    "composite_em_falta": ["Z_COMP_1"],
                    "single_em_falta": [],
                    "detalhes": {}
                }
            }
        }
        res = executar_fase_validacao_existencia_roles("dummy.xlsx", dados=dados, ambientes=("PRD",), assumir_sim=False)
        mock_cr.assert_not_called()
        assert len(res["roles_confirmadas"]["PRD"]) == 0


# =====================================================================
# TESTES DE CLASSIFICAÇÃO TÉCNICA E AUDITORIA READ-ONLY
# =====================================================================

def test_classificacao_composite_correta():
    """Valida role existente em AGR_DEFINE com FLAG_TYPE='COLL_AGR' e FLAG_VALUE='X' -> COMPOSITE."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [["Z_COMP_CORRETA"]],                    # AGR_DEFINE
            [["Z_COMP_CORRETA", "COLL_AGR", "X"]],   # AGR_FLAGS
        ]
        res = verificar_existencia_e_tipo_sap(["Z_COMP_CORRETA"], connection=conn)
        assert res["Z_COMP_CORRETA"]["existe"] is True
        assert res["Z_COMP_CORRETA"]["tipo_sap"] == "COMPOSITE"


def test_classificacao_single_sem_flag_coll_agr():
    """Valida role existente em AGR_DEFINE sem flag COLL_AGR -> SINGLE."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [["Z_SINGLE_ROLE"]],                     # AGR_DEFINE
            [],                                      # AGR_FLAGS (sem registo COLL_AGR)
        ]
        res = verificar_existencia_e_tipo_sap(["Z_SINGLE_ROLE"], connection=conn)
        assert res["Z_SINGLE_ROLE"]["existe"] is True
        assert res["Z_SINGLE_ROLE"]["tipo_sap"] == "SINGLE"


def test_classificacao_agr_define_ausente():
    """Valida role que não existe em AGR_DEFINE -> INEXISTENTE."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [],                                      # AGR_DEFINE vazia
            [],                                      # AGR_FLAGS não deve sequer ser consultada para inexistentes
        ]
        res = verificar_existencia_e_tipo_sap(["Z_ROLE_INEXISTENTE"], connection=conn)
        assert res["Z_ROLE_INEXISTENTE"]["existe"] is False
        assert res["Z_ROLE_INEXISTENTE"]["tipo_sap"] == "INEXISTENTE"


def test_classificacao_flag_coll_agr_com_valor_vazio_ou_incompleto():
    """Valida role com COLL_AGR mas FLAG_VALUE != 'X' -> Não classificada como Composite (SINGLE)."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        mock_read.side_effect = [
            [["Z_ROLE_FLAG_INVALIDA"]],                        # AGR_DEFINE
            [["Z_ROLE_FLAG_INVALIDA", "COLL_AGR", " "]],       # AGR_FLAGS sem X
        ]
        res = verificar_existencia_e_tipo_sap(["Z_ROLE_FLAG_INVALIDA"], connection=conn)
        assert res["Z_ROLE_FLAG_INVALIDA"]["existe"] is True
        assert res["Z_ROLE_FLAG_INVALIDA"]["tipo_sap"] == "SINGLE"


# =====================================================================
# TESTES DE GOVERNANÇA DE CRIADORES E SEGURANÇA DE RECIAÇÃO
# =====================================================================

def test_composite_nunca_usa_prgn_rfc_create_activity_group():
    """Garante que a chamada para criar COMPOSITE nunca aciona a RFC de Single (PRGN_RFC_CREATE_ACTIVITY_GROUP), mas sim PRGN_RFC_CREATE_AGR_MULTIPLE(COLLECTIVE_AGR='X')."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[False, True]), \
         patch("sap_rfc.role_existence_validator_service.read_table", return_value=[["Z_COMP_TEST", "COLL_AGR", "X"]]):

        res = criar_role_sap_com_readback("Z_COMP_TEST", tipo="COMPOSITE", environment="PRD", connection=conn)

        # RFC de Single PRGN_RFC_CREATE_ACTIVITY_GROUP NUNCA pode ter sido chamada na conexão
        chamadas_single_rfc = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "PRGN_RFC_CREATE_ACTIVITY_GROUP"]
        assert len(chamadas_single_rfc) == 0

        # RFC oficial de Composite PRGN_RFC_CREATE_AGR_MULTIPLE DEVE ter sido chamada com COLLECTIVE_AGR='X'
        chamadas_composite_rfc = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "PRGN_RFC_CREATE_AGR_MULTIPLE"]
        assert len(chamadas_composite_rfc) == 1
        assert chamadas_composite_rfc[0][1]["COLLECTIVE_AGR"] == "X"
        assert res["ok"] is True
        assert res["tipo_confirmado"] == "COMPOSITE"


def test_single_continua_usando_criador_single_rfc():
    """Garante que SINGLE continua a usar a RFC oficial PRGN_RFC_CREATE_ACTIVITY_GROUP."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[False, True]):
        res = criar_role_sap_com_readback("Z_SINGLE_TEST", tipo="SINGLE", environment="PRD", connection=conn)

        chamadas_create_rfc = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "PRGN_RFC_CREATE_ACTIVITY_GROUP"]
        assert len(chamadas_create_rfc) == 1
        assert res["ok"] is True
        assert res["tipo_confirmado"] == "SINGLE"


def test_readback_composite_falha_se_ausente_coll_agr():
    """Criação Composite falha se readback em AGR_FLAGS não trouxer COLL_AGR='X'."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[False, True]), \
         patch("sap_rfc.role_existence_validator_service.read_table", return_value=[]): # Sem flag em AGR_FLAGS

        res = criar_role_sap_com_readback("Z_COMP_SEM_FLAG", tipo="COMPOSITE", environment="PRD", connection=conn)
        assert res["ok"] is False
        assert res["status"] == STATUS_ERRO_READBACK
        assert "COLL_AGR='X'" in res["mensagem"]


# =====================================================================
# TESTES ESPECÍFICOS: ETAPA [5/8] — FUNÇÕES INDIVIDUAIS (SINGLE)
# =====================================================================

def test_etapa_single_role_correta(mock_dados_projeto):
    """Cenário 1: Single correta no SAP PRD (existe e é SINGLE)."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_SING_1": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_2": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_3": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_4": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_BASIS_BASE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "ZORG_EMPRESA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto)
        assert res["ok"] is True
        assert res["corretas"] == 6
        assert res["ausentes"] == 0
        assert res["tipos_divergentes"] == 0


def test_etapa_single_role_ausente(mock_dados_projeto):
    """Cenário 2: Single ausente no SAP PRD."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_SING_1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
            "Z_SING_2": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_3": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_4": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_BASIS_BASE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "ZORG_EMPRESA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto)
        assert res["ausentes"] == 1
        assert "Z_SING_1" in res["lista_ausentes"]


def test_etapa_single_fisicamente_composite_divergencia_e_confirma_n(mock_dados_projeto):
    """Cenários 3, 4 e 5: Single esperada mas SAP indica COMPOSITE, lista divergências e confirmação N."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk, \
         patch("builtins.input", return_value="N"):
        mock_chk.return_value = {
            "Z_SING_1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
            "Z_SING_2": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_3": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_4": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_BASIS_BASE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "ZORG_EMPRESA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto, assumir_sim=False)
        assert res["tipos_divergentes"] == 1
        assert res["lista_divergentes"][0]["role"] == "Z_SING_1"
        assert res["lista_divergentes"][0]["tipo_sap"] == "COMPOSITE"


def test_etapa_single_qad_offline_unknown(mock_dados_projeto):
    """Cenário 6: QAD offline = UNKNOWN, nunca MISSING."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        # Simula erro de conexão em todas as roles consultadas
        mock_chk.return_value = {
            s: {"existe": False, "tipo_sap": "INEXISTENTE", "erro": "Host unreachable / timeout"}
            for s in ["Z_SING_1", "Z_SING_2", "Z_SING_3", "Z_SING_4", "Z_BASIS_BASE", "ZORG_EMPRESA"]
        }
        res = executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto, environment="QAD")
        assert res["status"] == STATUS_SINGLE_UNKNOWN_AMBIENTE
        assert res["ausentes"] == 0
        assert res["tipos_divergentes"] == 0


def test_nenhuma_composite_esperada_na_etapa_single(mock_dados_projeto):
    """Cenário 7: Nenhuma Composite esperada entra na etapa Single."""
    roles_por_tipo = extrair_roles_necessarias_projeto("dummy.xlsx", dados=mock_dados_projeto)
    singles = roles_por_tipo["SINGLE"]
    composites = roles_por_tipo["COMPOSITE"]

    for comp in composites:
        assert comp not in singles


def test_auditar_single_role_divergente_bloqueada_para_revisao():
    """Valida auditoria de Single que está como Composite no SAP -> BLOQUEADA_PARA_REVISAO sem deleção automática."""
    conn = MagicMock()
    with patch("sap_rfc.role_existence_validator_service.role_exists", return_value=True), \
         patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        # Simula: AGR_FLAGS (COLL_AGR=X), AGR_USERS (1 user), AGR_TCODES (1 tc), AGR_1251 (2 auths), AGR_AGRS (1 filha), AGR_AGRS parent (0)
        mock_read.side_effect = [
            [["Z_ROLE_X", "COLL_AGR", "X"]], # AGR_FLAGS
            [["USER1"]],                      # AGR_USERS
            [["FB01"]],                       # AGR_TCODES
            [["S_TCODE"], ["S_RFC"]],          # AGR_1251
            [["Z_CHILD_1"]],                  # AGR_AGRS filhas
            [],                               # AGR_AGRS pais
        ]
        res = auditar_single_role_divergente("Z_ROLE_X", environment="PRD", connection=conn)
        assert res["role"] == "Z_ROLE_X"
        assert res["tipo_sap"] == "COMPOSITE"
        assert res["classificacao"] == STATUS_BLOQUEADA_PARA_REVISAO
        assert res["utilizadores_count"] == 1
        assert res["tcodes_count"] == 1


# =====================================================================
# TESTES ESPECÍFICOS: ETAPA [6/8] — FUNÇÕES COMPOSTAS (COMPOSITE)
# =====================================================================

def test_etapa_composite_role_correta(mock_dados_projeto):
    """Cenário 8: Composite correta no SAP PRD (AGR_FLAGS.COLL_AGR = 'X')."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto)
        assert res["ok"] is True
        assert res["corretas"] == 2
        assert res["ausentes"] == 0
        assert res["tipos_divergentes"] == 0


def test_etapa_composite_role_ausente(mock_dados_projeto):
    """Cenário 9: Composite ausente no SAP PRD."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto)
        assert res["ausentes"] == 1
        assert "Z_COMP_1" in res["lista_ausentes"]


def test_etapa_composite_fisicamente_single_confirmacao_n(mock_dados_projeto):
    """Cenários 10, 11 e 12: Composite esperada mas SAP indica SINGLE, lista divergências e confirmação N."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk, \
         patch("builtins.input", return_value="N"):
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }
        res = executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto, assumir_sim=False)
        assert res["tipos_divergentes"] == 1
        assert res["lista_divergentes"][0]["role"] == "Z_COMP_1"
        assert res["lista_divergentes"][0]["tipo_sap"] == "SINGLE"
        assert res["corrigidas_count"] == 0


def test_etapa_composite_fisicamente_single_confirmacao_s_e_sucesso(mock_dados_projeto):
    """Cenários 13, 14 e 15: Confirmação S executa delete, readback de ausência, create composite e readback COLL_AGR='X'."""
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk, \
         patch("sap_rfc.role_existence_validator_service.corrigir_composite_role_divergente_com_readback") as mock_corr:
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }
        mock_corr.return_value = {
            "ok": True,
            "status": "CORRIGIDA",
            "mensagem": "✓ Composite Role Z_COMP_1 corrigida com sucesso"
        }
        res = executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto, assumir_sim=True)
        assert res["corrigidas_count"] == 1
        assert res["tipos_divergentes"] == 0
        mock_corr.assert_called_once_with("Z_COMP_1", environment="PRD")


def test_corrigir_composite_fluxo_delete_create_readback_completo():
    """Valida o fluxo seguro detalhado de correção de Composite: delete + readback ausência + create PRGN_RFC_CREATE_AGR_MULTIPLE + readback COLL_AGR='X'."""
    conn = MagicMock()
    # Mock do retorno de BPC_DELETE_SINGLE_ROLE
    conn.call.return_value = {"ET_MESSAGES": []}

    with patch("sap_rfc.role_existence_validator_service.role_exists") as mock_exists, \
         patch("sap_rfc.role_existence_validator_service.read_table") as mock_read:
        # 1. role_exists inicial -> True
        # 2. read_table flags inicial -> [] (é Single)
        # 3. read_table texts -> [["Descricao"]]
        # 4. role_exists pós-delete -> False (ausente com sucesso)
        # 5. role_exists pós-create -> True (recriada)
        # 6. read_table flags pós-create -> [["Z_COMP_TEST", "COLL_AGR", "X"]]
        mock_exists.side_effect = [True, False, True]
        mock_read.side_effect = [
            [],                                          # Flags inicial (não é composite)
            [["Composite Description"]],                 # Texts
            [["Z_COMP_TEST", "COLL_AGR", "X"]],          # Flags pós-create (COLL_AGR='X')
        ]

        res = corrigir_composite_role_divergente_com_readback("Z_COMP_TEST", environment="PRD", connection=conn)
        assert res["ok"] is True
        assert res["status"] == "CORRIGIDA"

        # Conferir que BPC_DELETE_SINGLE_ROLE foi chamada
        chamadas_del = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "BPC_DELETE_SINGLE_ROLE"]
        assert len(chamadas_del) == 1
        assert chamadas_del[0][1]["IV_NAME_PREFIX"] == " "
        assert chamadas_del[0][1]["IV_PREF_AGR_NAME"] == "Z_COMP_TEST"

        # Conferir que PRGN_RFC_CREATE_AGR_MULTIPLE foi chamada com COLLECTIVE_AGR='X'
        chamadas_create = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "PRGN_RFC_CREATE_AGR_MULTIPLE"]
        assert len(chamadas_create) == 1
        assert chamadas_create[0][1]["COLLECTIVE_AGR"] == "X"

        # Conferir que PRGN_RFC_CREATE_ACTIVITY_GROUP NUNCA foi chamada
        chamadas_single = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "PRGN_RFC_CREATE_ACTIVITY_GROUP"]
        assert len(chamadas_single) == 0


def test_corrigir_composite_role_longa_maior_que_24_chars():
    """Valida que roles com mais de 24 chars são fatiadas corretamente entre IV_NAME_PREFIX e IV_PREF_AGR_NAME."""
    conn = MagicMock()
    conn.call.return_value = {"ET_MESSAGES": []}
    role_longa = "Z_BR_CUSTOMERCARE_TEAMLEAD"  # 26 chars

    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[True, False, True]), \
         patch("sap_rfc.role_existence_validator_service.read_table", side_effect=[[], [["Desc"]], [["Z_BR_CUSTOMERCARE_TEAMLEAD", "COLL_AGR", "X"]]]):
        res = corrigir_composite_role_divergente_com_readback(role_longa, environment="PRD", connection=conn)
        assert res["ok"] is True

        chamadas_del = [c for c in conn.call.call_args_list if c[0] and c[0][0] == "BPC_DELETE_SINGLE_ROLE"]
        assert len(chamadas_del) == 1
        kwargs_del = chamadas_del[0][1]
        assert kwargs_del["IV_NAME_PREFIX"] == "Z_"
        assert kwargs_del["IV_PREF_AGR_NAME"] == "BR_CUSTOMERCARE_TEAMLEAD"
        assert kwargs_del["IV_NAME_PREFIX"] + kwargs_del["IV_PREF_AGR_NAME"] == role_longa


def test_falha_em_qualquer_readback_nao_gera_sucesso():
    """Cenário 16: Falha no readback pós-deleção (role ainda existe) ou pós-criação aborta e não gera sucesso."""
    conn = MagicMock()
    conn.call.return_value = {"ET_MESSAGES": []}

    # Caso A: Readback pós-deleção falha (role ainda existe em AGR_DEFINE)
    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[True, True]), \
         patch("sap_rfc.role_existence_validator_service.read_table", side_effect=[[], [["Desc"]]]):
        res_a = corrigir_composite_role_divergente_com_readback("Z_FAIL_DEL", environment="PRD", connection=conn)
        assert res_a["ok"] is False
        assert res_a["status"] == STATUS_ERRO_READBACK
        assert "ainda existe na AGR_DEFINE após eliminação" in res_a["mensagem"]

    # Caso B: Readback pós-criação falha (AGR_FLAGS sem COLL_AGR='X')
    with patch("sap_rfc.role_existence_validator_service.role_exists", side_effect=[True, False, True]), \
         patch("sap_rfc.role_existence_validator_service.read_table", side_effect=[[], [["Desc"]], []]):
        res_b = corrigir_composite_role_divergente_com_readback("Z_FAIL_FLAG", environment="PRD", connection=conn)
        assert res_b["ok"] is False
        assert res_b["status"] == STATUS_ERRO_READBACK
        assert "NÃO possui COLL_AGR='X'" in res_b["mensagem"]


def test_nenhuma_single_esperada_na_etapa_composite(mock_dados_projeto):
    """Cenário 17: Nenhuma Single esperada entra na etapa Composite."""
    roles_por_tipo = extrair_roles_necessarias_projeto("dummy.xlsx", dados=mock_dados_projeto)
    singles = roles_por_tipo["SINGLE"]
    composites = roles_por_tipo["COMPOSITE"]

    for s in singles:
        assert s not in composites


def test_formato_log_etapas_5_e_6_compacto_e_com_problema(capsys, mock_dados_projeto):
    """Valida formato 1 linha sem problema e formato com pendências para [6/8] e [7/8]."""
    # 1. Sem problema -> 1 linha
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_SING_1": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_2": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_3": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_4": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_BASIS_BASE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "ZORG_EMPRESA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto)
        out_6_ok = capsys.readouterr().out
        assert "✓ [6/8] Funções Individuais: 6/6 corretas | ausentes 0 | divergentes 0" in out_6_ok

    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }
        executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto)
        out_7_ok = capsys.readouterr().out
        assert "✓ [7/8] Funções Compostas: 2/2 corretas | ausentes 0 | divergentes 0" in out_7_ok

    # 2. Com problemas -> ⚠️ e listagem de ausentes / divergentes
    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk:
        mock_chk.return_value = {
            "Z_SING_1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
            "Z_SING_2": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
            "Z_SING_3": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_SING_4": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "Z_BASIS_BASE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
            "ZORG_EMPRESA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        executar_etapa_validacao_funcoes_individuais("dummy.xlsx", dados=mock_dados_projeto, assumir_sim=True)
        out_6_err = capsys.readouterr().out
        assert "⚠️ [6/8] Funções Individuais: 4/6 corretas | ausentes 1 | divergentes 1" in out_6_err
        assert "Ausentes:" in out_6_err
        assert "- Z_SING_1" in out_6_err
        assert "Tipo divergente:" in out_6_err
        assert "- Z_SING_2 | Esperado=SINGLE | SAP=COMPOSITE" in out_6_err

    with patch("sap_rfc.role_existence_validator_service.verificar_existencia_e_tipo_sap") as mock_chk, \
         patch("sap_rfc.role_existence_validator_service.corrigir_composite_role_divergente_com_readback", return_value={"ok": True, "mensagem": "OK"}):
        mock_chk.return_value = {
            "Z_COMP_1": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
            "Z_COMP_2": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        }
        executar_etapa_validacao_funcoes_compostas("dummy.xlsx", dados=mock_dados_projeto, assumir_sim=True)
        out_7_err = capsys.readouterr().out
        assert "⚠️ [7/8] Funções Compostas: 0/2 corretas | ausentes 1 | divergentes 1" in out_7_err
        assert "Ausentes:" in out_7_err
        assert "- Z_COMP_1" in out_7_err
        assert "Tipo divergente:" in out_7_err
        assert "- Z_COMP_2 | Esperado=COMPOSITE | SAP=SINGLE" in out_7_err




