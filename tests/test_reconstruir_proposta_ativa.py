# -*- coding: utf-8 -*-
"""
tests/test_reconstruir_proposta_ativa.py
======================================================================
Testes unitários para a rotina de Reconstrução de Funções Individuais
na folha "Proposta Ativa" a partir das matrizes departamentais e da folha "Proposta".
======================================================================
"""

import sys
import pytest
import importlib.util
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

script_path = PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py"
spec = importlib.util.spec_from_file_location("projeto_perfil_mod", script_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


def test_a_uma_tcode_uma_funcao():
    """TESTE A — UMA TCODE / UMA FUNÇÃO."""
    tcode_map = {"TCODE_A": ["FUNCAO_A"]}
    tcodes_user = ["TCODE_A"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["FUNCAO_A"]
    assert sem_f == []


def test_b_uma_tcode_duas_funcoes():
    """TESTE B — UMA TCODE / DUAS FUNÇÕES."""
    tcode_map = {"TCODE_A": ["FUNCAO_A", "FUNCAO_B"]}
    tcodes_user = ["TCODE_A"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["FUNCAO_A", "FUNCAO_B"]
    assert sem_f == []


def test_c_varias_tcodes_mesma_funcao():
    """TESTE C — VÁRIAS TCODES / MESMA FUNÇÃO (Deduplicado na lista final)."""
    tcode_map = {
        "TCODE_A": ["FUNCAO_A"],
        "TCODE_B": ["FUNCAO_A"],
        "TCODE_C": ["FUNCAO_A"],
    }
    tcodes_user = ["TCODE_A", "TCODE_B", "TCODE_C"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["FUNCAO_A"]
    assert len(roles) == 1
    assert sem_f == []


def test_d_varias_tcodes_funcoes_sobrepostas():
    """TESTE D — VÁRIAS TCODES / FUNÇÕES SOBREPOSTAS."""
    tcode_map = {
        "TCODE_A": ["FUNCAO_A", "FUNCAO_B"],
        "TCODE_B": ["FUNCAO_B", "FUNCAO_C"],
    }
    tcodes_user = ["TCODE_A", "TCODE_B"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["FUNCAO_A", "FUNCAO_B", "FUNCAO_C"]
    assert sem_f == []


def test_e_ordem_deterministica():
    """TESTE E — ORDEM DETERMINÍSTICA DE DESCOBERTA."""
    tcode_map = {
        "TCODE_X": ["FUNCAO_Z", "FUNCAO_A"],
        "TCODE_Y": ["FUNCAO_M", "FUNCAO_Z"],
    }
    tcodes_user = ["TCODE_X", "TCODE_Y"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["FUNCAO_Z", "FUNCAO_A", "FUNCAO_M"]


def test_f_g_h_deteccao_dinamica_colunas():
    """TESTE F, G, H — Detecção dinâmica de 'Funções Individuais' nas colunas J, K e M."""
    class FakeDFJ:
        columns = ["User", "Nome", "Email", "Dep", "Cargo", "Comp", "Desc", "A1", "A2", "Funções Individuais"]
    class FakeDFK:
        columns = ["User", "Nome", "Email", "Dep", "Cargo", "Comp", "Desc", "A1", "A2", "A3", "Funções Individuais"]
    class FakeDFM:
        columns = ["User", "Nome", "Email", "Dep", "Cargo", "Comp", "Desc", "A1", "A2", "A3", "A4", "A5", "Funções Individuais"]

    assert mod.encontrar_coluna_funcoes_individuais_df(FakeDFJ()) == 9   # Coluna J (índice 9)
    assert mod.encontrar_coluna_funcoes_individuais_df(FakeDFK()) == 10  # Coluna K (índice 10)
    assert mod.encontrar_coluna_funcoes_individuais_df(FakeDFM()) == 12  # Coluna M (índice 12)


def test_i_cabecalho_ausente_gera_erro():
    """TESTE I — Se 'Funções Individuais' não existir, lança erro sem fallback."""
    class FakeDFInvalido:
        columns = ["User", "Nome", "Email", "Departamento", "Composite Role"]

    with pytest.raises(ValueError, match="CABECALHO_FUNCOES_INDIVIDUAIS_NAO_ENCONTRADO"):
        mod.encontrar_coluna_funcoes_individuais_df(FakeDFInvalido())


def test_j_vk13_multifuncao_real():
    """TESTE J — VK13 MULTIFUNÇÃO: Retorna Z_SALES_PRICECOND_CREATE e Z_SALES_PRICECOND_REPORT."""
    caminho = mod.encontrar_excel_padrao()
    dados = mod.carregar_projeto_perfil(caminho)
    tcode_map = mod.mapear_tcodes_sheet_proposta(dados)

    roles_vk13 = tcode_map.get("VK13", [])
    assert "Z_SALES_PRICECOND_CREATE" in roles_vk13
    assert "Z_SALES_PRICECOND_REPORT" in roles_vk13
    assert len(roles_vk13) >= 2


def test_composite_role_permanece_inalterada():
    """TESTE — Composite Role permanece inalterada."""
    row_original = {
        "Usuario": "S270",
        "Composite Role": "Z_BR_PURCHSERV_TEAMLEAD",
        "Funcoes": ["Z_OLD_1", "Z_OLD_2"]
    }
    novas_singles = ["Z_ASSET_CREATE", "Z_PROJECT_CREATE"]
    row_reconstruida = dict(row_original)
    row_reconstruida["Funcoes"] = novas_singles
    assert row_reconstruida["Composite Role"] == row_original["Composite Role"]


def test_transacao_inexistente_na_proposta():
    """TESTE — Transação inexistente na Proposta (registada em sem_funcao)."""
    tcode_map = {"AS02": ["Z_ASSET_CREATE"]}
    tcodes_user = ["AS02", "TCODE_DESCONHECIDO"]
    roles, sem_f = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert roles == ["Z_ASSET_CREATE"]
    assert sem_f == ["TCODE_DESCONHECIDO"]


def test_proposta_ativa_sem_alteracoes():
    """TESTE 14.A — Utilizador sem alterações -> Proposta Ativa permanece igual."""
    tcode_map = {"ME23N": ["Z_PURCHASE_ORDER_DISPLAY"]}
    tcodes_user = ["ME23N"]
    roles_calc, _ = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    roles_atuais = ["Z_PURCHASE_ORDER_DISPLAY"]
    assert roles_calc == roles_atuais


def test_proposta_ativa_nova_tcode():
    """TESTE 14.B — Utilizador recebe nova TCODE -> nova Função Individual entra na Proposta Ativa."""
    tcode_map = {"ME21N": ["Z_PURCHASE_ORDER_CREATE"], "ME23N": ["Z_PURCHASE_ORDER_DISPLAY"]}
    tcodes_user = ["ME23N", "ME21N"]
    roles_calc, _ = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert "Z_PURCHASE_ORDER_CREATE" in roles_calc
    assert "Z_PURCHASE_ORDER_DISPLAY" in roles_calc


def test_proposta_ativa_idempotencia():
    """TESTE 14.E — Execução repetida -> idempotente (mesmas funções calculadas)."""
    tcode_map = {"ME21N": ["Z_PURCHASE_ORDER_CREATE"], "ME23N": ["Z_PURCHASE_ORDER_DISPLAY"]}
    tcodes_user = ["ME23N", "ME21N"]
    calc1, _ = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    calc2, _ = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)
    assert calc1 == calc2


def test_federico_proposta_ativa_tres_novas_tcodes():
    """TESTE 14.F — Federico: entrada ME21N, ME22N, ME23N -> Z_PURCHASE_ORDER_CREATE, Z_PURCHASE_ORDER_DISPLAY, Z_PURCHASE_ORDER_REPORT presentes na Proposta Ativa."""
    tcode_map = {
        "ME21N": ["Z_PURCHASE_ORDER_CREATE"],
        "ME22N": ["Z_PURCHASE_ORDER_DISPLAY"],
        "ME23N": ["Z_PURCHASE_ORDER_REPORT"],
        "ME51N": ["Z_PURCHASE_REQ_CREATE"],
        "ME52N": ["Z_PURCHASE_REQ_DISPLAY"],
        "ME53N": ["Z_PURCHASE_REQ_REPORT"],
        "MIGO": ["Z_GOODS_MOVEMENTS"],
        "KS03": ["Z_COSTCENTER_CREATE"],
    }
    tcodes_federico = ["ME21N", "ME22N", "ME23N", "ME51N", "ME52N", "ME53N", "MIGO", "KS03"]
    roles_federico, _ = mod.mapeamento_tcodes_para_funcoes_individuais(tcodes_federico, tcode_map)

    assert len(roles_federico) == 8
    assert "Z_PURCHASE_ORDER_CREATE" in roles_federico
    assert "Z_PURCHASE_ORDER_DISPLAY" in roles_federico
    assert "Z_PURCHASE_ORDER_REPORT" in roles_federico
