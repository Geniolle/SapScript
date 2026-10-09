# -*- coding: utf-8 -*-
"""
tests/test_reconciliar_proposta_ativa_com_prd.py
======================================================================
Testes unitários abrangentes para o Requisito 23:
A. Single Role direta em PRD e já existe na Proposta Ativa -> nenhuma alteração.
B. Single Role direta em PRD não existe na Proposta Ativa -> candidata a adicionar.
C. Composite Role atribuída -> não deve ser adicionada às Funções Individuais.
D. Single Role apenas herdada por Composite -> não deve ser adicionada como individual.
E. Role expirada -> ignorar.
F. Função existente no Excel mas não em PRD -> reportar, não remover.
G. Role duplicada no Excel -> não duplicar novamente.
H. Utilizador externo ao Excel -> ignorar.
I. Execução repetida -> idempotente.
======================================================================
"""

import sys
import pytest
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def simular_reconciliacao_proposta_ativa_utilizador(
    uid_excel: str,
    funcs_atuais_excel: list,
    comp_role_excel: str,
    registos_agr_users_prd: list,
    funcoes_excluidas_definicoes: set = None,
    today_str: str = None
):
    """
    Helper puro de testes que executa a lógica de reconciliação para 1 utilizador.
    """
    if today_str is None:
        today_str = datetime.now().strftime("%Y%m%d")

    if funcoes_excluidas_definicoes is None:
        funcoes_excluidas_definicoes = set()

    # Normalizar funcoes excluidas (case-insensitive + trim)
    funcoes_excluidas_norm = {str(f).strip().upper() for f in funcoes_excluidas_definicoes if f and str(f).strip()}

    # 1. Normalizar e deduplicar funcoes atuais
    funcs_ex_set = set()
    for f in funcs_atuais_excel:
        if f:
            val = str(f).strip().upper()
            if val and val.startswith("Z") and not val.startswith("Z_BR_") and len(val) >= 4:
                funcs_ex_set.add(val)

    # 2. Filtrar AGR_USERS para Single Roles diretas ativas
    funcs_prd_set = set()
    for r in registos_agr_users_prd:
        uname, agr_name, f_dat, t_dat, col_flag = [str(x or "").strip() for x in r]
        uname = uname.upper()
        agr_name = agr_name.upper()

        if uname == uid_excel.upper() and agr_name:
            if (not f_dat or f_dat <= today_str) and (not t_dat or t_dat >= today_str):
                if col_flag != "X" and not agr_name.startswith("Z_BR_"):
                    funcs_prd_set.add(agr_name)

    # 3. Filtrar pelas DEFINIÇÕES
    funcs_excluidas = funcs_prd_set.intersection(funcoes_excluidas_norm)
    funcs_elegiveis = funcs_prd_set - funcoes_excluidas_norm

    # 4. Comparação
    a_adicionar = sorted(list(funcs_elegiveis - funcs_ex_set))
    nao_confirmadas = sorted(list(funcs_ex_set - funcs_prd_set))

    funcs_finais = sorted(list(funcs_ex_set.union(set(a_adicionar))))

    return {
        "user_id": uid_excel,
        "funcs_atuais_excel": sorted(list(funcs_ex_set)),
        "funcs_prd_diretas": sorted(list(funcs_prd_set)),
        "funcs_excluidas_definicoes": sorted(list(funcs_excluidas)),
        "funcs_elegiveis": sorted(list(funcs_elegiveis)),
        "a_adicionar": a_adicionar,
        "nao_confirmadas": nao_confirmadas,
        "funcs_finais": funcs_finais
    }


def test_req23_a_role_direta_ja_existe_na_proposta_ativa():
    """A. Single Role direta em PRD já existe na Proposta Ativa -> Nenhuma alteração."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_PURCHASE_ORDER_CREATE"],
        comp_role_excel="Z_BR_TEST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == []
    assert res["nao_confirmadas"] == []
    assert res["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_req23_b_role_direta_nao_existe_na_proposta_ativa():
    """B. Single Role direta em PRD não existe na Proposta Ativa -> Candidata a adicionar."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_PURCHASE_ORDER_DISPLAY"],
        comp_role_excel="Z_BR_TEST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", ""),
            ("S123", "Z_PURCHASE_ORDER_DISPLAY", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]
    assert res["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE", "Z_PURCHASE_ORDER_DISPLAY"]


def test_req23_c_composite_role_atribuida():
    """C. Composite Role atribuída -> Não deve ser adicionada às Funções Individuais."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="Z_BR_PURCHREQ&PO_SPECIALIST",
        registos_agr_users_prd=[
            ("S123", "Z_BR_PURCHREQ&PO_SPECIALIST", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == []
    assert "Z_BR_PURCHREQ&PO_SPECIALIST" not in res["funcs_finais"]


def test_req23_d_single_role_apenas_herdada_por_composite():
    """D. Single Role apenas herdada por Composite (COL_FLAG == 'X') -> Não adicionar como individual."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="Z_BR_PURCHREQ&PO_SPECIALIST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", "X")  # COL_FLAG='X' -> Herdada
        ]
    )
    assert res["a_adicionar"] == []


def test_req23_e_role_expirada():
    """E. Role expirada em PRD -> Ignorar."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_EXPIRED_ROLE", "20200101", "20221231", "")  # Expirada
        ],
        today_str="20261009"
    )
    assert res["a_adicionar"] == []


def test_req23_f_funcao_excel_nao_em_prd():
    """F. Função existente no Excel mas não em PRD -> Reportar em não confirmadas, NÃO remover."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_OLD_ROLE_IN_EXCEL"],
        comp_role_excel="",
        registos_agr_users_prd=[]
    )
    assert res["nao_confirmadas"] == ["Z_OLD_ROLE_IN_EXCEL"]
    assert res["funcs_finais"] == ["Z_OLD_ROLE_IN_EXCEL"]  # Preservada!


def test_req23_g_role_duplicada_no_excel():
    """G. Role duplicada no Excel -> Não duplicar novamente."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_ROLE_A", "Z_ROLE_A"],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_ROLE_A", "20200101", "99991231", "")
        ]
    )
    assert res["funcs_atuais_excel"] == ["Z_ROLE_A"]
    assert res["funcs_finais"] == ["Z_ROLE_A"]


def test_req23_h_utilizador_externo_ao_excel():
    """H. Utilizador externo ao Excel -> Ignorar."""
    registos = [("S999_EXTERNO", "Z_ROLE_X", "20200101", "99991231", "")]
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=registos
    )
    assert res["a_adicionar"] == []


def test_req23_i_execucao_repetida_idempotente():
    """I. Execução repetida -> Idempotente (2ª execução resulta em a_adicionar = 0)."""
    registos_prd = [
        ("S123", "Z_ROLE_NEW", "20200101", "99991231", "")
    ]
    res1 = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=registos_prd
    )
    assert res1["a_adicionar"] == ["Z_ROLE_NEW"]

    res2 = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=res1["funcs_finais"],
        comp_role_excel="",
        registos_agr_users_prd=registos_prd
    )
    assert res2["a_adicionar"] == []
    assert res2["funcs_finais"] == ["Z_ROLE_NEW"]


# Novas validações obrigatórias do Requisito 8 (Sheet DEFINIÇÕES)

def test_req8_a_role_consta_em_definicoes():
    """A. Role direta PRD consta na DEFINIÇÕES -> NÃO entra em FUNCOES_A_ADICIONAR."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_BASIS_BASE", "20200101", "99991231", ""),
            ("S123", "Z_MY_HOME", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={"Z_BASIS_BASE", "Z_MY_HOME"}
    )
    assert res["funcs_excluidas_definicoes"] == ["Z_BASIS_BASE", "Z_MY_HOME"]
    assert res["funcs_elegiveis"] == []
    assert res["a_adicionar"] == []


def test_req8_b_role_nao_consta_em_definicoes():
    """B. Role direta PRD não consta na DEFINIÇÕES e não está na Proposta Ativa -> entra em FUNCOES_A_ADICIONAR."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_CUSTOM_ROLE_VALID", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={"Z_BASIS_BASE"}
    )
    assert res["funcs_excluidas_definicoes"] == []
    assert res["funcs_elegiveis"] == ["Z_CUSTOM_ROLE_VALID"]
    assert res["a_adicionar"] == ["Z_CUSTOM_ROLE_VALID"]


def test_req8_c_role_definicoes_ja_no_excel():
    """C. Role consta na DEFINIÇÕES e já existe na Proposta Ativa -> NÃO remover."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_BASIS_BASE"],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_BASIS_BASE", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={"Z_BASIS_BASE"}
    )
    assert res["a_adicionar"] == []
    assert res["funcs_finais"] == ["Z_BASIS_BASE"]  # Mantida no Excel!


def test_req8_d_comparacao_case_insensitive_trim():
    """D. Comparação deve ser case-insensitive e trim-safe."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", " z_my_home ", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={"Z_MY_HOME"}
    )
    assert res["funcs_excluidas_definicoes"] == ["Z_MY_HOME"]
    assert res["a_adicionar"] == []


def test_req8_e_role_vazia_em_definicoes():
    """E. Role vazia na DEFINIÇÕES -> ignorar."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_CUSTOM_ROLE", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={" ", "", None}
    )
    assert res["a_adicionar"] == ["Z_CUSTOM_ROLE"]


def test_req8_f_role_duplicada_em_definicoes():
    """F. Role duplicada na DEFINIÇÕES -> considerar apenas uma vez."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_BASIS_BASE", "20200101", "99991231", "")
        ],
        funcoes_excluidas_definicoes={"Z_BASIS_BASE", "Z_BASIS_BASE", "z_basis_base"}
    )
    assert res["funcs_excluidas_definicoes"] == ["Z_BASIS_BASE"]
    assert res["a_adicionar"] == []



def test_req23_a_role_direta_ja_existe_na_proposta_ativa():
    """A. Single Role direta em PRD já existe na Proposta Ativa -> Nenhuma alteração."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_PURCHASE_ORDER_CREATE"],
        comp_role_excel="Z_BR_TEST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == []
    assert res["nao_confirmadas"] == []
    assert res["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE"]


def test_req23_b_role_direta_nao_existe_na_proposta_ativa():
    """B. Single Role direta em PRD não existe na Proposta Ativa -> Candidata a adicionar."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_PURCHASE_ORDER_DISPLAY"],
        comp_role_excel="Z_BR_TEST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", ""),
            ("S123", "Z_PURCHASE_ORDER_DISPLAY", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == ["Z_PURCHASE_ORDER_CREATE"]
    assert res["funcs_finais"] == ["Z_PURCHASE_ORDER_CREATE", "Z_PURCHASE_ORDER_DISPLAY"]


def test_req23_c_composite_role_atribuida():
    """C. Composite Role atribuída -> Não deve ser adicionada às Funções Individuais."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="Z_BR_PURCHREQ&PO_SPECIALIST",
        registos_agr_users_prd=[
            ("S123", "Z_BR_PURCHREQ&PO_SPECIALIST", "20200101", "99991231", "")
        ]
    )
    assert res["a_adicionar"] == []
    assert "Z_BR_PURCHREQ&PO_SPECIALIST" not in res["funcs_finais"]


def test_req23_d_single_role_apenas_herdada_por_composite():
    """D. Single Role apenas herdada por Composite (COL_FLAG == 'X') -> Não adicionar como individual."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="Z_BR_PURCHREQ&PO_SPECIALIST",
        registos_agr_users_prd=[
            ("S123", "Z_PURCHASE_ORDER_CREATE", "20200101", "99991231", "X")  # COL_FLAG='X' -> Herdada
        ]
    )
    assert res["a_adicionar"] == []


def test_req23_e_role_expirada():
    """E. Role expirada em PRD -> Ignorar."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_EXPIRED_ROLE", "20200101", "20221231", "")  # Expirada
        ],
        today_str="20261009"
    )
    assert res["a_adicionar"] == []


def test_req23_f_funcao_excel_nao_em_prd():
    """F. Função existente no Excel mas não em PRD -> Reportar em não confirmadas, NÃO remover."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_OLD_ROLE_IN_EXCEL"],
        comp_role_excel="",
        registos_agr_users_prd=[]
    )
    assert res["nao_confirmadas"] == ["Z_OLD_ROLE_IN_EXCEL"]
    assert res["funcs_finais"] == ["Z_OLD_ROLE_IN_EXCEL"]  # Preservada!


def test_req23_g_role_duplicada_no_excel():
    """G. Role duplicada no Excel -> Não duplicar novamente."""
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=["Z_ROLE_A", "Z_ROLE_A"],
        comp_role_excel="",
        registos_agr_users_prd=[
            ("S123", "Z_ROLE_A", "20200101", "99991231", "")
        ]
    )
    assert res["funcs_atuais_excel"] == ["Z_ROLE_A"]
    assert res["funcs_finais"] == ["Z_ROLE_A"]


def test_req23_h_utilizador_externo_ao_excel():
    """H. Utilizador externo ao Excel -> Ignorar."""
    # A população é delimitada pelo Excel. Se o ID não for processado, não entra.
    registos = [("S999_EXTERNO", "Z_ROLE_X", "20200101", "99991231", "")]
    # Testamos que para o S123 do Excel, o S999_EXTERNO é totalmente irrelevante.
    res = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=registos
    )
    assert res["a_adicionar"] == []


def test_req23_i_execucao_repetida_idempotente():
    """I. Execução repetida -> Idempotente (2ª execução resulta em a_adicionar = 0)."""
    registos_prd = [
        ("S123", "Z_ROLE_NEW", "20200101", "99991231", "")
    ]
    # Passagem 1: Excel tem []
    res1 = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=[],
        comp_role_excel="",
        registos_agr_users_prd=registos_prd
    )
    assert res1["a_adicionar"] == ["Z_ROLE_NEW"]

    # Passagem 2: Excel agora tem funcs_finais da Passagem 1
    res2 = simular_reconciliacao_proposta_ativa_utilizador(
        uid_excel="S123",
        funcs_atuais_excel=res1["funcs_finais"],
        comp_role_excel="",
        registos_agr_users_prd=registos_prd
    )
    assert res2["a_adicionar"] == []
    assert res2["funcs_finais"] == ["Z_ROLE_NEW"]
