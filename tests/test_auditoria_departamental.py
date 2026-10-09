# -*- coding: utf-8 -*-
"""
tests/test_auditoria_departamental.py
======================================================================
Testes unitários para a rotina de Auditoria Departamental Pós-Processamento.
Garante a correta classificação de divergências entre Excel, PRD e QAS.
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
projeto_perfil_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(projeto_perfil_mod)

comparar_utilizador_ambientes = projeto_perfil_mod.comparar_utilizador_ambientes

def test_comparacao_ok():
    """Teste 1 — tudo igual (Excel = PRD = QAS)."""
    excel_esp = {
        "composta": "ZC_BUYER",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01"},
        "roles_herdadas": {"ZMM_SINGLE_01", "ZMM_SINGLE_02"},
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_SINGLE_01", "ZMM_SINGLE_02"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01"},
        "roles_herdadas": {"ZMM_SINGLE_01", "ZMM_SINGLE_02"},
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_SINGLE_01", "ZMM_SINGLE_02"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01"},
        "roles_herdadas": {"ZMM_SINGLE_01", "ZMM_SINGLE_02"},
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_SINGLE_01", "ZMM_SINGLE_02"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert res["status_prd"] == "OK"
    assert res["status_qas"] == "OK"
    assert res["status_geral"] == "OK"
    assert "OK" in res["classificacoes"]
    assert len(res["faltam_prd"]) == 0
    assert len(res["faltam_qas"]) == 0


def test_comparacao_falta_prd():
    """Teste 2 — falta no PRD."""
    excel_esp = {
        "composta": "ZC_BUYER",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert res["status_prd"] == "DIVERGENTE"
    assert res["status_qas"] == "OK"
    assert "DIVERGENTE_PRD" in res["classificacoes"]
    assert "ROLE_EM_FALTA" in res["classificacoes"]
    assert res["faltam_prd"] == ["ZMM_ROLE_02"]


def test_comparacao_falta_qas():
    """Teste 3 — falta no QAS."""
    excel_esp = {
        "composta": "ZC_BUYER",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01", "ZMM_ROLE_02"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER", "ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZC_BUYER", "ZMM_ROLE_01"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert res["status_prd"] == "OK"
    assert res["status_qas"] == "DIVERGENTE"
    assert "DIVERGENTE_QAS" in res["classificacoes"]
    assert "ROLE_EM_FALTA" in res["classificacoes"]
    assert res["faltam_qas"] == ["ZMM_ROLE_02"]


def test_comparacao_role_extra_prd():
    """Teste 4 — role extra PRD."""
    excel_esp = {
        "composta": None,
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01", "ZMM_EXTRA_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01", "ZMM_EXTRA_01"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert "DIVERGENTE_PRD" in res["classificacoes"]
    assert "ROLE_EXTRA" in res["classificacoes"]
    assert res["extras_prd"] == ["ZMM_EXTRA_01"]


def test_comparacao_prd_diferente_qas():
    """Teste 5 — PRD ≠ QAS."""
    excel_esp = {
        "composta": None,
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01", "ZMM_EXTRA_PRD"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01", "ZMM_EXTRA_PRD"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01", "ZMM_EXTRA_QAS"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01", "ZMM_EXTRA_QAS"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert "PRD_QAS_DIVERGENTE" in res["classificacoes"]
    assert res["prd_nao_qas"] == ["ZMM_EXTRA_PRD"]
    assert res["qas_nao_prd"] == ["ZMM_EXTRA_QAS"]


def test_comparacao_composite_role_heranca():
    """Teste 6 — composite role (não confundir herdadas com diretas em falta)."""
    excel_esp = {
        "composta": "ZC_BUYER",
        "roles_diretas": {"ZC_BUYER"},
        "roles_herdadas": {"ZMM_SINGLE_01"},
        "roles_totais": {"ZC_BUYER", "ZMM_SINGLE_01"}
    }
    # No PRD, a composite está atribuída e a single entra via herança (COL_FLAG='X')
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER"},
        "roles_herdadas": {"ZMM_SINGLE_01"},
        "roles_totais": {"ZC_BUYER", "ZMM_SINGLE_01"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZC_BUYER"},
        "roles_herdadas": {"ZMM_SINGLE_01"},
        "roles_totais": {"ZC_BUYER", "ZMM_SINGLE_01"}
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert res["status_geral"] == "OK"
    assert res["faltam_prd"] == []
    assert res["extras_prd"] == []


def test_comparacao_utilizador_inexistente():
    """Teste 7 — utilizador inexistente num dos ambientes (ex: QAS)."""
    excel_esp = {
        "composta": None,
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    qas_dados = {
        "status_conta": "NAO_ENCONTRADO",
        "roles_diretas": set(),
        "roles_herdadas": set(),
        "roles_totais": set()
    }

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados)
    assert "UTILIZADOR_NAO_ENCONTRADO" in res["classificacoes"]
    assert res["status_qas"] == "DIVERGENTE"


def test_comparacao_protegida_por_exclusao():
    """Teste 8 — role extra que bate num padrão da sheet EXCLUÇÃO."""
    excel_esp = {
        "composta": None,
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    # Role extra: ZMM_APROVA_PEDC_COD_01 (supondo que padroes_exclusao tem ZMM_APROVA_PEDC_COD_*)
    prd_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01", "ZMM_APROVA_PEDC_COD_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01", "ZMM_APROVA_PEDC_COD_01"}
    }
    qas_dados = {
        "status_conta": "ATIVA",
        "roles_diretas": {"ZMM_ROLE_01"},
        "roles_herdadas": set(),
        "roles_totais": {"ZMM_ROLE_01"}
    }
    padroes_exclusao = ["ZMM_APROVA_PEDC_COD_*"]

    res = comparar_utilizador_ambientes("S5092", "Clayton Lopes", excel_esp, prd_dados, qas_dados, padroes_exclusao=padroes_exclusao)
    assert "PROTEGIDA_POR_EXCLUSAO" in res["classificacoes"]
    assert res["protegidas_exclusao_prd"] == ["ZMM_APROVA_PEDC_COD_01"]
    assert res["extras_prd"] == []  # Não é classificada como extra limpa
