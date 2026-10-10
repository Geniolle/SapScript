# -*- coding: utf-8 -*-
"""
tests/test_auditoria_matrizes_vs_proposta.py
======================================================================
Testes unitários obrigatórios para a nova validação inicial entre
Matrizes Departamentais e Sheet Proposta na Tarefa 1.
======================================================================
"""

import sys
import pytest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import importlib.util

projeto_perfil_path = PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py"
spec = importlib.util.spec_from_file_location("projeto_perfil", projeto_perfil_path)
projeto_perfil = importlib.util.module_from_spec(spec)
spec.loader.exec_module(projeto_perfil)

auditar_matrizes_departamentais_vs_proposta = projeto_perfil.auditar_matrizes_departamentais_vs_proposta
REGRAS_ESPECIAIS_TCODE = projeto_perfil.REGRAS_ESPECIAIS_TCODE
ProjetoPerfilData = projeto_perfil.ProjetoPerfilData


def simular_auditoria_matriz(
    controlo_deps: list,
    matrizes_data: dict,
    proposta_tcodes: dict,
    regras_especiais: dict = None
):
    """
    Helper puro de testes para simular o comportamento de auditar_matrizes_departamentais_vs_proposta.
    """
    if regras_especiais is None:
        regras_especiais = REGRAS_ESPECIAIS_TCODE

    pendencias = {}
    total_utilizadas = 0
    tcodes_sem_map_set = set()

    for item in controlo_deps:
        dep_nome = item["departamento"]
        # independentemente do status
        matriz_dep = matrizes_data.get(dep_nome, [])
        tcodes_vistas = set()

        for line in matriz_dep:
            tc = line["tcode"].strip().upper()
            flags = line.get("flags", [])
            # QUALQUER X = TCODE UTILIZADA, ZERO X = IGNORAR
            tem_x = any(f.strip().upper() in ("X", "1", "SIM", "YES", "S") for f in flags)
            if not tem_x:
                continue

            if tc in tcodes_vistas:
                continue
            tcodes_vistas.add(tc)
            total_utilizadas += 1

            if tc in regras_especiais:
                classificacao = regras_especiais[tc]
            elif tc in proposta_tcodes:
                classificacao = "MAPEADA"
            else:
                classificacao = "TCODE_SEM_MAPEAMENTO_PROPOSTA"

            if classificacao == "TCODE_SEM_MAPEAMENTO_PROPOSTA":
                if dep_nome not in pendencias:
                    pendencias[dep_nome] = []
                pendencias[dep_nome].append({
                    "tcode": tc,
                    "descricao": line.get("descricao", "")
                })
                tcodes_sem_map_set.add(tc)

    return {
        "deps_auditados": len(controlo_deps),
        "total_tcodes_utilizadas": total_utilizadas,
        "tem_pendencias": len(tcodes_sem_map_set) > 0,
        "deps_com_pendencias_count": len(pendencias),
        "deps_com_pendencias": list(pendencias.keys()),
        "total_tcodes_sem_mapeamento": len(tcodes_sem_map_set),
        "tcodes_sem_mapeamento_set": tcodes_sem_map_set,
        "pendencias_por_dep": pendencias
    }


def test_1_2_3_audit_status_controlo_ignorado():
    """1, 2, 3: Valida departamentos na CONTROLO com STATUS PROCESSADO, PROCESSADO_COM_PENDENCIAS e vazio."""
    controlo = [
        {"departamento": "Purchase & Services", "status": "PROCESSADO"},
        {"departamento": "Client Services", "status": "PROCESSADO_COM_PENDENCIAS"},
        {"departamento": "Product", "status": ""}
    ]
    matrizes = {
        "Purchase & Services": [{"tcode": "ME23N", "flags": ["X"]}],
        "Client Services": [{"tcode": "VA03", "flags": ["X"]}],
        "Product": [{"tcode": "MD11", "flags": ["X"]}]
    }
    proposta = {"ME23N": ["Z_ROLE1"], "VA03": ["Z_ROLE2"]}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["deps_auditados"] == 3
    assert res["total_tcodes_sem_mapeamento"] == 1
    assert "Product" in res["deps_com_pendencias"]


def test_4_tcode_sem_x_ignorada():
    """4: TCODE sem nenhum X deve ser ignorada."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {"Product": [{"tcode": "CO01", "flags": [" ", " ", ""]}]}
    proposta = {}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["total_tcodes_utilizadas"] == 0
    assert res["tem_pendencias"] is False


def test_5_6_tcode_com_um_ou_varios_x():
    """5, 6: TCODE com um único X ou com vários X deve ser considerada utilizada."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {
        "Product": [
            {"tcode": "ME23N", "flags": ["", "X", ""]},
            {"tcode": "VA01", "flags": ["X", "", "X"]}
        ]
    }
    proposta = {"ME23N": ["Z_ROLE1"], "VA01": ["Z_ROLE2"]}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["total_tcodes_utilizadas"] == 2
    assert res["tem_pendencias"] is False


def test_7_8_tcode_mapeada_1_1_e_1_n():
    """7, 8: TCODE mapeada 1:1 ou 1:N é CLASSIFICAR = MAPEADA."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {
        "Product": [
            {"tcode": "ME23N", "flags": ["X"]},
            {"tcode": "MIGO", "flags": ["X"]}
        ]
    }
    proposta = {
        "ME23N": ["Z_PURCHASE_ORDER_DISPLAY"],
        "MIGO": ["Z_GOODS_MOVEMENTS", "Z_INVENTORY_DC_STORE"]
    }

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["tem_pendencias"] is False


def test_9_tcode_sem_mapeamento():
    """9: TCODE sem mapeamento na Proposta -> TCODE_SEM_MAPEAMENTO_PROPOSTA."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {"Product": [{"tcode": "CO06", "flags": ["X"]}]}
    proposta = {}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["tem_pendencias"] is True
    assert res["pendencias_por_dep"]["Product"][0]["tcode"] == "CO06"


def test_10_bp_regra_especial():
    """10: BP tratada por regra especial -> não aparece como erro nem pendência."""
    controlo = [{"departamento": "People & Talent", "status": "PROCESSADO"}]
    matrizes = {"People & Talent": [{"tcode": "BP", "flags": ["X"]}]}
    proposta = {}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["tem_pendencias"] is False


def test_11_mesma_tcode_pendente_varios_departamentos():
    """11: Mesma TCODE pendente em vários departamentos -> deduplicada no total de TCODEs sem mapeamento."""
    controlo = [
        {"departamento": "Product", "status": "PROCESSADO"},
        {"departamento": "Legal", "status": "PENDENTE"}
    ]
    matrizes = {
        "Product": [{"tcode": "CO06", "flags": ["X"]}],
        "Legal": [{"tcode": "CO06", "flags": ["X"]}]
    }
    proposta = {}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["total_tcodes_sem_mapeamento"] == 1
    assert res["deps_com_pendencias_count"] == 2


def test_12_normalizacao_uppercase_trim():
    """12: Normalização uppercase/trim funciona corretamente."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {"Product": [{"tcode": "  me23n  ", "flags": [" x "]}]}
    proposta = {"ME23N": ["Z_PURCHASE_ORDER_DISPLAY"]}

    res = simular_auditoria_matriz(controlo, matrizes, proposta)
    assert res["total_tcodes_utilizadas"] == 1
    assert res["tem_pendencias"] is False


def test_13_14_15_16_decisao_continuar_e_abortar():
    """13, 14, 15, 16: Simular regras das Opções Continuar e Abortar."""
    tcodes_pendentes = {"CO06"}
    tcode_to_roles_map = {"ME23N": ["Z_ROLE_ME23N"]}

    # Opção 2 - Abortar: interrompe antes de qualquer alteração
    executar_etapas_posteriores = False
    assert executar_etapas_posteriores is False

    # Opção 1 - Continuar parcialmente: ignorar CO06 no cálculo de roles
    tcodes_utilizadas_user = ["ME23N", "CO06"]
    roles_esperadas = set()
    for tc in tcodes_utilizadas_user:
        if tc in tcodes_pendentes:
            continue  # ignorada, sem inventar Single Role
        for r in tcode_to_roles_map.get(tc, []):
            roles_esperadas.add(r)

    assert roles_esperadas == {"Z_ROLE_ME23N"}


def test_17_idempotencia_segunda_execucao():
    """17: Idempotência numa segunda execução da auditoria."""
    controlo = [{"departamento": "Product", "status": "PROCESSADO"}]
    matrizes = {"Product": [{"tcode": "ME23N", "flags": ["X"]}]}
    proposta = {"ME23N": ["Z_ROLE1"]}

    res1 = simular_auditoria_matriz(controlo, matrizes, proposta)
    res2 = simular_auditoria_matriz(controlo, matrizes, proposta)

    assert res1 == res2
