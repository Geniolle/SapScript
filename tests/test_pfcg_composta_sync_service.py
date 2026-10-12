# -*- coding: utf-8 -*-
"""
tests/test_pfcg_composta_sync_service.py
======================================================================
Bateria de Testes Unitários Aprofundada para a Fase 4
(Sincronização PFCG_COMPOSTA com RFC exclusivo, preflight, delta e readback)
======================================================================
"""

import openpyxl
import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

from sap_rfc.pfcg_composta_sync_service import (
    obter_pendencias_pfcg_composta,
    validar_roles_filhas_nao_compostas,
    ler_relacoes_agr_agrs,
    sincronizar_pfcg_composta,
    COL_STATUS, COL_MSG, COL_TIMESTEMP, COL_PRD
)


@pytest.fixture
def excel_fake_pfcg_composta(tmp_path):
    """Cria uma planilha Excel temporária simulando a sheet PFCG_COMPOSTA."""
    excel_path = tmp_path / "test_pfcg_composta.xlsx"
    wb = openpyxl.Workbook()
    
    # Sheet default
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    
    # Cabeçalho
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    
    # Linhas de teste
    # Linha 2: Pendente (STATUS vazio)
    ws.append([742, "Z_BR_PURCHREQ_MANAGER", "PURCHASE REQUISITION MANAGER", "Z_FI_VENDOR_READ", "", "", "", ""])
    # Linha 3: Pendente (STATUS vazio)
    ws.append([743, "Z_BR_PURCHREQ_MANAGER", "PURCHASE REQUISITION MANAGER", "Z_MM_PO_CREATE", "", "", "", ""])
    # Linha 4: Já Concluído (STATUS preenchido)
    ws.append([700, "Z_OLD_ROLE", "OLD DESC", "Z_SINGLE_OLD", "Concluído", "Validado em SAP QAD e PRD via RFC", "2026-10-09 10:00:00", "Validado"])
    
    wb.save(excel_path)
    wb.close()
    return str(excel_path)


def test_obter_pendencias_pfcg_composta(excel_fake_pfcg_composta):
    pendencias = obter_pendencias_pfcg_composta(excel_fake_pfcg_composta)
    assert len(pendencias) == 2
    assert pendencias[0]["id"] == 742
    assert pendencias[0]["composite"] == "Z_BR_PURCHREQ_MANAGER"
    assert pendencias[0]["single"] == "Z_FI_VENDOR_READ"
    assert pendencias[1]["id"] == 743
    assert pendencias[1]["single"] == "Z_MM_PO_CREATE"


def test_validar_roles_filhas_nao_compostas_sucesso():
    mock_conn = MagicMock()
    mock_guard = MagicMock()
    with patch("sap_rfc.pfcg_composta_sync_service.read_table", return_value=[]):
        validas, invalidas = validar_roles_filhas_nao_compostas(mock_conn, mock_guard, ["Z_FI_SINGLE_1", "Z_MM_SINGLE_2"])
        assert validas == {"Z_FI_SINGLE_1", "Z_MM_SINGLE_2"}
        assert len(invalidas) == 0


def test_validar_roles_filhas_nao_compostas_com_composite_invalida():
    mock_conn = MagicMock()
    mock_guard = MagicMock()
    mock_data = [["Z_COMPOSITE_FILHA", "COLL_AGR", "X"]]
    with patch("sap_rfc.pfcg_composta_sync_service.read_table", return_value=mock_data):
        validas, invalidas = validar_roles_filhas_nao_compostas(mock_conn, mock_guard, ["Z_FI_SINGLE_1", "Z_COMPOSITE_FILHA"])
        assert validas == {"Z_FI_SINGLE_1"}
        assert invalidas == {"Z_COMPOSITE_FILHA"}


def test_validar_roles_filhas_nao_compostas_varias_e_duplicadas():
    """Valida chunks, duplicados e lista vazia em validar_roles_filhas_nao_compostas."""
    mock_conn = MagicMock()
    mock_guard = MagicMock()
    
    # 1. Lista vazia
    v_vazia, inv_vazia = validar_roles_filhas_nao_compostas(mock_conn, mock_guard, [])
    assert len(v_vazia) == 0 and len(inv_vazia) == 0

    # 2. Várias roles com duplicados
    roles = [f"Z_ROLE_{i}" for i in range(35)] + ["Z_ROLE_0", "Z_ROLE_1"]
    with patch("sap_rfc.pfcg_composta_sync_service.read_table", return_value=[["Z_ROLE_5", "COLL_AGR", "X"]]):
        validas, invalidas = validar_roles_filhas_nao_compostas(mock_conn, mock_guard, roles)
        assert len(validas) == 34
        assert invalidas == {"Z_ROLE_5"}


def test_ler_relacoes_agr_agrs_uma_composite():
    mock_conn = MagicMock()
    mock_guard = MagicMock()
    mock_data = [
        ["Z_BR_PURCHREQ_MANAGER", "Z_FI_VENDOR_READ"],
        ["Z_BR_PURCHREQ_MANAGER", "Z_MM_PO_CREATE"]
    ]
    with patch("sap_rfc.pfcg_composta_sync_service.read_table", return_value=mock_data) as mock_read:
        rel, ok, err = ler_relacoes_agr_agrs(mock_conn, mock_guard, ["Z_BR_PURCHREQ_MANAGER"])
        assert ok is True
        assert err is None
        assert rel["Z_BR_PURCHREQ_MANAGER"] == {"Z_FI_VENDOR_READ", "Z_MM_PO_CREATE"}
        # Garantir que options geradas usaram make_option_in válido
        assert mock_read.call_args[1]["options"] == [{"TEXT": "AGR_NAME = 'Z_BR_PURCHREQ_MANAGER'"}]


def test_ler_relacoes_agr_agrs_varias_e_duplicados():
    """Testa várias Composite Roles, eliminação de duplicados e chunks com make_option_in."""
    mock_conn = MagicMock()
    mock_guard = MagicMock()
    
    # Vazia
    rel_vazia, ok_vazia, _ = ler_relacoes_agr_agrs(mock_conn, mock_guard, [])
    assert rel_vazia == {} and ok_vazia is True

    # 25 composites com repetições
    comps = [f"ZC_ROLE_{i}" for i in range(25)] + ["ZC_ROLE_0", "ZC_ROLE_1"]
    mock_data = [["ZC_ROLE_0", "Z_SINGLE_A"], ["ZC_ROLE_1", "Z_SINGLE_B"]]
    with patch("sap_rfc.pfcg_composta_sync_service.read_table", return_value=mock_data) as mock_read:
        rel, ok, err = ler_relacoes_agr_agrs(mock_conn, mock_guard, comps)
        assert ok is True
        assert err is None
        assert len(rel) == 25
        # Dividido em 2 chunks (20 + 5)
        assert mock_read.call_count == 2
        # Verificar sintaxe do 1º chunk (sempre com OR a partir do 2º)
        options_c1 = mock_read.call_args_list[0][1]["options"]
        assert options_c1[0]["TEXT"].startswith("AGR_NAME = 'ZC_ROLE_0'")
        assert options_c1[1]["TEXT"].startswith("OR AGR_NAME = 'ZC_ROLE_1'")


def test_sincronizar_pfcg_composta_preflight_falha_ambos(excel_fake_pfcg_composta):
    """Testa que se o preflight PRD falhar, a execução aborta de imediato (PREFLIGHT_FAILED)."""
    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(False, "Ambiente Offline")):
        res = sincronizar_pfcg_composta(excel_fake_pfcg_composta, dry_run=True)
        assert res["ok"] is False
        assert res["status"] == "PREFLIGHT_FAILED"


def test_sincronizar_pfcg_composta_dry_run_sucesso(excel_fake_pfcg_composta):
    """Testa execução dry-run com preflight PRD OK e readback positivo."""
    mock_conn = MagicMock()
    fake_params = {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"}

    mock_roles_info = {
        "Z_BR_PURCHREQ_MANAGER": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        "Z_FI_VENDOR_READ": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "Z_MM_PO_CREATE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
    }

    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(True, None)), \
         patch("sap_rfc.pfcg_composta_sync_service.build_connection_params_for_env", return_value=fake_params), \
         patch("sap_rfc.pfcg_composta_sync_service.Connection", return_value=mock_conn), \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap", return_value=mock_roles_info), \
         patch("sap_rfc.pfcg_composta_sync_service.validar_roles_filhas_nao_compostas", return_value=({"Z_FI_VENDOR_READ", "Z_MM_PO_CREATE"}, set())), \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", side_effect=[
             ({"Z_BR_PURCHREQ_MANAGER": set()}, True, None),
             ({"Z_BR_PURCHREQ_MANAGER": {"Z_FI_VENDOR_READ", "Z_MM_PO_CREATE"}}, True, None)
         ]):

        res = sincronizar_pfcg_composta(excel_fake_pfcg_composta, dry_run=True)

        assert res["ok"] is True
        assert res["status"] == "SUCESSO"
        assert res["total_pendencias"] == 2
        assert res["confirmadas"] == 2

        # Verificar se o Excel PERMANECEU INALTERADO no modo dry-run
        wb = openpyxl.load_workbook(excel_fake_pfcg_composta, data_only=True)
        ws = wb["PFCG_COMPOSTA"]
        assert ws.cell(row=2, column=COL_STATUS).value is None or ws.cell(row=2, column=COL_STATUS).value == ""
        wb.close()


def test_sincronizar_pfcg_composta_sem_pendencias(tmp_path):
    excel_path = tmp_path / "test_empty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    ws.append([700, "Z_ROLE", "TEXT", "Z_SINGLE", "Concluído", "MSG", "TS", "Validado"])
    wb.save(excel_path)
    wb.close()

    res = sincronizar_pfcg_composta(str(excel_path), dry_run=True)
    assert res["ok"] is True
    assert res["status"] == "SEM_PENDENCIAS"
    assert res["total_pendencias"] == 0


def test_sincronizar_pfcg_composta_classificacao_historicas(tmp_path):
    """Valida a separação relacional de Novo Lote (IDs 742-849) e Pendências Históricas (IDs < 742)."""
    excel_path = tmp_path / "test_historico.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    ws.append([739, "Z_BR_STOREMAINT", "DESC", "Z_SINGLE_1", "", "", "", ""])
    ws.append([742, "Z_BR_PURCHREQ", "DESC", "Z_SINGLE_2", "", "", "", ""])
    wb.save(excel_path)
    wb.close()

    pend = obter_pendencias_pfcg_composta(str(excel_path))
    assert len(pend) == 2
    assert pend[0]["id"] == 739
    assert pend[0]["categoria"] == "PENDENCIA_HISTORICA"
    assert pend[1]["id"] == 742
    assert pend[1]["categoria"] == "NOVO_LOTE"


def test_prevalidar_pfcg_composta_fluxo_completo(tmp_path):
    """
    Testa o fluxo completo de pré-validação com separação estrita:
    - candidatas
    - ja_existiam em PRD (não entra no delta executável)
    - ausentes em PRD
    - invalidas (Composite -> Composite excluídas)
    - executaveis (apenas ausente + Single válida)
    - duplicados na planilha não aumentam lote executável
    """
    from sap_rfc.pfcg_composta_sync_service import prevalidar_pfcg_composta

    excel_path = tmp_path / "test_preval.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    
    # 1. Já existe em PRD
    ws.append([1, "ZC_COMP1", "DESC", "ZS_EXISTE", "", "", "", ""])
    # 2. Ausente e Single Válida (deve ser executável)
    ws.append([2, "ZC_COMP1", "DESC", "ZS_NOVA", "", "", "", ""])
    # 3. Ausente mas Filha é Composite (inválida)
    ws.append([3, "ZC_COMP1", "DESC", "ZC_FILHA_COMPOSITE", "", "", "", ""])
    # 4. Duplicada da linha 2 (não deve inflar executaveis)
    ws.append([4, "ZC_COMP1", "DESC", "ZS_NOVA", "", "", "", ""])
    # 5. Já concluída (não entra como pendente)
    ws.append([5, "ZC_COMP1", "DESC", "ZS_OLD", "Concluído", "MSG", "TS", "Validado"])

    wb.save(excel_path)
    wb.close()

    mock_conn = MagicMock()
    fake_params = {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"}

    mock_roles_info = {
        "ZC_COMP1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        "ZS_EXISTE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "ZS_NOVA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "ZC_FILHA_COMPOSITE": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
    }

    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(True, None)), \
         patch("sap_rfc.pfcg_composta_sync_service.build_connection_params_for_env", return_value=fake_params), \
         patch("sap_rfc.pfcg_composta_sync_service.Connection", return_value=mock_conn), \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap", return_value=mock_roles_info), \
         patch("sap_rfc.pfcg_composta_sync_service.validar_roles_filhas_nao_compostas", return_value=(
             {"ZS_EXISTE", "ZS_NOVA"}, {"ZC_FILHA_COMPOSITE"}
         )), \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", return_value=(
             {"ZC_COMP1": {"ZS_EXISTE"}}, True, None
         )):

        res = prevalidar_pfcg_composta(str(excel_path), ambiente_prd="PRD")

        assert res["ok"] is True
        assert res["analisadas"] == 4
        assert res["ja_existiam"] == 1
        assert res["ausentes"] == 3
        assert res["invalidas_composite_composite"] == 1
        assert res["executaveis"] == 1  # 2 ausentes de ZS_NOVA desduplicadas para 1
        assert len(res["invalidas"]) == 1
        assert res["invalidas"][0]["composite"] == "ZC_COMP1"
        assert res["invalidas"][0]["single"] == "ZC_FILHA_COMPOSITE"
        assert res["invalidas"][0]["tipo_child"] == "COMPOSITE"


def test_bloqueio_seguranca_divergencia_contagem_aborta_sem_escrita(tmp_path):
    """
    Testa que se houver divergência entre a quantidade confirmada/esperada
    e o lote de executáveis calculado pelo serviço, a operação ABORTA sem escrita.
    """
    excel_path = tmp_path / "test_bloqueio.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    ws.append([1, "ZC_COMP1", "DESC", "ZS_NOVA", "", "", "", ""])
    wb.save(excel_path)
    wb.close()

    mock_conn = MagicMock()
    fake_params = {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"}

    mock_roles_info = {
        "ZC_COMP1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        "ZS_NOVA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
    }

    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(True, None)), \
         patch("sap_rfc.pfcg_composta_sync_service.build_connection_params_for_env", return_value=fake_params), \
         patch("sap_rfc.pfcg_composta_sync_service.Connection", return_value=mock_conn), \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap", return_value=mock_roles_info), \
         patch("sap_rfc.pfcg_composta_sync_service.validar_roles_filhas_nao_compostas", return_value=({"ZS_NOVA"}, set())), \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", return_value=({"ZC_COMP1": set()}, True, None)):

        # Passamos qtd_esperada_executaveis = 999 (diverge de 1)
        res = sincronizar_pfcg_composta(
            str(excel_path),
            dry_run=False,
            only_prd=True,
            qtd_esperada_executaveis=999
        )

        assert res["ok"] is False
        assert res["status"] == "BLOQUEIO_SEGURANCA_DIVERGENCIA"
        assert "Inconsistência entre pré-validação e lote de execução" in res["mensagem"]
        # Assegurar que nenhuma chamada RFC de escrita foi feita (PRGN_RFC_ADD_AGRS_TO_COLL_AGR)
        chamadas_escrita = [c for c in mock_conn.call.call_args_list if c[0][0] == "PRGN_RFC_ADD_AGRS_TO_COLL_AGR"]
        assert len(chamadas_escrita) == 0


def test_execucao_apenas_delta_valido_e_readback_confirmado(tmp_path):
    """
    Testa que apenas REALMENTE_PENDENTE + SINGLE válida entra no payload de escrita,
    e o readback confirma fisicamente apenas o que foi gravado.
    """
    excel_path = tmp_path / "test_delta.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    ws.append([1, "ZC_COMP1", "DESC", "ZS_EXISTE", "", "", "", ""])
    ws.append([2, "ZC_COMP1", "DESC", "ZS_NOVA", "", "", "", ""])
    ws.append([3, "ZC_COMP1", "DESC", "ZC_INVALIDA", "", "", "", ""])
    wb.save(excel_path)
    wb.close()

    mock_conn = MagicMock()
    fake_params = {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"}

    mock_roles_info = {
        "ZC_COMP1": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        "ZS_EXISTE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "ZS_NOVA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "ZC_INVALIDA": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
    }

    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(True, None)), \
         patch("sap_rfc.pfcg_composta_sync_service.build_connection_params_for_env", return_value=fake_params), \
         patch("sap_rfc.pfcg_composta_sync_service.Connection", return_value=mock_conn), \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap", return_value=mock_roles_info), \
         patch("sap_rfc.pfcg_composta_sync_service.validar_roles_filhas_nao_compostas", return_value=(
             {"ZS_EXISTE", "ZS_NOVA"}, {"ZC_INVALIDA"}
         )), \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", side_effect=[
             # 1. Chamada na prevalidacao
             ({"ZC_COMP1": {"ZS_EXISTE"}}, True, None),
             # 2. Chamada no readback pós-escrita (agora contém ZS_NOVA também)
             ({"ZC_COMP1": {"ZS_EXISTE", "ZS_NOVA"}}, True, None)
         ]):

        res = sincronizar_pfcg_composta(
            str(excel_path),
            dry_run=False,
            only_prd=True,
            qtd_esperada_executaveis=1
        )

        assert res["ok"] is True
        assert res["analisadas"] == 3
        assert res["ja_existiam"] == 1
        assert res["ausentes"] == 2
        assert res["invalidas_composite_composite"] == 1
        assert res["executaveis"] == 1
        assert res["criadas"] == 1
        assert res["confirmadas"] == 1

        # Verificar chamada RFC: APENAS ZS_NOVA foi enviada
        chamadas_rfc = [call for call in mock_conn.call.call_args_list if call[0][0] == "PRGN_RFC_ADD_AGRS_TO_COLL_AGR"]
        assert len(chamadas_rfc) == 1
        kwargs_chamada = chamadas_rfc[0][1]
        assert kwargs_chamada["ACTIVITY_GROUP"] == "ZC_COMP1"
        assert kwargs_chamada["ACTIVITY_GROUPS"] == [{"AGR_NAME": "ZS_NOVA", "TEXT": "DESC"}]

        # Verificar gravação no Excel: apenas a linha 3 (ZS_NOVA, row_idx=3) virou 'Concluído'
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb["PFCG_COMPOSTA"]
        assert ws.cell(row=2, column=COL_STATUS).value in (None, "")   # Linha 2 (já existia) não alterada
        assert ws.cell(row=3, column=COL_STATUS).value == "Concluído" # Linha 3 (criada e confirmada)
        assert ws.cell(row=4, column=COL_STATUS).value in (None, "")   # Linha 4 (inválida) não alterada
        wb.close()


def test_reconciliar_ja_existentes_pfcg_composta_excel_completo(tmp_path):
    """
    Testa a reconciliação de linhas já existentes no Excel:
    1. STATUS atualizado para 'Concluído'
    2. MSG atualizado
    3. TIMESTEMP preenchido
    4. PRD = 'Validado'
    5. Linhas REALMENTE_PENDENTE e INVALIDA_COMPOSITE_COMPOSITE não são alteradas
    6. Backup criado apenas quando há atualização
    7. Segunda execução de prevalidar não volta a contar as reconciliadas
    """
    from sap_rfc.pfcg_composta_sync_service import (
        reconciliar_ja_existentes_pfcg_composta_excel,
        prevalidar_pfcg_composta,
        obter_pendencias_pfcg_composta
    )

    excel_path = tmp_path / "test_reconciliacao.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    
    # Linha 2: Já existe em PRD (deve ser reconciliada)
    ws.append([1, "ZC_COMP1", "DESC", "ZS_EXISTE", "Pendente QAD", "", "", ""])
    # Linha 3: Ausente em PRD (deve permanecer inalterada)
    ws.append([2, "ZC_COMP1", "DESC", "ZS_PENDENTE", "", "", "", ""])
    # Linha 4: Inválida Composite (deve permanecer inalterada)
    ws.append([3, "ZC_COMP1", "DESC", "ZC_INVALIDA", "", "", "", ""])
    wb.save(excel_path)
    wb.close()

    itens_ja_existem = [{"row_idx": 2, "composite": "ZC_COMP1", "single": "ZS_EXISTE"}]

    # Reconciliar no Excel
    res_rec = reconciliar_ja_existentes_pfcg_composta_excel(str(excel_path), itens_ja_existem, criar_backup=True)

    assert res_rec["ok"] is True
    assert res_rec["linhas_reconciliadas"] == 1
    assert res_rec["backup_criado"] is not None

    # Verificar células da linha 2 reconciliada
    wb_pos = openpyxl.load_workbook(excel_path, data_only=True)
    ws_pos = wb_pos["PFCG_COMPOSTA"]
    assert ws_pos.cell(row=2, column=COL_STATUS).value == "Concluído"
    assert "Readback OK" in str(ws_pos.cell(row=2, column=COL_MSG).value)
    assert ws_pos.cell(row=2, column=COL_TIMESTEMP).value is not None
    assert ws_pos.cell(row=2, column=COL_PRD).value == "Validado"

    # Linhas 3 e 4 permanecem vazias / intocadas
    assert ws_pos.cell(row=3, column=COL_STATUS).value in (None, "")
    assert ws_pos.cell(row=4, column=COL_STATUS).value in (None, "")
    wb_pos.close()

    # Segunda leitura de pendências: a linha 2 não deve mais constar
    pendencias_segunda = obter_pendencias_pfcg_composta(str(excel_path))
    assert len(pendencias_segunda) == 2
    ids_pend = [p["id"] for p in pendencias_segunda]
    assert 1 not in ids_pend
    assert 2 in ids_pend
    assert 3 in ids_pend


def test_prevalidar_bloqueia_parent_inexistente_e_child_inexistente(tmp_path):
    """
    Testa que relações com:
    - PARENT_NAO_EXISTE são bloqueadas e agrupadas em pais_inexistentes
    - CHILD_NAO_EXISTE são bloqueadas e agrupadas em filhos_inexistentes
    - PARENT_TIPO_DIVERGENTE são bloqueadas
    Nenhuma delas entra em itens_executaveis.
    """
    from sap_rfc.pfcg_composta_sync_service import prevalidar_pfcg_composta

    excel_path = tmp_path / "test_inexistentes.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PFCG_COMPOSTA"
    ws.append(["ID", "AGR_NAME_COMPOSTA", "TEXT", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP", "PRD"])
    
    # 1. Pai não existe
    ws.append([1, "ZC_PAI_INEXISTENTE", "DESC", "ZS_VALIDA", "", "", "", ""])
    # 2. Filho não existe
    ws.append([2, "ZC_PAI_OK", "DESC", "ZS_FILHO_INEXISTENTE", "", "", "", ""])
    # 3. Pai existe mas é Single (tipo divergente)
    ws.append([3, "ZC_PAI_SINGLE", "DESC", "ZS_VALIDA", "", "", "", ""])
    # 4. Relação perfeita (Pai Composite OK, Filho Single OK, ausente de AGR_AGRS)
    ws.append([4, "ZC_PAI_OK", "DESC", "ZS_VALIDA", "", "", "", ""])

    wb.save(excel_path)
    wb.close()

    mock_roles_info = {
        "ZC_PAI_INEXISTENTE": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
        "ZC_PAI_OK": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        "ZC_PAI_SINGLE": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
        "ZS_FILHO_INEXISTENTE": {"existe": False, "tipo_sap": "INEXISTENTE", "erro": None},
        "ZS_VALIDA": {"existe": True, "tipo_sap": "SINGLE", "erro": None},
    }

    mock_conn = MagicMock()
    fake_params = {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"}

    with patch("sap_rfc.pfcg_composta_sync_service.preflight_test_connection", return_value=(True, None)), \
         patch("sap_rfc.pfcg_composta_sync_service.build_connection_params_for_env", return_value=fake_params), \
         patch("sap_rfc.pfcg_composta_sync_service.Connection", return_value=mock_conn), \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap", return_value=mock_roles_info), \
         patch("sap_rfc.pfcg_composta_sync_service.validar_roles_filhas_nao_compostas", return_value=({"ZS_VALIDA"}, set())), \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", return_value=({
             "ZC_PAI_INEXISTENTE": set(), "ZC_PAI_OK": set(), "ZC_PAI_SINGLE": set()
         }, True, None)):

        res = prevalidar_pfcg_composta(str(excel_path), ambiente_prd="PRD")

        assert res["ok"] is True
        assert res["analisadas"] == 4
        assert res["parent_nao_existe"] == 1
        assert res["child_nao_existe"] == 1
        assert res["parent_tipo_divergente"] == 1
        assert res["executaveis"] == 1
        assert len(res["itens_executaveis"]) == 1
        assert res["itens_executaveis"][0]["composite"] == "ZC_PAI_OK"
        assert res["itens_executaveis"][0]["single"] == "ZS_VALIDA"

        # Validar agrupamentos
        assert "ZC_PAI_INEXISTENTE" in res["pais_inexistentes"]
        assert "ZS_FILHO_INEXISTENTE" in res["filhos_inexistentes"]


