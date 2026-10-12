import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
import sys
import importlib.util

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Carregar dinamica/seguramente o modulo Projeto Perfil.py com espaco no nome
script_path = PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py"
spec = importlib.util.spec_from_file_location("projeto_perfil_mod", script_path)
pp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pp)


def test_despachante_pendencias_tres_processos_ssn(tmp_path):
    """Testa o despachante quando há 3 processos e o utilizador responde S, S, N."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    mock_preval = MagicMock(return_value={
        "ok": True,
        "analisadas": 372,
        "ja_existiam": 59,
        "ausentes": 313,
        "invalidas_composite_composite": 15,
        "executaveis": 298,
        "erros": [],
        "invalidas": [{"composite": "Z_C1", "single": "Z_C2"}],
        "itens_executaveis": []
    })

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 372, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch.object(pp, "auditar_prevalidacao_cua_remove_prd", return_value={"ok": True, "validas": 126, "invalidas": 360, "ja_nao_existem": 102}), \
         patch("builtins.input", side_effect=["S", "S", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_rec = MagicMock(return_value={"ok": True, "linhas_reconciliadas": 59, "backup_criado": "backup.xlsx"})
        mock_cua_add = MagicMock(return_value={"ok": True, "candidatas": 348, "ja_existentes": 282, "tentadas": 66, "confirmadas_readback": 66, "falhadas": 0, "pendentes": 0, "status": "SUCESSO", "detalhe": "66 confirmadas por readback"})
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "prevalidar_pfcg_composta": mock_preval,
            "reconciliar_ja_existentes_pfcg_composta_excel": mock_rec,
            "executar_atribuicoes_cua_sap": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        # Asserções
        assert resumo["PFCG_COMPOSTA"]["executado"] is True
        assert resumo["PFCG_COMPOSTA"]["resultado"] == "SUCESSO"
        mock_rec.assert_called_once()
        mock_pfcg_comp.executar.assert_called_once()

        assert resumo["CUA_ADICIONAR"]["executado"] is True
        assert resumo["CUA_ADICIONAR"]["resultado"] == "SUCESSO"
        mock_cua_add.assert_called_once()

        assert resumo["CUA_REMOVE"]["executado"] is False
        assert resumo["CUA_REMOVE"]["resultado"] == "IGNORADO_PELO_UTILIZADOR"
        mock_cua_rm.executar.assert_not_called()


def test_despachante_pendencias_tres_processos_nsn(tmp_path):
    """Testa quando o utilizador recusa o 1º processo, aceita o 2º e recusa o 3º (N, S, N)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    mock_preval = MagicMock(return_value={
        "ok": True,
        "analisadas": 372,
        "ja_existiam": 59,
        "ausentes": 313,
        "invalidas_composite_composite": 15,
        "executaveis": 298,
        "erros": [],
        "invalidas": [],
        "itens_executaveis": []
    })

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 372, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch("builtins.input", side_effect=["N", "S", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_rec = MagicMock(return_value={"ok": True, "linhas_reconciliadas": 59, "backup_criado": "backup.xlsx"})
        mock_cua_add = MagicMock(return_value={"ok": True, "candidatas": 348, "ja_existentes": 282, "tentadas": 66, "confirmadas_readback": 66, "falhadas": 0, "pendentes": 0, "status": "SUCESSO", "detalhe": "66 confirmadas por readback"})
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "prevalidar_pfcg_composta": mock_preval,
            "reconciliar_ja_existentes_pfcg_composta_excel": mock_rec,
            "executar_atribuicoes_cua_sap": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["PFCG_COMPOSTA"]["executado"] is False
        mock_rec.assert_called_once()
        mock_pfcg_comp.executar.assert_not_called()

        assert resumo["CUA_ADICIONAR"]["executado"] is True
        mock_cua_add.assert_called_once()

        assert resumo["CUA_REMOVE"]["executado"] is False
        mock_cua_rm.executar.assert_not_called()


def test_despachante_pendencias_recusa_total_nnn(tmp_path):
    """Testa a recusa de todos os processos (N, N, N)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    mock_preval = MagicMock(return_value={
        "ok": True,
        "analisadas": 372,
        "ja_existiam": 59,
        "ausentes": 313,
        "invalidas_composite_composite": 15,
        "executaveis": 298,
        "erros": [],
        "invalidas": [],
        "itens_executaveis": []
    })

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 372, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch("builtins.input", side_effect=["N", "N", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_rec = MagicMock(return_value={"ok": True, "linhas_reconciliadas": 59, "backup_criado": "backup.xlsx"})
        mock_cua_add = MagicMock()
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "prevalidar_pfcg_composta": mock_preval,
            "reconciliar_ja_existentes_pfcg_composta_excel": mock_rec,
            "executar_atribuicoes_cua_sap": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["PFCG_COMPOSTA"]["executado"] is False
        assert resumo["CUA_ADICIONAR"]["executado"] is False
        assert resumo["CUA_REMOVE"]["executado"] is False
        mock_rec.assert_called_once()
        mock_pfcg_comp.executar.assert_not_called()
        mock_cua_add.assert_not_called()
        mock_cua_rm.executar.assert_not_called()


def test_despachante_falha_reconciliacao_excel_aborta(tmp_path):
    """Testa que se a reconciliação do Excel falhar, aborta imediatamente sem executar SAP."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    mock_preval = MagicMock(return_value={
        "ok": True,
        "analisadas": 372,
        "ja_existiam": 147,
        "ausentes": 225,
        "invalidas_composite_composite": 15,
        "executaveis": 210,
        "erros": [],
        "invalidas": [],
        "itens_ja_existem": [{"row_idx": 2}]
    })

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 372}):
        mock_pfcg_comp = MagicMock()
        mock_rec = MagicMock(return_value={"ok": False, "erro": "Arquivo Excel bloqueado por outro processo"})

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "prevalidar_pfcg_composta": mock_preval,
            "reconciliar_ja_existentes_pfcg_composta_excel": mock_rec
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=True, mock_executores=mocks)

        assert resumo["PFCG_COMPOSTA"]["executado"] is False
        assert resumo["PFCG_COMPOSTA"]["resultado"] == "ERRO_EXCEL"
        mock_rec.assert_called_once()
        mock_pfcg_comp.executar.assert_not_called()


def test_despachante_dupla_confirmacao_cua_remove(tmp_path):
    """Testa a dupla confirmação necessária para a CUA_REMOVE (S na 1ª pergunta, N na 2ª pergunta)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_remove_prd", return_value={"ok": True, "validas": 126, "invalidas": 360, "ja_nao_existem": 102}), \
         patch("builtins.input", side_effect=["S", "N"]):

        mock_cua_rm = MagicMock()
        mocks = {"processo_cua_remove": mock_cua_rm}

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["CUA_REMOVE"]["executado"] is False
        assert resumo["CUA_REMOVE"]["resultado"] == "IGNORADO_PELO_UTILIZADOR"
        mock_cua_rm.executar.assert_not_called()


def test_despachante_dupla_confirmacao_cua_remove_confirmada(tmp_path):
    """Testa a dupla confirmação aceita para a CUA_REMOVE (S na 1ª pergunta, S na 2ª pergunta)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_remove_prd", return_value={"ok": True, "validas": 126, "invalidas": 360, "ja_nao_existem": 102}), \
         patch("builtins.input", side_effect=["S", "S"]):

        mock_cua_rm = MagicMock()
        mocks = {"processo_cua_remove": mock_cua_rm}

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["CUA_REMOVE"]["executado"] is True
        assert resumo["CUA_REMOVE"]["resultado"] == "SUCESSO"
        mock_cua_rm.executar.assert_called_once()


def test_cua_adicionar_executor_indisponivel_zero_executadas(tmp_path):
    """Testa que se o executor de CUA_ADICIONAR falha ou está indisponível, retorna 0 executadas e resultado != SUCESSO."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"CUA_ADICIONAR": 66}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 0, "realmente_pendentes": 66}), \
         patch("builtins.input", return_value="S"):

        mock_cua_fail = MagicMock(return_value={
            "ok": False,
            "candidatas": 66,
            "ja_existentes": 0,
            "tentadas": 0,
            "confirmadas_readback": 0,
            "falhadas": 0,
            "pendentes": 66,
            "status": "PENDENTE_AMBIENTE_INDISPONIVEL",
            "detalhe": "Nenhum ambiente SAP conectado"
        })

        mocks = {"executar_atribuicoes_cua_sap": mock_cua_fail}
        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["CUA_ADICIONAR"]["executado"] is False
        assert resumo["CUA_ADICIONAR"]["resultado"] == "PENDENTE_AMBIENTE_INDISPONIVEL"
        assert resumo["CUA_ADICIONAR"]["resultado"] != "SUCESSO"
        assert "0 confirmadas por readback" in resumo["CUA_ADICIONAR"]["detalhe"]
        assert "66 pendentes" in resumo["CUA_ADICIONAR"]["detalhe"]


def test_cua_adicionar_readback_negativo_nao_gera_sucesso(tmp_path):
    """Testa que se a BAPI/COMMIT rodar mas o readback falhar (role não confirmada fisicamente), não gera SUCESSO."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"CUA_ADICIONAR": 66}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 0, "realmente_pendentes": 66}), \
         patch("builtins.input", return_value="S"):

        mock_cua_rb_fail = MagicMock(return_value={
            "ok": False,
            "candidatas": 66,
            "ja_existentes": 0,
            "tentadas": 66,
            "confirmadas_readback": 0,
            "falhadas": 66,
            "pendentes": 66,
            "status": "FALHA",
            "detalhe": "0 confirmadas por readback / 0 já existentes / 66 falhas"
        })

        mocks = {"executar_atribuicoes_cua_sap": mock_cua_rb_fail}
        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["CUA_ADICIONAR"]["executado"] is True
        assert resumo["CUA_ADICIONAR"]["resultado"] == "FALHA"
        assert resumo["CUA_ADICIONAR"]["resultado"] != "SUCESSO"
        assert "0 confirmadas por readback" in resumo["CUA_ADICIONAR"]["detalhe"]


def test_despachante_resumo_usa_confirmadas_e_nao_candidatas(tmp_path):
    """Testa que o resumo final utiliza confirmadas_readback e nunca a quantidade de candidatas confirmadas pelo user."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"CUA_ADICIONAR": 66}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch("builtins.input", return_value="S"):

        # 66 pendentes, mas apenas 40 confirmadas por readback e 26 falhas
        mock_cua_parcial = MagicMock(return_value={
            "ok": False,
            "candidatas": 348,
            "ja_existentes": 282,
            "tentadas": 66,
            "confirmadas_readback": 40,
            "falhadas": 26,
            "pendentes": 26,
            "status": "FALHA",
            "detalhe": "40 confirmadas por readback / 282 já existentes / 26 falhas"
        })

        mocks = {"executar_atribuicoes_cua_sap": mock_cua_parcial}
        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["CUA_ADICIONAR"]["resultado"] == "FALHA"
        assert "40 confirmadas por readback" in resumo["CUA_ADICIONAR"]["detalhe"]
        assert "26 falhas" in resumo["CUA_ADICIONAR"]["detalhe"]
        assert "66 executadas" not in resumo["CUA_ADICIONAR"]["detalhe"]


def test_rfc_logon_failure_aborta_automaticamente_etapas_sap(tmp_path):
    """Cenário 19: RFC_LOGON_FAILURE / Password must be changed no PRD aborta todas as etapas SAP dependentes sem perguntar."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    mock_val_roles = MagicMock(return_value={
        "ok": False,
        "ambientes": {
            "PRD": {
                "online": False,
                "status": "INDISPONIVEL",
                "validado": False,
                "erro_conexao": "LogonError: Password must be changed"
            }
        },
        "roles_confirmadas": {"PRD": set(), "QAD": set()}
    })

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_CREATE": 10, "PFCG_COMPOSTA": 20, "CUA_ADICIONAR": 30}), \
         patch("builtins.input") as mock_input:

        mocks = {"executar_fase_validacao_existencia_roles": mock_val_roles}
        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        # Não deve fazer qualquer pergunta ao utilizador
        mock_input.assert_not_called()

        # Todas as etapas devem estar abortadas com resultado indicando falha de logon
        assert "PFCG_CREATE" in resumo
        assert resumo["PFCG_CREATE"]["resultado"] == "ABORTADO_RFC_LOGON_FAILURE"
        assert resumo["PFCG_COMPOSTA"]["resultado"] == "ABORTADO_RFC_LOGON_FAILURE"
        assert resumo["CUA_ADICIONAR"]["resultado"] == "ABORTADO_RFC_LOGON_FAILURE"


def test_child_tipo_divergente_bloqueada_em_pfcg_composta(tmp_path):
    """Cenários 20 e 21: Child esperada como SINGLE mas no SAP é COMPOSITE é bloqueada em PFCG_COMPOSTA com CHILD_TIPO_DIVERGENTE."""
    from sap_rfc.pfcg_composta_sync_service import prevalidar_pfcg_composta

    caminho_excel = str(tmp_path / "fake_comp.xlsx")

    with patch("sap_rfc.pfcg_composta_sync_service.obter_pendencias_pfcg_composta") as mock_pend, \
         patch("sap_rfc.pfcg_composta_sync_service.verificar_existencia_e_tipo_sap") as mock_sap, \
         patch("sap_rfc.pfcg_composta_sync_service.ler_relacoes_agr_agrs", return_value=({}, True, None)):

        mock_pend.return_value = [
            {"row_idx": 2, "composite": "Z_PARENT_COMP", "single": "Z_CHILD_ERR_COMP", "status": ""}
        ]

        # Parent é COMPOSITE OK, Child existe fisicamente no SAP como COMPOSITE (divergente!)
        mock_sap.return_value = {
            "Z_PARENT_COMP": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
            "Z_CHILD_ERR_COMP": {"existe": True, "tipo_sap": "COMPOSITE", "erro": None},
        }

        res = prevalidar_pfcg_composta(caminho_excel, ambiente_prd="PRD")

        assert res["ok"] is True
        assert res["child_tipo_divergente"] == 1
        assert res["executaveis"] == 0
        assert len(res["itens_bloqueados"]) == 1
        assert res["itens_bloqueados"][0]["estado_prd"] == "CHILD_TIPO_DIVERGENTE"


