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

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 231, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch.object(pp, "auditar_prevalidacao_cua_remove_prd", return_value={"ok": True, "validas": 126, "invalidas": 360, "ja_nao_existem": 102}), \
         patch("builtins.input", side_effect=["S", "S", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_cua_add = MagicMock()
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "processo_cua_adicionar": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        # Asserções
        assert resumo["PFCG_COMPOSTA"]["executado"] is True
        assert resumo["PFCG_COMPOSTA"]["resultado"] == "SUCESSO"
        mock_pfcg_comp.executar.assert_called_once()

        assert resumo["CUA_ADICIONAR"]["executado"] is True
        assert resumo["CUA_ADICIONAR"]["resultado"] == "SUCESSO"
        mock_cua_add.executar.assert_called_once()

        assert resumo["CUA_REMOVE"]["executado"] is False
        assert resumo["CUA_REMOVE"]["resultado"] == "IGNORADO_PELO_UTILIZADOR"
        mock_cua_rm.executar.assert_not_called()


def test_despachante_pendencias_tres_processos_nsn(tmp_path):
    """Testa quando o utilizador recusa o 1º processo, aceita o 2º e recusa o 3º (N, S, N)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 231, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch("builtins.input", side_effect=["N", "S", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_cua_add = MagicMock()
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "processo_cua_adicionar": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["PFCG_COMPOSTA"]["executado"] is False
        mock_pfcg_comp.executar.assert_not_called()

        assert resumo["CUA_ADICIONAR"]["executado"] is True
        mock_cua_add.executar.assert_called_once()

        assert resumo["CUA_REMOVE"]["executado"] is False
        mock_cua_rm.executar.assert_not_called()


def test_despachante_pendencias_recusa_total_nnn(tmp_path):
    """Testa a recusa de todos os processos (N, N, N)."""
    caminho_excel = str(tmp_path / "fake.xlsx")

    with patch.object(pp, "obter_pendencias_status", return_value={"PFCG_COMPOSTA": 231, "CUA_ADICIONAR": 348, "CUA_REMOVE": 588}), \
         patch.object(pp, "auditar_prevalidacao_cua_adicionar_prd", return_value={"ok": True, "ja_existem": 282, "realmente_pendentes": 66}), \
         patch("builtins.input", side_effect=["N", "N", "N"]):

        mock_pfcg_comp = MagicMock()
        mock_cua_add = MagicMock()
        mock_cua_rm = MagicMock()

        mocks = {
            "processo_pfcg_composta": mock_pfcg_comp,
            "processo_cua_adicionar": mock_cua_add,
            "processo_cua_remove": mock_cua_rm
        }

        resumo = pp.verificar_e_perguntar_pendencias(caminho_excel, assumir_sim=False, mock_executores=mocks)

        assert resumo["PFCG_COMPOSTA"]["executado"] is False
        assert resumo["CUA_ADICIONAR"]["executado"] is False
        assert resumo["CUA_REMOVE"]["executado"] is False
        mock_pfcg_comp.executar.assert_not_called()
        mock_cua_add.executar.assert_not_called()
        mock_cua_rm.executar.assert_not_called()


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
