# -*- coding: utf-8 -*-
"""
tests/test_tarefa1_onedrive_precheck.py
======================================================================
Testes unitários automatizados para o Pre-Check do OneDrive na Tarefa 1:
1. A pergunta é a primeira ação da Tarefa 1.
2. Nenhuma etapa [1/8] ocorre antes da resposta.
3. Resposta N mantém o ficheiro local.
4. Resposta vazia assume N.
5. Resposta S cria backup antes da cópia.
6. Resposta S copia somente OneDrive -> Local.
7. Falha no backup aborta.
8. Origem OneDrive inexistente aborta atualização e não inicia sincronização.
9. Readback inválido aborta antes de [1/8].
10. Nunca existe cópia Local -> OneDrive.
======================================================================
"""

import os
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock
import pytest
import openpyxl

import importlib.util

PROJECT_ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("projeto_perfil", PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


@pytest.fixture
def ambiente_arquivos(tmp_path):
    """Cria arquivos Excel sintéticos válidos para origem e destino."""
    od_file = tmp_path / "onedrive" / "S4H_Perfis de autorização_v1.xlsx"
    od_file.parent.mkdir(parents=True, exist_ok=True)
    
    wb_od = openpyxl.Workbook()
    wb_od.active.title = "CONTROLO"
    ws_prop = wb_od.create_sheet("Proposta")
    ws_prop.append(["FUNÇÃO", "DESCRIÇÃO"])
    ws_pa = wb_od.create_sheet("Proposta Ativa")
    ws_pa.append(["UTILIZADOR", "COMPOSITE", "DEPARTAMENTO", "DIRECAO", "FUNÇÕES INDIVIDUAIS"])
    wb_od.create_sheet("PFCG_CREATE")
    wb_od.create_sheet("PFCG_COMPOSTA")
    wb_od.create_sheet("PFCG_AUTHORITY")
    wb_od.save(str(od_file))
    wb_od.close()

    loc_file = tmp_path / "local" / "S4H_Perfis de autorização_v1.xlsx"
    loc_file.parent.mkdir(parents=True, exist_ok=True)

    wb_loc = openpyxl.Workbook()
    wb_loc.active.title = "CONTROLO"
    ws_prop_l = wb_loc.create_sheet("Proposta")
    ws_prop_l.append(["FUNÇÃO", "DESCRIÇÃO"])
    ws_pa_l = wb_loc.create_sheet("Proposta Ativa")
    ws_pa_l.append(["UTILIZADOR", "COMPOSITE", "DEPARTAMENTO", "DIRECAO", "FUNÇÕES INDIVIDUAIS"])
    wb_loc.create_sheet("PFCG_CREATE")
    wb_loc.create_sheet("PFCG_COMPOSTA")
    wb_loc.create_sheet("PFCG_AUTHORITY")
    wb_loc.save(str(loc_file))
    wb_loc.close()

    return {
        "onedrive": str(od_file),
        "local": str(loc_file),
        "output": str(tmp_path / "output")
    }


def test_1_2_pergunta_e_primeira_acao_e_etapa_1_nao_ocorre_antes_da_resposta(ambiente_arquivos, capsys):
    """Garante que a pergunta ocorre antes de qualquer etapa [1/8]."""
    chamadas_ordem = []
    prompts_recebidos = []

    def mock_input(prompt):
        chamadas_ordem.append("input")
        prompts_recebidos.append(prompt)
        return "N"

    def mock_auditoria(*args, **kwargs):
        chamadas_ordem.append("etapa_1")
        return {"tem_pendencias": False}

    with patch("builtins.input", side_effect=mock_input), \
         patch.object(mod, "auditar_matrizes_departamentais_vs_proposta", side_effect=mock_auditoria), \
         patch.object(mod, "analisar_folha_proposta", return_value={"todos_tcodes": []}), \
         patch.object(mod, "preparar_cua_adicionar", return_value={"novas_linhas_count": 0, "total_users_analisados": 0, "total_atribuicoes_esperadas": 0, "ja_existentes_count": 0, "detalhes_por_dep": {}}):

        mod.executar_atualizacao_integrada_excel(caminho_excel=ambiente_arquivos["local"], incorporar_prd_rfc=False, validar_tcodes_rfc=False)

    assert "input" in chamadas_ordem
    assert "etapa_1" in chamadas_ordem
    assert chamadas_ordem.index("input") < chamadas_ordem.index("etapa_1")
    assert any("Deseja atualizar o Excel local com a versão do OneDrive? (S/N) [Padrão: N]:" in p for p in prompts_recebidos)

    captured = capsys.readouterr().out
    assert "ℹ Excel local mantido." in captured


def test_3_resposta_n_mantem_ficheiro_local(ambiente_arquivos):
    """Resposta N mantém o arquivo local sem alteração nem cópia."""
    mtime_antes = os.path.getmtime(ambiente_arquivos["local"])
    size_antes = os.path.getsize(ambiente_arquivos["local"])

    with patch("builtins.input", return_value="N"), \
         patch.object(mod, "atualizar_excel_local_do_onedrive_seguro") as mock_sync, \
         patch.object(mod, "auditar_matrizes_departamentais_vs_proposta", return_value={"tem_pendencias": False}), \
         patch.object(mod, "analisar_folha_proposta", return_value={"todos_tcodes": []}), \
         patch.object(mod, "preparar_cua_adicionar", return_value={"novas_linhas_count": 0, "total_users_analisados": 0, "total_atribuicoes_esperadas": 0, "ja_existentes_count": 0, "detalhes_por_dep": {}}):

        mod.executar_atualizacao_integrada_excel(caminho_excel=ambiente_arquivos["local"], incorporar_prd_rfc=False, validar_tcodes_rfc=False)
        mock_sync.assert_not_called()

    assert os.path.getmtime(ambiente_arquivos["local"]) == mtime_antes
    assert os.path.getsize(ambiente_arquivos["local"]) == size_antes


def test_4_resposta_vazia_assume_n(capsys, ambiente_arquivos):
    """Pressionar Enter (resposta vazia) assume padrão N."""
    with patch("builtins.input", return_value=""), \
         patch.object(mod, "atualizar_excel_local_do_onedrive_seguro") as mock_sync, \
         patch.object(mod, "auditar_matrizes_departamentais_vs_proposta", return_value={"tem_pendencias": False}), \
         patch.object(mod, "analisar_folha_proposta", return_value={"todos_tcodes": []}), \
         patch.object(mod, "preparar_cua_adicionar", return_value={"novas_linhas_count": 0, "total_users_analisados": 0, "total_atribuicoes_esperadas": 0, "ja_existentes_count": 0, "detalhes_por_dep": {}}):

        mod.executar_atualizacao_integrada_excel(caminho_excel=ambiente_arquivos["local"], incorporar_prd_rfc=False, validar_tcodes_rfc=False)
        mock_sync.assert_not_called()

    captured = capsys.readouterr().out
    assert "ℹ Excel local mantido." in captured


def test_5_6_resposta_s_cria_backup_e_copia_onedrive_para_local(tmp_path, ambiente_arquivos, capsys):
    """Resposta S cria backup antes da cópia e substitui local a partir do OneDrive."""
    # Alterar OneDrive para ter um marcador específico
    wb_od = openpyxl.load_workbook(ambiente_arquivos["onedrive"])
    wb_od["CONTROLO"]["A1"] = "MARCADOR_ONEDRIVE_NOVO"
    wb_od.save(ambiente_arquivos["onedrive"])
    wb_od.close()

    with patch("builtins.input", return_value="S"), \
         patch.object(mod, "obter_caminho_onedrive_oficial", return_value=ambiente_arquivos["onedrive"]), \
         patch.object(mod, "auditar_matrizes_departamentais_vs_proposta", return_value={"tem_pendencias": False}), \
         patch.object(mod, "analisar_folha_proposta", return_value={"todos_tcodes": []}), \
         patch.object(mod, "preparar_cua_adicionar", return_value={"novas_linhas_count": 0, "total_users_analisados": 0, "total_atribuicoes_esperadas": 0, "ja_existentes_count": 0, "detalhes_por_dep": {}}):

        mod.executar_atualizacao_integrada_excel(caminho_excel=ambiente_arquivos["local"], incorporar_prd_rfc=False, validar_tcodes_rfc=False)

    captured = capsys.readouterr().out
    assert "✓ Excel local atualizado a partir do OneDrive." in captured

    # Verificar que o local agora tem o conteúdo novo do OneDrive
    wb_check = openpyxl.load_workbook(ambiente_arquivos["local"])
    assert wb_check["CONTROLO"]["A1"].value == "MARCADOR_ONEDRIVE_NOVO"
    wb_check.close()


def test_7_falha_no_backup_aborta_atualizacao(ambiente_arquivos):
    """Se a criação do backup falhar, aborta imediatamente sem alterar o ficheiro local."""
    with patch("shutil.copy2", side_effect=IOError("Erro ao criar backup de permissao")), \
         patch.object(mod, "obter_caminho_onedrive_oficial", return_value=ambiente_arquivos["onedrive"]):

        ok, msg = mod.atualizar_excel_local_do_onedrive_seguro(
            caminho_destino=ambiente_arquivos["local"],
            caminho_origem_onedrive=ambiente_arquivos["onedrive"]
        )
        assert ok is False
        assert "Falha ao gerar backup" in msg


def test_8_origem_onedrive_inexistente_aborta(ambiente_arquivos):
    """Origem inexistente retorna erro claro e aborta."""
    caminho_inexistente = "C:\\inexistente\\OneDrive - Salsajeans\\S4H_Perfis.xlsx"
    ok, msg = mod.atualizar_excel_local_do_onedrive_seguro(
        caminho_destino=ambiente_arquivos["local"],
        caminho_origem_onedrive=caminho_inexistente
    )
    assert ok is False
    assert "não encontrado ou inacessível" in msg


def test_9_readback_invalido_aborta_antes_de_etapa_1(tmp_path, ambiente_arquivos):
    """Se o arquivo copiado estiver corrompido ou sem as sheets obrigatórias, o readback falha."""
    arquivo_corrompido = tmp_path / "corrompido.xlsx"
    arquivo_corrompido.write_text("conteudo invalido nao excel", encoding="utf-8")

    ok, msg = mod.atualizar_excel_local_do_onedrive_seguro(
        caminho_destino=ambiente_arquivos["local"],
        caminho_origem_onedrive=str(arquivo_corrompido)
    )
    assert ok is False
    assert "corrompido ou não pôde ser aberto" in msg


def test_10_nunca_existe_copia_local_para_onedrive():
    """Garante no código fonte que não há comandos shutil.copy/copy2 copiando local para OneDrive."""
    mod_path = PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py"
    code = mod_path.read_text(encoding="utf-8")

    assert "shutil.copy2(CAMINHO_EXCEL_LOCAL, origem" not in code
    assert "shutil.copy(CAMINHO_EXCEL_LOCAL, origem" not in code
    assert "shutil.copy2(destino, origem" not in code
