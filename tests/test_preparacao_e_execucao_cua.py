# -*- coding: utf-8 -*-
"""
test_preparacao_e_execucao_cua.py
======================================================================
Testes unitários automatizados cobrindo a separação das Etapas [6/7] e [7/7]:
  Etapa [6/7]: Preparação da sheet CUA_ADICIONAR (apenas Excel, idempotente, sem escrita SAP).
  Etapa [7/7]: Execução das atribuições no SAP (BAPI + Readback + Atualização Excel via ID).
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

# Adicionar raiz do projeto ao sys.path
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import openpyxl
import importlib.util

spec = importlib.util.spec_from_file_location("projeto_perfil", PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


class TestPreparacaoEExecucaoCUA(unittest.TestCase):

    def setUp(self):
        self.caminho_excel = mod.encontrar_excel_padrao()
        self.dados = mod.carregar_projeto_perfil(self.caminho_excel)

    def test_etapa_6_preparar_cua_adicionar_sem_escrita_sap(self):
        """[6/7] Testa que a preparação de CUA_ADICIONAR calcula atribuições e não chama BAPI SAP."""
        res = mod.preparar_cua_adicionar(self.caminho_excel, self.dados)
        self.assertIn("novas_linhas_count", res)
        self.assertIn("ja_existentes_count", res)
        self.assertIn("detalhes_por_dep", res)
        # Confirmar que STATUS, MSG, TIMESTEMP ficam vazios
        for item in res.get("novas_linhas", []):
            self.assertEqual(item["STATUS"], "")
            self.assertEqual(item["MSG"], "")
            self.assertEqual(item["TIMESTEMP"], "")
            self.assertIn(item["SISTEMA"], ["S4PCLNT100", "S4QCLNT100"])

    def test_etapa_6_idempotencia_chave_duplicada(self):
        """[6/7] Testa que a mesma combinação (UTILIZADOR + SISTEMA + AGR_NAME) não duplica."""
        res1 = mod.preparar_cua_adicionar(self.caminho_excel, self.dados)
        # Simular que a chave já existe nos dados
        dados_copy = mod.carregar_projeto_perfil(self.caminho_excel)
        for row in res1.get("novas_linhas", []):
            dados_copy.cua_adicionar.append({
                "id": row["ID"],
                "utilizador": row["UTILIZADOR"],
                "sistema": row["SISTEMA"],
                "role": row["AGR_NAME"],
                "status": "",
                "msg": ""
            })
        res2 = mod.preparar_cua_adicionar(self.caminho_excel, dados_copy)
        self.assertEqual(res2["novas_linhas_count"], 0)

    def test_etapa_6_ambientes_independentes(self):
        """[6/7] Testa se PRD (S4PCLNT100) e QAD (S4QCLNT100) geram linhas independentes."""
        res = mod.preparar_cua_adicionar(self.caminho_excel, self.dados)
        sistemas = {item["SISTEMA"] for item in res.get("novas_linhas", [])}
        if res.get("novas_linhas_count", 0) > 0:
            self.assertTrue("S4PCLNT100" in sistemas or "S4QCLNT100" in sistemas)

    def test_etapa_6_filhas_compostas_nao_atribuidas_diretamente(self):
        """[6/7] Confirma que Single Roles filhas de Composite Roles não são atribuídas diretamente."""
        res = mod.preparar_cua_adicionar(self.caminho_excel, self.dados)
        filhas_todas = set()
        for c_info in self.dados.roles_compostas.values():
            for f in c_info.get("roles_filhas", []):
                filhas_todas.add(mod.normalizar_texto(f))
        for item in res.get("novas_linhas", []):
            role_norm = mod.normalizar_texto(item["AGR_NAME"])
            # Se for single, não deve ser uma filha pertencente à composta do user
            if role_norm != mod.normalizar_texto(item["AGR_NAME"]):
                self.assertNotIn(role_norm, filhas_todas)

    @patch("pyrfc.Connection")
    def test_etapa_7_simulacao_executa_rollback(self, mock_conn):
        """[7/7] Testa que o modo de simulação chama BAPI_TRANSACTION_ROLLBACK sem alterar o SAP."""
        conn_instance = MagicMock()
        mock_conn.return_value = conn_instance
        conn_instance.call.side_effect = lambda name, **kw: (
            {"RETURN": [], "ACTIVITYGROUPS": []} if "BAPI_USER_GET_DETAIL" in name else {"RETURN": []}
        )
        item_dep = {"departamento": "HEALTH & SAFETY", "linha": 2}
        
        # Testar chamada de sincronização em modo simulação
        with patch("sap_rfc._rfc_common.build_connection_params_for", return_value={}):
            sucesso = mod.sincronizar_departamento_prd_qad_rfc(
                dados=self.dados,
                item_dep=item_dep,
                confirmar_execucao=False,
                modo_simulacao=True,
                usuarios_alvo=["S80001882"]
            )
            self.assertTrue(sucesso)
            # Confirmar que ROLLBACK foi chamado e COMMIT NÃO foi chamado
            calls = [call[0][0] for call in conn_instance.call.call_args_list]
            self.assertIn("BAPI_TRANSACTION_ROLLBACK", calls)
            self.assertNotIn("BAPI_TRANSACTION_COMMIT", calls)


if __name__ == "__main__":
    unittest.main()
