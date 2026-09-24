from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from unittest import TestCase, mock

# Ensure project root is on sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


class ProjetoPerfilAgentMenuTests(TestCase):
    """Verifica se o submenu e as 6 ações nativas estão no cockpit.agent.js."""

    def setUp(self) -> None:
        self.js_path = PROJECT_ROOT / "sap_script_web_cockpit_v2" / "web_api" / "static" / "js" / "cockpit.agent.js"
        self.assertTrue(self.js_path.exists(), "cockpit.agent.js deve existir")
        self.js_content = self.js_path.read_text(encoding="utf-8")

    def test_menu_contains_all_six_native_options(self) -> None:
        expected_options = [
            ("projeto-perfil-execucao", "Execução"),
            ("projeto-perfil-departamento", "Departamento"),
            ("projeto-perfil-utilizador", "Utilizador"),
            ("projeto-perfil-pesquisa", "Pesquisa"),
            ("projeto-perfil-corrigir", "Corrigir"),
            ("projeto-perfil-su53", "SU53"),
        ]
        for opt_id, label in expected_options:
            with self.subTest(option_id=opt_id):
                self.assertIn(f"id: '{opt_id}'", self.js_content)
                self.assertIn(f"label: '{label}'", self.js_content)

    def test_menu_action_handlers_defined(self) -> None:
        handlers = [
            "action.id === 'projeto-perfil-execucao'",
            "action.id === 'projeto-perfil-departamento'",
            "action.id === 'projeto-perfil-utilizador'",
            "action.id === 'projeto-perfil-pesquisa'",
            "action.id === 'projeto-perfil-corrigir'",
            "action.id === 'projeto-perfil-su53'",
        ]
        for h in handlers:
            with self.subTest(handler=h):
                self.assertIn(h, self.js_content)

    def test_awaiting_input_handlers_defined(self) -> None:
        inputs = [
            "ASI_PROJETO_PERFIL_UTILIZADOR_INPUT",
            "ASI_PROJETO_PERFIL_PESQUISA_TCODE_INPUT",
            "ASI_PROJETO_PERFIL_PESQUISA_USER_INPUT",
            "ASI_PROJETO_PERFIL_DEPTO_INPUT",
            "ASI_PROJETO_PERFIL_SU53_USER_INPUT",
        ]
        for inp in inputs:
            with self.subTest(input_const=inp):
                self.assertIn(inp, self.js_content)


class ProjetoPerfilWorkerTaskTests(TestCase):
    """Verifica se os task handlers do worker SAP estão registados."""

    def test_task_handlers_registered(self) -> None:
        from sap_script_web_cockpit_v2.worker.sap_tasks import TASK_HANDLERS

        expected_tasks = [
            "projeto_perfil_execucao",
            "projeto_perfil_departamento",
            "projeto_perfil_utilizador",
            "projeto_perfil_pesquisa",
            "projeto_perfil_corrigir",
            "projeto_perfil_su53",
        ]
        for t in expected_tasks:
            with self.subTest(task_name=t):
                self.assertIn(t, TASK_HANDLERS)
                self.assertTrue(callable(TASK_HANDLERS[t]))


class ProjetoPerfilCliBridgeTests(TestCase):
    """Testa a CLI bridge sap_rfc.projeto_perfil_cli para ações sem bloquear em input()."""

    def test_cli_departamento_listar(self) -> None:
        from sap_rfc.projeto_perfil_cli import run_cli

        ret = run_cli(["--action", "departamento", "--subacao", "listar"])
        self.assertEqual(ret, 0)

    def test_cli_pesquisa_tcode(self) -> None:
        from sap_rfc.projeto_perfil_cli import run_cli

        ret = run_cli(["--action", "pesquisa", "--tcode", "ME23N"])
        self.assertEqual(ret, 0)


class ProjetoPerfilServiceTests(TestCase):
    """Testa o serviço Python unificado do Projeto Perfil."""

    def test_fluxo_departamento_listar(self) -> None:
        from sap_rfc.projeto_perfil_service import fluxo_departamento

        res = fluxo_departamento(subacao="listar")
        self.assertTrue(res.get("ok"))
        self.assertEqual(res.get("subacao"), "listar")
        self.assertIsInstance(res.get("departamentos"), list)
        self.assertGreater(len(res.get("departamentos", [])), 0)

    def test_pesquisa_transacao_existente(self) -> None:
        from sap_rfc.projeto_perfil_service import pesquisa_atribuir_transacao

        res = pesquisa_atribuir_transacao("ME23N")
        self.assertTrue(res.get("ok"))
        self.assertEqual(res.get("tcode"), "ME23N")
        self.assertTrue(res.get("encontrado"))
        self.assertGreater(len(res.get("roles", [])), 0)

    def test_pesquisa_transacao_inexistente(self) -> None:
        from sap_rfc.projeto_perfil_service import pesquisa_atribuir_transacao

        res = pesquisa_atribuir_transacao("ZZZZ_INEXISTENTE_999")
        self.assertTrue(res.get("ok"))
        self.assertFalse(res.get("encontrado"))
        self.assertEqual(res.get("roles"), [])

    def test_auditoria_utilizador_vazio(self) -> None:
        from sap_rfc.projeto_perfil_service import auditoria_utilizador

        res = auditoria_utilizador("")
        self.assertFalse(res.get("ok"))

    def test_diagnostico_su53_vazio(self) -> None:
        from sap_rfc.projeto_perfil_service import diagnostico_su53

        res = diagnostico_su53("")
        self.assertFalse(res.get("ok"))

    def test_execucao_completa_assumir_sim_nao_bloqueia_input(self) -> None:
        """Valida que executar_processo_completo com assumir_sim=True não chama input()."""
        from sap_rfc.projeto_perfil_service import executar_processo_completo

        # Mock input to fail immediately if called
        with mock.patch("builtins.input", side_effect=AssertionError("input() não deve ser chamado no Agente!")):
            # Se a fila estiver vazia ou com pendências, assumir_sim=True deve processar sem pedir input
            res = executar_processo_completo(assumir_sim=True)
            self.assertIn("ok", res)
            self.assertIn("status", res)

