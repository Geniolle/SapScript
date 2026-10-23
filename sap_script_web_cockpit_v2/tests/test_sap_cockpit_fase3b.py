"""Testes da Fase 3b do SAP Cockpit.

- Equivalencia das funcoes puras centralizadas em sap_cockpit_core.py contra o comportamento original
  (fixture gerada a partir de HEAD b34330e, antes da centralizacao).
- Validacao de assinatura antes de executar(), sem retry.
- Regressao dos dois caminhos de reexecucao do SAP Cockpit.py (listar_requests e executar_processo).
  O desktop nao e importado: executa codigo de topo (SAP). As funcoes sao extraidas por AST e
  executadas com stubs, sem SAP.

Nao toca em SAP, nao le .env, nao imprime credenciais.
"""

from __future__ import annotations

import ast
import builtins
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
COCKPIT_DIR = HERE.parent
REPO_ROOT = COCKPIT_DIR.parent
for _p in (str(COCKPIT_DIR), str(REPO_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import inspect  # noqa: E402

import sap_cockpit_core as core  # noqa: E402

DESKTOP_PATH = REPO_ROOT / "SAP Cockpit.py"
WEB_PATH = COCKPIT_DIR / "sap_cockpit_web_ready.py"
GOLDEN = json.loads((HERE / "fixtures" / "sap_cockpit_golden_puras.json").read_text(encoding="utf-8"))


def _extrair(path: Path, nome: str, ns: dict) -> dict:
    """Compila apenas a definicao `nome` do ficheiro, sem executar o resto do modulo."""
    src = path.read_text(encoding="utf-8")
    arvore = ast.parse(src, filename=str(path))
    for no in arvore.body:
        if isinstance(no, ast.FunctionDef) and no.name == nome:
            exec(compile(ast.Module(body=[no], type_ignores=[]), str(path), "exec"), ns)
            return ns
    raise AssertionError(f"{nome} nao encontrado em {path.name}")


class EquivalenciaFuncoesPurasTest(unittest.TestCase):
    def test_parse_env_line_igual_ao_original(self):
        for caso in GOLDEN["parse"]:
            with self.subTest(entrada=caso["in"]):
                self.assertEqual(list(core.parse_env_line(caso["in"])), caso["d"])

    def test_is_scripting_disabled_error_igual_ao_original(self):
        for caso in GOLDEN["disabled"]:
            mensagem, excinfo = caso["in"]
            exc = Exception(mensagem if mensagem is not None else "")
            if excinfo is not None:
                exc.excepinfo = excinfo
            with self.subTest(entrada=caso["in"]):
                self.assertEqual(core.is_scripting_disabled_error(exc), caso["d"])

    def test_desktop_e_web_usam_a_implementacao_central(self):
        for caminho in (DESKTOP_PATH, WEB_PATH):
            arvore = ast.parse(caminho.read_text(encoding="utf-8"))
            definidas = {n.name for n in arvore.body if isinstance(n, ast.FunctionDef)}
            with self.subTest(ficheiro=caminho.name):
                self.assertNotIn("_parse_env_line", definidas)
                self.assertNotIn("_is_scripting_disabled_error", definidas)
                importa = [
                    a.name
                    for n in arvore.body
                    if isinstance(n, ast.ImportFrom) and n.module == "sap_cockpit_core"
                    for a in n.names
                ]
                self.assertIn("parse_env_line", importa)
                self.assertIn("is_scripting_disabled_error", importa)


class ValidarArgumentosExecTest(unittest.TestCase):
    def test_argumento_obrigatorio_em_falta_levanta_erro(self):
        def executar(ambiente_cockpit, obrigatorio):
            return None

        with self.assertRaises(core.ArgumentosExecInvalidos):
            core.validar_argumentos_exec(executar, "DEV", {}, "p.py")

    def test_argumento_desconhecido_levanta_erro(self):
        def executar(ambiente_cockpit):
            return None

        with self.assertRaises(core.ArgumentosExecInvalidos):
            core.validar_argumentos_exec(executar, "DEV", {"x": 1}, "p.py")

    def test_assinatura_compativel_passa(self):
        def executar(ambiente_cockpit, obrigatorio=None, **extra):
            return None

        core.validar_argumentos_exec(executar, "DEV", {"obrigatorio": 1, "outro": 2}, "p.py")

    def test_validacao_nao_chama_a_funcao(self):
        chamadas = []

        def executar(ambiente_cockpit):
            chamadas.append(1)

        core.validar_argumentos_exec(executar, "DEV", {}, "p.py")
        self.assertEqual(chamadas, [])

    def test_erro_inclui_nome_do_processo_sem_valores(self):
        def executar(ambiente_cockpit, obrigatorio):
            return None

        with self.assertRaises(core.ArgumentosExecInvalidos) as ctx:
            core.validar_argumentos_exec(executar, "DEV", {"segredo": "valor-teste"}, "meu.py")
        self.assertIn("meu.py", str(ctx.exception))


class DesktopSemRetryTest(unittest.TestCase):
    """Regressao de SAP Cockpit.py::executar_processo (antes: exec_fn(...) seguido de exec_fn(ambiente))."""

    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        builtins.CONTA_EXEC_DESKTOP = 0
        self.erros = []
        self.infos = []

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)
        builtins.__dict__.pop("CONTA_EXEC_DESKTOP", None)

    def _ns(self, respostas):
        respostas = list(respostas)
        ns = {
            "os": os,
            "importlib": __import__("importlib"),
            "inspect": inspect,
            "sys": sys,
            "info": lambda *a, **k: None,
            "ok": lambda *a, **k: None,
            "warn": lambda *a, **k: None,
            "erro": lambda msg, *a, **k: self.erros.append(str(msg)),
            "mostrar_subprocessos": lambda *a, **k: None,
            "_resetar_env_request": lambda: None,
            "perguntar_opcao_request": lambda *a, **k: {},
            "input": lambda *a, **k: respostas.pop(0),
            "validar_argumentos_exec": core.validar_argumentos_exec,
        }
        _extrair(DESKTOP_PATH, "_analisar_exec_signature", ns)
        _extrair(DESKTOP_PATH, "executar_processo", ns)
        return ns

    def _correr(self, script: str, respostas):
        _escrever(self.pasta, "a.py", script)
        ns = self._ns(respostas)
        ns["executar_processo"]("DEV", self.pasta, "S4D", None)

    def test_typeerror_interno_do_script_nao_reexecuta(self):
        self._correr(
            "import builtins\n"
            "def executar(ambiente_cockpit):\n"
            "    builtins.CONTA_EXEC_DESKTOP += 1\n"
            "    raise TypeError('erro interno do script')\n",
            ["1", "2"],
        )
        self.assertEqual(builtins.CONTA_EXEC_DESKTOP, 1)
        self.assertTrue(any("Erro ao executar" in e for e in self.erros), self.erros)

    def test_argumento_em_falta_nao_executa_o_script(self):
        self._correr(
            "import builtins\n"
            "def executar(ambiente_cockpit, obrigatorio):\n"
            "    builtins.CONTA_EXEC_DESKTOP += 1\n",
            ["1", "2"],
        )
        self.assertEqual(builtins.CONTA_EXEC_DESKTOP, 0)
        self.assertTrue(any("Argumentos incompativeis" in e for e in self.erros), self.erros)

    def test_execucao_normal_corre_uma_vez(self):
        self._correr(
            "import builtins\n"
            "def executar(ambiente_cockpit):\n"
            "    builtins.CONTA_EXEC_DESKTOP += 1\n",
            ["1", "2"],
        )
        self.assertEqual(builtins.CONTA_EXEC_DESKTOP, 1)
        self.assertEqual(self.erros, [])


class DesktopListarRequestsSemRetryTest(unittest.TestCase):
    """Regressao de SAP Cockpit.py::perguntar_opcao_request (antes: repetia listar_requests sem argumentos)."""

    def test_falha_de_listar_requests_nao_repete_a_chamada(self):
        chamadas = []

        def listar_requests(**kwargs):
            chamadas.append(kwargs)
            if "use_new_mode" in kwargs:
                raise TypeError("erro interno de listar_requests")
            return []

        modulo = type("ModPesq", (), {"listar_requests": staticmethod(listar_requests)})
        erros = []
        ns = {
            "linha": lambda *a, **k: None,
            "destaque": lambda *a, **k: None,
            "erro": lambda msg, *a, **k: erros.append(str(msg)),
            "warn": lambda *a, **k: None,
            "ok": lambda *a, **k: None,
            "print": lambda *a, **k: None,
            "input": lambda *a, **k: "3",
            "validar_request": lambda v: v,
            "carregar_pesquisar_request": lambda: modulo,
            "escolher_request_por_linha": lambda *a, **k: None,
        }
        _extrair(DESKTOP_PATH, "perguntar_opcao_request", ns)

        resultado = ns["perguntar_opcao_request"]("S4D", None)

        self.assertEqual(len(chamadas), 1, chamadas)
        self.assertEqual(resultado["request_option"], "3")
        self.assertTrue(any("Falha ao gerar lista" in e for e in erros), erros)


class ExecucaoWebUmaVezTest(unittest.TestCase):
    """Garante a mesma regra na web: um pedido -> uma execucao de executar()."""

    def test_web_executa_uma_vez_quando_typeerror_e_interno(self):
        import importlib

        try:
            cockpit = importlib.import_module("sap_cockpit_web_ready")
        except ImportError as exc:
            self.skipTest(f"dependencias do worker em falta: {exc}")

        pasta = tempfile.mkdtemp()
        builtins.CONTA_EXEC_WEB = 0
        try:
            _escrever(pasta, "a.py", (
                "import builtins\n"
                "def executar(ambiente_cockpit):\n"
                "    builtins.CONTA_EXEC_WEB += 1\n"
                "    raise TypeError('erro interno')\n"
            ))
            with self.assertRaises(TypeError):
                cockpit._executar_um_script(
                    ambiente_cockpit="DEV", caminho_pasta=pasta, processo_escolhido="a.py",
                    sistema_desejado="S4D", session=None, payload={}, interactive=False,
                )
            self.assertEqual(builtins.CONTA_EXEC_WEB, 1)
        finally:
            shutil.rmtree(pasta, ignore_errors=True)
            builtins.__dict__.pop("CONTA_EXEC_WEB", None)


def _escrever(pasta: str, nome: str, conteudo: str) -> None:
    with open(os.path.join(pasta, nome), "w", encoding="utf-8") as f:
        f.write(conteudo)


if __name__ == "__main__":
    unittest.main()
