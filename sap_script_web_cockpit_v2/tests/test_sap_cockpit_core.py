"""Testes da Fase 2 do SAP Cockpit: seguranca de subprocessos, sem retry duplicado, contratos e import sem efeitos.

Requer as dependencias Windows do worker (pywin32, rich). Com o venv do cockpit
os testes sao saltados; executar com:
    worker\\.venv\\Scripts\\python.exe -m unittest tests.test_sap_cockpit_core
"""

from __future__ import annotations

import builtins
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
COCKPIT_DIR = HERE.parent
REPO_ROOT = COCKPIT_DIR.parent
for _p in (str(COCKPIT_DIR), str(REPO_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import sap_cockpit_web_ready as cockpit
    from cockpit_contracts import CockpitContractError, CockpitRunRequest, CockpitRunResult
    _IMPORT_ERROR = None
except ImportError as exc:  # venv sem pywin32/rich
    cockpit = None
    _IMPORT_ERROR = exc


def _escrever(pasta: str, nome: str, conteudo: str) -> str:
    caminho = os.path.join(pasta, nome)
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(conteudo)
    return caminho


@unittest.skipIf(cockpit is None, f"dependencias do worker em falta: {_IMPORT_ERROR}")
class ResolverCaminhosTest(unittest.TestCase):
    def setUp(self):
        self.raiz = tempfile.mkdtemp()
        self.processos = os.path.join(self.raiz, "Processos")
        self.irmao = os.path.join(self.raiz, "Processos_evil")
        os.makedirs(os.path.join(self.processos, "Proc"))
        os.makedirs(self.irmao)
        self._patch = mock.patch.object(cockpit, "PROCESSOS_DIR", self.processos)
        self._patch.start()

    def tearDown(self):
        self._patch.stop()
        shutil.rmtree(self.raiz, ignore_errors=True)

    def test_pasta_irma_com_mesmo_prefixo_e_rejeitada(self):
        # Antes: startswith aceitava "Processos_evil" como dentro de "Processos".
        with self.assertRaises(cockpit.SapCockpitError):
            cockpit._resolve_processo_path(self.irmao)

    def test_pasta_dentro_de_processos_e_aceite(self):
        self.assertEqual(
            os.path.normcase(cockpit._resolve_processo_path("Proc")),
            os.path.normcase(os.path.join(self.processos, "Proc")),
        )

    def test_subprocesso_com_travessia_e_rejeitado(self):
        for nome in ("../x.py", "..\\x.py", "sub/x.py", "C:x.py", "x/../../y"):
            with self.subTest(nome=nome), self.assertRaises(cockpit.SapCockpitError):
                cockpit._resolve_subprocesso_path(os.path.join(self.processos, "Proc"), nome)

    def test_subprocesso_absoluto_e_rejeitado(self):
        alvo = _escrever(self.raiz, "fora.py", "x = 1\n")
        with self.assertRaises(cockpit.SapCockpitError):
            cockpit._resolve_subprocesso_path(os.path.join(self.processos, "Proc"), alvo)

    def test_subprocesso_valido_fica_dentro_da_pasta(self):
        pasta = os.path.join(self.processos, "Proc")
        caminho = cockpit._resolve_subprocesso_path(pasta, "script")
        self.assertEqual(os.path.basename(caminho), "script.py")
        self.assertEqual(os.path.normcase(os.path.dirname(caminho)), os.path.normcase(pasta))

    def test_web_config_nao_le_fora_da_pasta(self):
        _escrever(self.raiz, "fora.py", "WEB_CONFIG = {'manages_own_session': True}\n")
        pasta = os.path.join(self.processos, "Proc")
        with self.assertRaises(cockpit.SapCockpitError):
            cockpit._read_subprocess_web_config(pasta, "../fora.py")


@unittest.skipIf(cockpit is None, f"dependencias do worker em falta: {_IMPORT_ERROR}")
class ExecucaoSemRetryTest(unittest.TestCase):
    def setUp(self):
        self.pasta = tempfile.mkdtemp()
        builtins.CONTA_EXEC_TESTE = 0
        self._patch = mock.patch.object(cockpit, "PROCESSOS_DIR", self.pasta)
        self._patch.start()

    def tearDown(self):
        self._patch.stop()
        shutil.rmtree(self.pasta, ignore_errors=True)
        builtins.__dict__.pop("CONTA_EXEC_TESTE", None)

    def _executar(self, nome):
        return cockpit._executar_um_script(
            ambiente_cockpit="DEV",
            caminho_pasta=self.pasta,
            processo_escolhido=nome,
            sistema_desejado="DEV",
            session=None,
            payload={},
            interactive=False,
        )

    def test_typeerror_dentro_do_script_nao_reexecuta(self):
        # Antes: TypeError interno levava a segunda chamada exec_fn(ambiente) => execucao duplicada.
        _escrever(self.pasta, "a.py", (
            "import builtins\n"
            "def executar(ambiente_cockpit):\n"
            "    builtins.CONTA_EXEC_TESTE += 1\n"
            "    raise TypeError('erro interno do script')\n"
        ))
        with self.assertRaises(TypeError):
            self._executar("a.py")
        self.assertEqual(builtins.CONTA_EXEC_TESTE, 1)

    def test_argumento_obrigatorio_em_falta_falha_antes_de_executar(self):
        _escrever(self.pasta, "b.py", (
            "import builtins\n"
            "def executar(ambiente_cockpit, obrigatorio):\n"
            "    builtins.CONTA_EXEC_TESTE += 1\n"
        ))
        with self.assertRaises(cockpit.SapCockpitError):
            self._executar("b.py")
        self.assertEqual(builtins.CONTA_EXEC_TESTE, 0)

    def test_script_sem_argumentos_extra_executa_uma_vez(self):
        _escrever(self.pasta, "c.py", (
            "import builtins\n"
            "def executar(ambiente_cockpit):\n"
            "    builtins.CONTA_EXEC_TESTE += 1\n"
            "    return {'status': 'ok'}\n"
        ))
        with mock.patch.object(cockpit, "read_sbar_status", return_value="sem-status"):
            resultado = self._executar("c.py")
        self.assertEqual(resultado, "ok")
        self.assertEqual(builtins.CONTA_EXEC_TESTE, 1)

    def test_nome_de_subprocesso_com_travessia_nao_carrega_modulo(self):
        with self.assertRaises(cockpit.SapCockpitError):
            self._executar("../outro.py")


@unittest.skipIf(cockpit is None, f"dependencias do worker em falta: {_IMPORT_ERROR}")
class ContratosTest(unittest.TestCase):
    def test_pedido_sem_processo_e_rejeitado(self):
        with self.assertRaises(CockpitContractError):
            CockpitRunRequest.from_payload({})

    def test_pedido_aceita_subprocesso_ou_script(self):
        self.assertEqual(CockpitRunRequest.from_payload({"processo": "P", "subprocesso": "a"}).subprocesso, "a")
        self.assertEqual(CockpitRunRequest.from_payload({"processo": "P", "script": "b"}).subprocesso, "b")

    def test_pedido_nao_dict_e_rejeitado(self):
        with self.assertRaises(CockpitContractError):
            CockpitRunRequest.from_payload(["P"])

    def test_resultado_tem_formato_que_o_worker_espera(self):
        self.assertEqual(CockpitRunResult(status="S", log="L").to_dict(), {"status": "S", "log": "L"})

    def test_run_sap_cockpit_sem_processo_falha_sem_tocar_no_sap(self):
        with mock.patch.object(cockpit, "obter_sessao_sap") as sessao:
            with self.assertRaises(CockpitContractError), mock.patch("sys.stdout"):
                cockpit.run_sap_cockpit({})
        sessao.assert_not_called()


@unittest.skipIf(cockpit is None, f"dependencias do worker em falta: {_IMPORT_ERROR}")
class ImportSemEfeitosTest(unittest.TestCase):
    def test_importar_motor_nao_abre_sap(self):
        import importlib
        import win32com.client

        sys.modules.pop("sap_cockpit_web_ready", None)
        with mock.patch.object(win32com.client, "GetObject", side_effect=AssertionError("SAP tocado no import")), \
                mock.patch.object(win32com.client, "Dispatch", side_effect=AssertionError("SAP tocado no import")):
            modulo = importlib.import_module("sap_cockpit_web_ready")
        self.assertTrue(hasattr(modulo, "run_sap_cockpit"))

    def test_execucao_direta_continua_a_ser_o_unico_ponto_de_entrada(self):
        fonte = (COCKPIT_DIR / "sap_cockpit_web_ready.py").read_text(encoding="utf-8")
        self.assertIn('if __name__ == "__main__":\n    main_terminal()', fonte)


if __name__ == "__main__":
    unittest.main()
