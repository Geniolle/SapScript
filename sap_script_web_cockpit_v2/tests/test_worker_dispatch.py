"""
Fase 4-resto: bloqueia o contrato de tasks do worker.

`worker/sap_tasks.py` importa pywin32, portanto nao e importavel fora do worker.
Este teste analisa o ficheiro por AST (sem o importar) e garante que:

- `TASK_HANDLERS` cobre exatamente o mesmo conjunto de tasks que a antiga
  cadeia `if task == "..."` (menos `sap_cockpit`, que fica explicito);
- `run_sap_task` mantem o dispatch por dict, o ramo `sap_cockpit` e o
  fallback "Rotina desconhecida";
- nenhuma das tasks migradas voltou a aparecer como `if task == "..."`.

    python -m unittest tests.test_worker_dispatch
"""

from __future__ import annotations

import ast
import re
import unittest
from pathlib import Path

SAP_TASKS = Path(__file__).resolve().parents[1] / "worker" / "sap_tasks.py"

# Contrato: todas as tasks que o worker sabe executar (inclui sap_cockpit).
EXPECTED_TASKS = {
    "sap_agent_analysis",
    "sap_cockpit_auto_trigger",
    "pfcg_role_analysis",
    "pfcg_role_transactions_analysis",
    "pfcg_role_users_analysis",
    "pfcg_transaction_roles",
    "pfcg_object_roles",
    "pfcg_user_roles",
    "user_data",
    "pfcg_create_excel_analysis",
    "pfcg_role_create_preview",
    "pfcg_role_create_rfc",
    "pfcg_composta_create_preview",
    "pfcg_composta_create",
    "pfcg_role_delete_preview",
    "pfcg_role_delete_rfc",
    "pfcg_transport_search",
    "sap_search_requests",
    "select_excel_file",
    "ping_status",
    "open_transaction",
    "sap_gui_chat_action",
    "fi_default_document",
    "f110_proposal",
    "f110_payment",
    "sap_cockpit",
}
# sap_cockpit fica como ramo explicito (streaming/threads/documentacao proprios).
DICT_TASKS = EXPECTED_TASKS - {"sap_cockpit"}


class WorkerDispatchTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = SAP_TASKS.read_text(encoding="utf-8")
        cls.tree = ast.parse(cls.source)

    def _task_handlers_keys(self) -> set[str]:
        for node in ast.walk(self.tree):
            targets = (
                node.targets
                if isinstance(node, ast.Assign)
                else [node.target]
                if isinstance(node, ast.AnnAssign)
                else []
            )
            if any(getattr(t, "id", None) == "TASK_HANDLERS" for t in targets):
                self.assertIsInstance(node.value, ast.Dict)
                return {k.value for k in node.value.keys}
        self.fail("TASK_HANDLERS nao encontrado em sap_tasks.py")

    def test_task_handlers_cobre_o_contrato(self) -> None:
        self.assertEqual(self._task_handlers_keys(), DICT_TASKS)

    def test_run_sap_task_tem_dispatch_cockpit_e_fallback(self) -> None:
        src = self.source
        self.assertIn("handler = TASK_HANDLERS.get(task)", src)
        self.assertIn('if task == "sap_cockpit":', src)
        self.assertIn('raise SapExecutionError(f"Rotina desconhecida: {task}")', src)

    def test_tasks_migradas_nao_ficaram_como_if_chain(self) -> None:
        if_chain = set(re.findall(r'if task == "([a-z0-9_]+)"', self.source))
        leftover = (if_chain & DICT_TASKS)
        self.assertEqual(leftover, set(), f"tasks migradas ainda em if-chain: {leftover}")
        self.assertEqual(if_chain, {"sap_cockpit"})


if __name__ == "__main__":
    unittest.main()
