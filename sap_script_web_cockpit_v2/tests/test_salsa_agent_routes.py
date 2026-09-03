"""
Testes da camada de rotas do Agente Salsa IT (`/api/salsa-it-agent/pfcg/*`).

Nao sobem servidor: importam `web_api.main` e chamam as funcoes das rotas
diretamente, com a base de dados de jobs redirecionada para um diretorio
temporario. Cobrem o contrato partilhado por ~12 pares de rotas:

    POST  ->  cria job, devolve `job_id` + `state`
    GET   ->  mapeia estado (pending/running/failed/succeeded) e valida o job

Executar (a partir de sap_script_web_cockpit_v2/, com o venv do cockpit):
    python -m unittest tests.test_salsa_agent_routes
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

_COCKPIT_DIR = Path(__file__).resolve().parents[1]
_REPO_ROOT = Path(__file__).resolve().parents[2]
# web_api.* (cockpit) e sap_rfc.* (raiz do repo, usado pela validacao PFCG).
for _p in (str(_COCKPIT_DIR), str(_REPO_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# DATA_DIR tem de estar definido antes de importar web_api.store (le no import).
_TMPDIR = tempfile.mkdtemp(prefix="salsa_agent_routes_")
os.environ["DATA_DIR"] = _TMPDIR

import web_api.store as store  # noqa: E402
from web_api import main  # noqa: E402
from fastapi import HTTPException  # noqa: E402


def _body(response) -> dict:
    return json.loads(bytes(response.body).decode("utf-8"))


class SalsaAgentRoutesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        # Garante isolamento mesmo que um .env defina DATA_DIR.
        store.DATA_DIR = Path(_TMPDIR)
        store.DB_PATH = store.DATA_DIR / "sap_script_jobs.sqlite3"
        store.init_db()

    # ---- POST: criacao de job -------------------------------------------------

    def test_pfcg_analyze_creates_job(self) -> None:
        resp = main.api_salsa_it_pfcg_analyze(
            main.SalsaItPfcgAnalyzeRequest(role_name="Z_FI_CLIENTES")
        )
        data = _body(resp)
        self.assertIn("job_id", data)
        self.assertEqual(data["state"], "pending")
        self.assertEqual(data["role_name"], "Z_FI_CLIENTES")

        job = store.get_job(data["job_id"])
        self.assertEqual(job["task"], "pfcg_role_analysis")
        self.assertEqual(job["params"].get("role_name"), "Z_FI_CLIENTES")
        self.assertEqual(job["params"].get("system"), "PRD")  # default

    def test_pfcg_analyze_passes_system(self) -> None:
        data = _body(
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="Z_X", system="qad")
            )
        )
        self.assertEqual(store.get_job(data["job_id"])["params"].get("system"), "QAD")

    def test_pfcg_analyze_rejects_bad_system(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="Z_X", system="XPTO")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_pfcg_analyze_rejects_invalid_role_name(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="nome invalido !!")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_select_excel_creates_job(self) -> None:
        resp = main.api_salsa_it_pfcg_create_select_excel()
        data = _body(resp)
        self.assertIn("job_id", data)
        self.assertEqual(store.get_job(data["job_id"])["task"], "select_excel_file")

    def test_transactions_and_users_analyze_create_jobs(self) -> None:
        req = main.SalsaItPfcgAnalyzeRequest(role_name="Z_ROLE_X")
        t = _body(main.api_salsa_it_pfcg_transactions_analyze(req))
        u = _body(main.api_salsa_it_pfcg_users_analyze(req))
        self.assertEqual(store.get_job(t["job_id"])["task"], "pfcg_role_transactions_analysis")
        self.assertEqual(store.get_job(u["job_id"])["task"], "pfcg_role_users_analysis")

    def test_transaction_roles_create_job_and_states(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_transaction_roles(
                main.SalsaItPfcgTransactionRolesRequest(tcode="fb01")
            )
        )
        self.assertEqual(created["tcode"], "FB01")
        job = store.get_job(created["job_id"])
        self.assertEqual(job["task"], "pfcg_transaction_roles")
        self.assertEqual(job["params"].get("tcode"), "FB01")

        got = _body(main.api_salsa_it_pfcg_transaction_roles_job(created["job_id"]))
        self.assertEqual(got["state"], "pending")

        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps(
                {
                    "ok": True,
                    "status": "OK",
                    "tcode": "FB01",
                    "tcode_description": "Enter Incoming Invoice",
                    "count": 1,
                    "roles": [
                        {"role": "Z_FI_X", "description": "FI X", "composite_parents": ["C_FI"], "segredo": "x"}
                    ],
                    "system": "PRD",
                    "client": "100",
                    "segredo_top": "x",
                }
            ),
            "",
        )
        done = _body(main.api_salsa_it_pfcg_transaction_roles_job(created["job_id"]))
        self.assertEqual(done["state"], "succeeded")
        self.assertEqual(done["result"]["count"], 1)
        self.assertEqual(done["result"]["roles"][0]["role"], "Z_FI_X")
        self.assertEqual(done["result"]["roles"][0]["composite_parents"], ["C_FI"])
        self.assertNotIn("segredo", done["result"]["roles"][0])
        self.assertNotIn("segredo_top", done["result"])

    def test_transaction_roles_rejects_bad_tcode(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_transaction_roles(
                main.SalsaItPfcgTransactionRolesRequest(tcode="")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_object_roles_create_job_and_shape(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_object_roles(
                main.SalsaItPfcgObjectRolesRequest(auth_object="s_tcode")
            )
        )
        self.assertEqual(created["auth_object"], "S_TCODE")
        job = store.get_job(created["job_id"])
        self.assertEqual(job["task"], "pfcg_object_roles")
        self.assertEqual(job["params"].get("auth_object"), "S_TCODE")

        got = _body(main.api_salsa_it_pfcg_object_roles_job(created["job_id"]))
        self.assertEqual(got["state"], "pending")

        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps(
                {
                    "ok": True,
                    "status": "OK",
                    "auth_object": "S_TCODE",
                    "auth_object_text": "Transaction Code Check at Transaction Start",
                    "count": 1,
                    "roles": [{"role": "Z_X", "description": "X", "composite_parents": [], "seg": "x"}],
                    "system": "PRD",
                    "client": "100",
                }
            ),
            "",
        )
        done = _body(main.api_salsa_it_pfcg_object_roles_job(created["job_id"]))
        self.assertEqual(done["state"], "succeeded")
        self.assertEqual(done["result"]["roles"][0]["role"], "Z_X")
        self.assertNotIn("seg", done["result"]["roles"][0])

    def test_object_roles_rejects_bad_object(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_object_roles(
                main.SalsaItPfcgObjectRolesRequest(auth_object="")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_user_roles_create_job_and_shape(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_user_roles(
                main.SalsaItPfcgUserRolesRequest(username="clopes")
            )
        )
        self.assertEqual(created["username"], "CLOPES")
        job = store.get_job(created["job_id"])
        self.assertEqual(job["task"], "pfcg_user_roles")
        self.assertEqual(job["params"].get("username"), "CLOPES")

        got = _body(main.api_salsa_it_pfcg_user_roles_job(created["job_id"]))
        self.assertEqual(got["state"], "pending")

        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps(
                {
                    "ok": True,
                    "status": "OK",
                    "username": "CLOPES",
                    "count": 1,
                    "roles": [
                        {
                            "role": "Z_FI",
                            "description": "FI",
                            "valid_from": "01/01/2026",
                            "valid_to": "31/12/9999",
                            "assignment_status": "ATIVO",
                            "seg": "x",
                        }
                    ],
                    "system": "PRD",
                    "client": "100",
                }
            ),
            "",
        )
        done = _body(main.api_salsa_it_pfcg_user_roles_job(created["job_id"]))
        self.assertEqual(done["state"], "succeeded")
        self.assertEqual(done["result"]["roles"][0]["role"], "Z_FI")
        self.assertEqual(done["result"]["roles"][0]["assignment_status"], "ATIVO")
        self.assertNotIn("seg", done["result"]["roles"][0])

    def test_user_roles_rejects_bad_user(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_user_roles(
                main.SalsaItPfcgUserRolesRequest(username="")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_user_data_create_job_and_shape(self) -> None:
        created = _body(
            main.api_salsa_it_user_data(
                main.SalsaItUserDataRequest(username="clopes", kind="master", system="qad")
            )
        )
        self.assertEqual(created["username"], "CLOPES")
        self.assertEqual(created["kind"], "master")
        job = store.get_job(created["job_id"])
        self.assertEqual(job["task"], "user_data")
        self.assertEqual(job["params"].get("kind"), "master")
        self.assertEqual(job["params"].get("system"), "QAD")

        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps(
                {
                    "ok": True, "status": "OK", "username": "CLOPES", "kind": "master",
                    "fields": [{"label": "Tipo", "value": "A - Dialog", "x": 1}],
                    "system": "QAD", "client": "100",
                }
            ),
            "",
        )
        done = _body(main.api_salsa_it_user_data_job(created["job_id"]))
        self.assertEqual(done["state"], "succeeded")
        self.assertEqual(done["result"]["fields"][0]["label"], "Tipo")
        self.assertNotIn("x", done["result"]["fields"][0])

    def test_user_data_rejects_bad_kind(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_user_data(
                main.SalsaItUserDataRequest(username="X", kind="xpto")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    def test_user_search_create_job_and_shape(self) -> None:
        created = _body(
            main.api_salsa_it_user_search(
                main.SalsaItUserSearchRequest(query="Silva", system="prd")
            )
        )
        self.assertEqual(created["query"], "Silva")
        job = store.get_job(created["job_id"])
        self.assertEqual(job["task"], "user_search")
        self.assertEqual(job["params"].get("query"), "Silva")

        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps({
                "ok": True, "status": "OK", "query": "Silva", "count": 1,
                "users": [{"username": "JSILVA", "full_name": "João Silva", "x": 1}],
                "system": "PRD", "client": "100",
            }),
            "",
        )
        done = _body(main.api_salsa_it_user_search_job(created["job_id"]))
        self.assertEqual(done["state"], "succeeded")
        self.assertEqual(done["result"]["users"][0]["username"], "JSILVA")
        self.assertNotIn("x", done["result"]["users"][0])

    def test_user_search_rejects_short_query(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_user_search(
                main.SalsaItUserSearchRequest(query="a")
            )
        self.assertEqual(ctx.exception.status_code, 400)

    # ---- GET: mapeamento de estado / validacao do job ----------------------

    def test_analyze_job_pending_state(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="Z_PENDING")
            )
        )
        got = _body(main.api_salsa_it_pfcg_analyze_job(created["job_id"]))
        self.assertEqual(got["state"], "pending")

    def test_analyze_job_unknown_id_returns_404(self) -> None:
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_analyze_job("nao-existe-1234")
        self.assertEqual(ctx.exception.status_code, 404)

    def test_analyze_job_wrong_task_returns_400(self) -> None:
        other = store.create_job("select_excel_file", {})
        with self.assertRaises(HTTPException) as ctx:
            main.api_salsa_it_pfcg_analyze_job(other["id"])
        self.assertEqual(ctx.exception.status_code, 400)

    def test_analyze_job_succeeded_is_shaped(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="Z_DONE")
            )
        )
        store.complete_job(
            created["job_id"],
            "succeeded",
            json.dumps(
                {
                    "ok": True,
                    "status": "VALID",
                    "role": "Z_DONE",
                    "description": "desc",
                    "language": "PT",
                    "system": "PRD",
                    "client": "100",
                    "segredo_interno": "nao deve sair",
                }
            ),
            "",
        )
        got = _body(main.api_salsa_it_pfcg_analyze_job(created["job_id"]))
        self.assertEqual(got["state"], "succeeded")
        self.assertEqual(got["result"]["role"], "Z_DONE")
        self.assertNotIn("segredo_interno", got["result"])  # whitelist de campos

    def test_analyze_job_failed_state(self) -> None:
        created = _body(
            main.api_salsa_it_pfcg_analyze(
                main.SalsaItPfcgAnalyzeRequest(role_name="Z_FAIL")
            )
        )
        store.complete_job(created["job_id"], "failed", "boom", "")
        got = _body(main.api_salsa_it_pfcg_analyze_job(created["job_id"]))
        self.assertEqual(got["state"], "failed")


if __name__ == "__main__":
    unittest.main()
