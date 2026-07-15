import time
import pytest
from fastapi.testclient import TestClient

# Adicionar o diretório web_api e worker ao sys.path se necessário
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2", "worker")))

from web_api.main import app, worker_last_seen, worker_last_seen_lock, WORKER_TOKEN
from web_api.store import create_job, get_job, get_connection

client = TestClient(app)

@pytest.fixture(autouse=True)
def clean_worker_state():
    """Limpa o dicionário de presença dos workers antes de cada teste."""
    with worker_last_seen_lock:
        worker_last_seen.clear()
    yield

def test_heartbeat_requires_valid_token():
    # 1. Enviar heartbeat com token inválido ou ausente -> Deve retornar 401
    response = client.post(
        "/api/worker/heartbeat",
        headers={"X-Worker-Token": "wrong-token"},
        json={"worker_name": "test-worker-1"}
    )
    assert response.status_code == 401
    assert "token" in response.json()["detail"].lower()

    # 2. Enviar heartbeat com token válido -> Deve retornar 200
    response = client.post(
        "/api/worker/heartbeat",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"worker_name": "test-worker-1"}
    )
    assert response.status_code == 200
    assert response.json()["status"] == "success"

def test_worker_status_online_offline_and_custom_threshold(monkeypatch):
    # Definir limite curto para testes (ex: 2 segundos) via monkeypatch
    monkeypatch.setattr("web_api.main.WORKER_OFFLINE_AFTER_SECONDS", 2.0)

    # Inicialmente, nenhum worker enviou ping -> Deve estar offline
    response = client.get("/api/worker/status")
    assert response.json()["status"] == "offline"

    # Enviar heartbeat do worker-1
    response = client.post(
        "/api/worker/heartbeat",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"worker_name": "test-worker-1"}
    )
    assert response.status_code == 200

    # Agora o status geral deve estar online
    response = client.get("/api/worker/status")
    assert response.json()["status"] == "online"

    # O status específico do worker-1 deve estar online
    response = client.get("/api/worker/status", params={"worker_name": "test-worker-1"})
    assert response.json()["status"] == "online"

    # Simular passagem do tempo (2.5 segundos) alterando manualmente a data no dict
    with worker_last_seen_lock:
        worker_last_seen["test-worker-1"] = time.time() - 2.5

    # Agora o status do worker-1 deve estar offline (ultrapassou o limite de 2 segundos)
    response = client.get("/api/worker/status", params={"worker_name": "test-worker-1"})
    assert response.json()["status"] == "offline"

    # O status geral também deve estar offline
    response = client.get("/api/worker/status")
    assert response.json()["status"] == "offline"

def test_multiple_workers_are_controlled_separately():
    # Enviar heartbeat do worker-A
    client.post(
        "/api/worker/heartbeat",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"worker_name": "worker-A"}
    )

    # Status do worker-A -> online
    response = client.get("/api/worker/status", params={"worker_name": "worker-A"})
    assert response.json()["status"] == "online"

    # Status do worker-B (não enviou heartbeat) -> offline
    response = client.get("/api/worker/status", params={"worker_name": "worker-B"})
    assert response.json()["status"] == "offline"

    # Status geral (pelo menos um online) -> online
    response = client.get("/api/worker/status")
    assert response.json()["status"] == "online"

def test_heartbeat_does_not_access_sqlite(monkeypatch):
    """Garante que o endpoint de heartbeat não realiza acessos à base de dados."""
    def mock_get_connection():
        raise RuntimeError("Acesso ao SQLite não é permitido no heartbeat")
    monkeypatch.setattr("web_api.store.get_connection", mock_get_connection)
    
    response = client.post(
        "/api/worker/heartbeat",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"worker_name": "test-worker-1"}
    )
    assert response.status_code == 200
    assert response.json()["status"] == "success"

def test_append_log_payload_and_preservation():
    """Garante que a rota POST /api/jobs/{job_id}/log não retorna o log completo e preserva os dados."""
    job = create_job(task="test_task", params={})
    job_id = job["id"]

    # 1. Enviar primeiro lote de logs
    response = client.post(
        f"/api/jobs/{job_id}/log",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"log_line": "Lote 1 de logs"}
    )
    assert response.status_code == 200
    res_data = response.json()
    assert res_data["status"] == "ok"
    assert res_data["job_id"] == job_id
    assert res_data["appended"] is True
    assert "log" not in res_data  # Não deve conter o log completo de volta

    # 2. Enviar segundo lote de logs
    client.post(
        f"/api/jobs/{job_id}/log",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"log_line": "Lote 2 de logs"}
    )

    # 3. Validar se os dados foram gravados na base de dados
    job_in_db = get_job(job_id)
    assert "Lote 1 de logs" in job_in_db["log"]
    assert "Lote 2 de logs" in job_in_db["log"]

def test_cancelled_job_returns_409():
    """Garante que se o job estiver cancelado (failed), a gravação de logs retorna 409."""
    job = create_job(task="test_task", params={})
    job_id = job["id"]

    # Atualizar estado do job para failed para simular cancelamento
    with get_connection() as conn:
        conn.execute("UPDATE jobs SET state = 'failed' WHERE id = ?", (job_id,))

    response = client.post(
        f"/api/jobs/{job_id}/log",
        headers={"X-Worker-Token": WORKER_TOKEN},
        json={"log_line": "Log após cancelamento"}
    )
    assert response.status_code == 409
    assert "cancelled" in response.json()["detail"].lower()

def test_worker_heartbeat_client_logging(monkeypatch, capsys):
    """Valida o comportamento do cliente de heartbeat: falha temporária, restabelecida e cooldown."""
    import worker as worker_mod
    
    monkeypatch.setenv("API_BASE_URL", "http://localhost:8000")
    monkeypatch.setenv("WORKER_NAME", "test-worker-client")
    
    post_calls = []
    class MockResponse:
        def __init__(self, status_code):
            self.status_code = status_code
        def raise_for_status(self):
            if self.status_code >= 400:
                raise Exception(f"HTTP Error {self.status_code}")

    resp_seq = [
        MockResponse(500),  # 1. Falha inicial (exibe erro)
        MockResponse(500),  # 2. Falha contínua (oculta devido ao cooldown de 60s)
        MockResponse(200),  # 3. Sucesso (exibe restabelecida)
    ]
    resp_iter = iter(resp_seq)

    def mock_post(self, url, headers=None, json=None, timeout=None):
        post_calls.append((url, json))
        return next(resp_iter)

    monkeypatch.setattr("requests.Session.post", mock_post)

    # Controlar a execução da thread
    run_count = 0
    def mock_is_set():
        nonlocal run_count
        if run_count >= 3:
            return True
        run_count += 1
        return False

    monkeypatch.setattr(worker_mod.stop_event, "is_set", mock_is_set)
    monkeypatch.setattr(worker_mod.stop_event, "wait", lambda t: None)

    worker_mod.heartbeat_loop()

    captured = capsys.readouterr()
    
    # 1. Deve registrar a primeira falha
    assert "Falha no heartbeat do worker" in captured.out
    
    # 2. Não deve poluir o console com o log da segunda falha contínua no cooldown
    assert "Falha contínua no heartbeat do worker" not in captured.out
    
    # 3. Deve avisar quando restabelecida a conexão
    assert "Conexão de heartbeat restabelecida" in captured.out
