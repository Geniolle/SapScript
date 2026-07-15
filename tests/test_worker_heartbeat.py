import time
import pytest
from fastapi.testclient import TestClient

# Adicionar o diretório web_api ao sys.path se necessário
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))

from web_api.main import app, worker_last_seen, worker_last_seen_lock, WORKER_TOKEN

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
