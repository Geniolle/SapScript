import json
import sys
from pathlib import Path

if len(sys.argv) < 2:
    raise SystemExit("ERRO: informe o caminho do credentials.json")

path = Path(sys.argv[1])

if not path.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {path}")

if not path.is_file():
    raise SystemExit(f"ERRO: caminho não é ficheiro: {path}")

if path.stat().st_size <= 20:
    raise SystemExit(f"ERRO: ficheiro vazio ou incompleto. Tamanho={path.stat().st_size} bytes")

try:
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
except Exception as exc:
    raise SystemExit(f"ERRO: JSON inválido: {exc}")

if data.get("type") != "service_account":
    raise SystemExit("ERRO: campo 'type' diferente de 'service_account'.")

if not data.get("client_email"):
    raise SystemExit("ERRO: campo 'client_email' ausente.")

if not data.get("private_key"):
    raise SystemExit("ERRO: campo 'private_key' ausente.")

print("OK_SERVICE_ACCOUNT")
print("CLIENT_EMAIL=" + data["client_email"])
print("OK: credentials.json válido.")
