from __future__ import annotations

from pathlib import Path
from datetime import datetime
import json


ROOT = Path.cwd()
WORKFLOW_ENGINE_PATH = ROOT / "workflow_engine.py"
STATE_PATH = ROOT / "cache" / "workflow_state.json"

STATE_KEY_TO_REMOVE = "IZ-56831|FI Extracto Cadeias de Pesquisa|04/05/2026"


def backup_file(path: Path) -> None:
    if not path.exists():
        return
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(path.suffix + f".bak_{timestamp}")
    backup_path.write_bytes(path.read_bytes())
    print(f"OK: backup criado: {backup_path}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


####################################################################################
# (3) CORRIGIR workflow_engine.py
####################################################################################

if not WORKFLOW_ENGINE_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {WORKFLOW_ENGINE_PATH}")

backup_file(WORKFLOW_ENGINE_PATH)

content = read_text(WORKFLOW_ENGINE_PATH)

old_block = '''def _load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with open(path, "r", encoding="utf-8") as file_obj:
        return json.load(file_obj)
'''

new_block = '''def _load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with open(path, "r", encoding="utf-8-sig") as file_obj:
        return json.load(file_obj)
'''

if new_block in content:
    print("INFO: workflow_engine.py já lê JSON com utf-8-sig.")
else:
    if old_block not in content:
        raise SystemExit("ERRO: bloco _load_json esperado não encontrado em workflow_engine.py.")
    content = content.replace(old_block, new_block)
    write_text(WORKFLOW_ENGINE_PATH, content)
    print("OK: workflow_engine.py atualizado para ler JSON com utf-8-sig.")


####################################################################################
# (4) NORMALIZAR cache/workflow_state.json SEM BOM E REMOVER CHAVE DO TICKET
####################################################################################

if not STATE_PATH.exists():
    print("INFO: workflow_state.json não existe. Nada para normalizar.")
else:
    backup_file(STATE_PATH)

    raw = STATE_PATH.read_text(encoding="utf-8-sig").strip()
    if not raw:
        state = {"processed": {}}
    else:
        state = json.loads(raw)

    if not isinstance(state, dict):
        state = {"processed": {}}

    processed = state.get("processed")
    if not isinstance(processed, dict):
        processed = {}
        state["processed"] = processed

    if STATE_KEY_TO_REMOVE in processed:
        del processed[STATE_KEY_TO_REMOVE]
        print(f"OK: chave removida do cache: {STATE_KEY_TO_REMOVE}")
    else:
        print(f"INFO: chave não existia no cache: {STATE_KEY_TO_REMOVE}")

    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("OK: workflow_state.json normalizado sem BOM.")


####################################################################################
# (5) VALIDAÇÕES
####################################################################################

final_content = read_text(WORKFLOW_ENGINE_PATH)

if 'encoding="utf-8-sig"' not in final_content:
    raise SystemExit("ERRO: workflow_engine.py não ficou com utf-8-sig.")

if STATE_PATH.exists():
    # Validação direta com utf-8 normal. Se tiver BOM, isto volta a falhar.
    with STATE_PATH.open("r", encoding="utf-8") as file_obj:
        json.load(file_obj)
    print("OK: workflow_state.json abre com utf-8 normal.")

print("OK: patch concluído.")
