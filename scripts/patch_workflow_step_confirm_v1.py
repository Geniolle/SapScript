from __future__ import annotations

from pathlib import Path
from datetime import datetime


ROOT = Path.cwd()
WORKFLOW_ENGINE_PATH = ROOT / "workflow_engine.py"
ENV_EXAMPLE_PATH = ROOT / ".env.example"


def backup_file(path: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(path.suffix + f".bak_{timestamp}")
    backup_path.write_text(path.read_text(encoding="utf-8-sig"), encoding="utf-8")
    print(f"OK: backup criado: {backup_path}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


if not WORKFLOW_ENGINE_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {WORKFLOW_ENGINE_PATH}")

if not ENV_EXAMPLE_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {ENV_EXAMPLE_PATH}")

backup_file(WORKFLOW_ENGINE_PATH)
backup_file(ENV_EXAMPLE_PATH)

workflow_engine = read_text(WORKFLOW_ENGINE_PATH)
env_example = read_text(ENV_EXAMPLE_PATH)


####################################################################################
# (3) ADICIONAR FUNÇÃO DE PAUSA NO workflow_engine.py
####################################################################################

if "def _pause_before_step_if_enabled(" in workflow_engine:
    print("INFO: função _pause_before_step_if_enabled já existe. Nada a inserir.")
else:
    marker = '''def _step_capture_evidence_enabled(step: Dict[str, Any]) -> bool:
    if "capture_evidence" in step:
        return _to_bool(str(step.get("capture_evidence", "")))
    if "capture_runtime_snapshot" in step:
        return _to_bool(str(step.get("capture_runtime_snapshot", "")))
    return False
'''

    insert_block = '''def _step_capture_evidence_enabled(step: Dict[str, Any]) -> bool:
    if "capture_evidence" in step:
        return _to_bool(str(step.get("capture_evidence", "")))
    if "capture_runtime_snapshot" in step:
        return _to_bool(str(step.get("capture_runtime_snapshot", "")))
    return False


def _pause_before_step_if_enabled(
    *,
    workflow_name: str,
    step_name: str,
    index: int,
    total: int,
    row_context: Dict[str, str],
) -> None:
    if not _to_bool(os.getenv("WORKFLOW_STEP_CONFIRM", "false")):
        return

    ticket_key = str(row_context.get("ticket_key", "") or "").strip() or "-"
    categoria = str(row_context.get("categoria_sap", "") or "").strip() or workflow_name
    request_number = str(row_context.get("request_number", "") or "").strip() or "-"

    message = (
        "\\n"
        "================================================================================\\n"
        "PAUSA DE VALIDAÇÃO DO WORKFLOW\\n"
        f"Ticket: {ticket_key}\\n"
        f"Categoria: {categoria}\\n"
        f"Workflow: {workflow_name}\\n"
        f"Step: {index}/{total} - {step_name}\\n"
        f"Request atual: {request_number}\\n"
        "Pressiona ENTER para executar este step, ou CTRL+C para interromper.\\n"
        "================================================================================\\n"
    )

    if not sys.stdin or not sys.stdin.isatty():
        logging.warning(
            "WORKFLOW_STEP_CONFIRM ativo, mas stdin nao e interativo; pausa ignorada para o step '%s'.",
            step_name,
        )
        return

    print(message, flush=True)
    input()
'''

    if marker not in workflow_engine:
        raise SystemExit("ERRO: bloco _step_capture_evidence_enabled não encontrado em workflow_engine.py.")

    workflow_engine = workflow_engine.replace(marker, insert_block)
    print("OK: função _pause_before_step_if_enabled inserida.")


####################################################################################
# (4) CHAMAR A PAUSA ANTES DE CADA STEP
####################################################################################

if "_pause_before_step_if_enabled(" in workflow_engine and "PAUSA DE VALIDAÇÃO DO WORKFLOW" in workflow_engine:
    old_call_marker = '''        ok, error, runtime_snapshot = _run_step(
            step=step,
            step_name=step_name,
            context=row_context,
            base_dir=base_dir,
            python_exec=python_exec,
            documentation=documentation,
        )'''

    new_call_block = '''        _pause_before_step_if_enabled(
            workflow_name=workflow_name,
            step_name=step_name,
            index=index,
            total=len(steps),
            row_context=row_context,
        )

        ok, error, runtime_snapshot = _run_step(
            step=step,
            step_name=step_name,
            context=row_context,
            base_dir=base_dir,
            python_exec=python_exec,
            documentation=documentation,
        )'''

    if new_call_block in workflow_engine:
        print("INFO: chamada da pausa já existe antes de _run_step. Nada a alterar.")
    else:
        if old_call_marker not in workflow_engine:
            raise SystemExit("ERRO: chamada _run_step esperada não encontrada em workflow_engine.py.")

        workflow_engine = workflow_engine.replace(old_call_marker, new_call_block)
        print("OK: pausa adicionada antes de cada step.")


####################################################################################
# (5) ATUALIZAR .env.example
####################################################################################

if "WORKFLOW_STEP_CONFIRM" in env_example:
    print("INFO: WORKFLOW_STEP_CONFIRM já existe no .env.example.")
else:
    env_block = '''
# Pausa manual entre steps do workflow.
# Usar apenas em testes locais/interativos.
# true  = pede ENTER antes de cada step configurado em workflows.json
# false = execução automática normal
WORKFLOW_STEP_CONFIRM=false
'''
    marker_env = "WORKFLOW_ENABLED=true"

    if marker_env not in env_example:
        raise SystemExit("ERRO: WORKFLOW_ENABLED=true não encontrado no .env.example.")

    env_example = env_example.replace(marker_env, marker_env + env_block)
    print("OK: .env.example atualizado com WORKFLOW_STEP_CONFIRM=false.")


####################################################################################
# (6) GRAVAR FICHEIROS
####################################################################################

write_text(WORKFLOW_ENGINE_PATH, workflow_engine)
write_text(ENV_EXAMPLE_PATH, env_example)


####################################################################################
# (7) VALIDAÇÕES
####################################################################################

final_workflow_engine = read_text(WORKFLOW_ENGINE_PATH)
final_env_example = read_text(ENV_EXAMPLE_PATH)

if "def _pause_before_step_if_enabled(" not in final_workflow_engine:
    raise SystemExit("ERRO: função de pausa não ficou no workflow_engine.py.")

if "WORKFLOW_STEP_CONFIRM" not in final_workflow_engine:
    raise SystemExit("ERRO: variável WORKFLOW_STEP_CONFIRM não ficou no workflow_engine.py.")

if "WORKFLOW_STEP_CONFIRM=false" not in final_env_example:
    raise SystemExit("ERRO: WORKFLOW_STEP_CONFIRM=false não ficou no .env.example.")

print("OK: patch concluído.")
