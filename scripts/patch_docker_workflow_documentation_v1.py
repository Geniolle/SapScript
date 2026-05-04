from __future__ import annotations

####################################################################################
# (1) IMPORTS
####################################################################################

from pathlib import Path
from datetime import datetime


####################################################################################
# (2) CAMINHOS
####################################################################################

ROOT = Path.cwd()
DOCKERFILE_PATH = ROOT / "Dockerfile"


####################################################################################
# (3) FUNÇÕES AUXILIARES
####################################################################################

def backup_file_v1(path: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(path.suffix + f".bak_{timestamp}")
    backup_path.write_text(path.read_text(encoding="utf-8-sig"), encoding="utf-8")
    print(f"OK: backup criado: {backup_path}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


####################################################################################
# (4) VALIDAR FICHEIROS
####################################################################################

if not DOCKERFILE_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {DOCKERFILE_PATH}")

if not (ROOT / "workflow_documentation.py").exists():
    raise SystemExit("ERRO: workflow_documentation.py não existe na raiz do projeto.")

backup_file_v1(DOCKERFILE_PATH)


####################################################################################
# (5) ALTERAR DOCKERFILE
####################################################################################

content = read_text(DOCKERFILE_PATH)

if "workflow_documentation.py" in content:
    print("INFO: Dockerfile já contém workflow_documentation.py. Nada a alterar.")
else:
    old_line = "COPY main.py workflow_engine.py workflows.json jira_download_anexos.py jira_sheet_daemon.py sap_session.py ./"
    new_line = "COPY main.py workflow_engine.py workflow_documentation.py workflows.json jira_download_anexos.py jira_sheet_daemon.py sap_session.py ./"

    if old_line not in content:
        raise SystemExit(
            "ERRO: linha COPY esperada não encontrada no Dockerfile. "
            "Revise o Dockerfile manualmente."
        )

    content = content.replace(old_line, new_line)
    print("OK: Dockerfile atualizado para copiar workflow_documentation.py.")

write_text(DOCKERFILE_PATH, content)


####################################################################################
# (6) VALIDAR RESULTADO
####################################################################################

final_content = read_text(DOCKERFILE_PATH)

required_files = [
    "main.py",
    "workflow_engine.py",
    "workflow_documentation.py",
    "workflows.json",
    "jira_download_anexos.py",
    "jira_sheet_daemon.py",
    "sap_session.py",
]

missing = [name for name in required_files if name not in final_content]

if missing:
    raise SystemExit("ERRO: ficheiros ainda ausentes no Dockerfile: " + ", ".join(missing))

print("OK: validação concluída. Dockerfile contém todos os módulos Python necessários para import inicial.")
