from __future__ import annotations

####################################################################################
# (1) IMPORTS
####################################################################################

from pathlib import Path
from datetime import datetime
import re


####################################################################################
# (2) CAMINHOS
####################################################################################

ROOT = Path.cwd()

DOCKERFILE_PATH = ROOT / "Dockerfile"
COMPOSE_PATH = ROOT / "docker-compose.yml"
DAEMON_PATH = ROOT / "jira_sheet_daemon.py"


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


def require_file(path: Path) -> None:
    if not path.exists():
        raise SystemExit(f"ERRO: ficheiro não encontrado: {path}")


####################################################################################
# (4) VALIDAR FICHEIROS
####################################################################################

for file_path in [DOCKERFILE_PATH, COMPOSE_PATH, DAEMON_PATH]:
    require_file(file_path)

for file_path in [DOCKERFILE_PATH, COMPOSE_PATH, DAEMON_PATH]:
    backup_file_v1(file_path)


####################################################################################
# (5) CORRIGIR DOCKERFILE: COPIAR sap_session.py PARA O CONTAINER
####################################################################################

dockerfile = read_text(DOCKERFILE_PATH)

if "sap_session.py" in dockerfile:
    print("INFO: Dockerfile já contém sap_session.py. Nada a alterar nessa parte.")
else:
    old_copy = "COPY main.py workflow_engine.py workflows.json jira_download_anexos.py jira_sheet_daemon.py ./"
    new_copy = "COPY main.py workflow_engine.py workflows.json jira_download_anexos.py jira_sheet_daemon.py sap_session.py ./"

    if old_copy not in dockerfile:
        raise SystemExit(
            "ERRO: linha COPY esperada não encontrada no Dockerfile. "
            "Revise manualmente antes de aplicar o patch."
        )

    dockerfile = dockerfile.replace(old_copy, new_copy)
    print("OK: Dockerfile atualizado para copiar sap_session.py.")

write_text(DOCKERFILE_PATH, dockerfile)


####################################################################################
# (6) CORRIGIR jira_sheet_daemon.py: PASSAR ESTADO_ALVO PARA filtrar_linhas()
####################################################################################

daemon = read_text(DAEMON_PATH)

if "sheet_main.ESTADO_ALVO" in daemon:
    print("INFO: jira_sheet_daemon.py já passa ESTADO_ALVO. Nada a alterar nessa parte.")
else:
    old_call = """linhas = sheet_main.filtrar_linhas(
        dados,
        sheet_main.RESPONSAVEL_ALVO,
        sheet_main.SUPPLIER_ALVO,
    )"""

    new_call = """linhas = sheet_main.filtrar_linhas(
        dados,
        sheet_main.RESPONSAVEL_ALVO,
        sheet_main.SUPPLIER_ALVO,
        sheet_main.ESTADO_ALVO,
    )"""

    if old_call not in daemon:
        raise SystemExit(
            "ERRO: chamada antiga de filtrar_linhas() não encontrada em jira_sheet_daemon.py. "
            "Revise manualmente antes de aplicar o patch."
        )

    daemon = daemon.replace(old_call, new_call)
    print("OK: jira_sheet_daemon.py atualizado para passar ESTADO_ALVO.")

write_text(DAEMON_PATH, daemon)


####################################################################################
# (7) CORRIGIR docker-compose.yml: EXPOR ESTADO_ALVO NO CONTAINER
####################################################################################

compose = read_text(COMPOSE_PATH)

if "ESTADO_ALVO:" in compose:
    print("INFO: docker-compose.yml já contém ESTADO_ALVO. Nada a alterar nessa parte.")
else:
    old_line = "      SUPPLIER_ALVO: ${SUPPLIER_ALVO:-Evolutive}"
    new_block = """      SUPPLIER_ALVO: ${SUPPLIER_ALVO:-Evolutive}
      ESTADO_ALVO: ${ESTADO_ALVO:-In Review}"""

    if old_line not in compose:
        raise SystemExit(
            "ERRO: linha SUPPLIER_ALVO não encontrada no docker-compose.yml. "
            "Revise manualmente antes de aplicar o patch."
        )

    compose = compose.replace(old_line, new_block)
    print("OK: docker-compose.yml atualizado com ESTADO_ALVO.")

write_text(COMPOSE_PATH, compose)


####################################################################################
# (8) VALIDAÇÕES DE CONTEÚDO
####################################################################################

dockerfile_final = read_text(DOCKERFILE_PATH)
daemon_final = read_text(DAEMON_PATH)
compose_final = read_text(COMPOSE_PATH)

if "sap_session.py" not in dockerfile_final:
    raise SystemExit("ERRO: sap_session.py não ficou no Dockerfile.")

if "sheet_main.ESTADO_ALVO" not in daemon_final:
    raise SystemExit("ERRO: ESTADO_ALVO não ficou na chamada de filtrar_linhas().")

if "ESTADO_ALVO:" not in compose_final:
    raise SystemExit("ERRO: ESTADO_ALVO não ficou no docker-compose.yml.")

if "SUPPLIER_ALVO: ${SUPPLIER_ALVO:-Evolutive}" not in compose_final:
    raise SystemExit("ERRO: SUPPLIER_ALVO padrão Evolutive não foi encontrado no compose.")

if "SHEET_NAME: ${SHEET_NAME:-DADOS}" not in compose_final:
    raise SystemExit("ERRO: SHEET_NAME padrão DADOS não foi encontrado no compose.")

print("OK: validações do patch concluídas.")
print("OK: Docker deverá ler Sheet=DADOS, Supplier=Evolutive e Estado=In Review.")
