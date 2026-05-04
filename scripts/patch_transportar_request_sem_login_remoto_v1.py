from __future__ import annotations

from pathlib import Path
from datetime import datetime


ROOT = Path.cwd()
TARGET_PATH = ROOT / "Processos" / "transportar_request.py"


def backup_file(path: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(path.suffix + f".bak_{timestamp}")
    backup_path.write_text(path.read_text(encoding="utf-8-sig"), encoding="utf-8")
    print(f"OK: backup criado: {backup_path}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


if not TARGET_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {TARGET_PATH}")

backup_file(TARGET_PATH)

content = read_text(TARGET_PATH)

start_marker = "def _handle_remote_login_popup_if_any("
end_marker = "\ndef _focus_request_visible_in_list"

start = content.find(start_marker)
if start == -1:
    raise SystemExit("ERRO: função _handle_remote_login_popup_if_any não encontrada.")

end = content.find(end_marker, start)
if end == -1:
    raise SystemExit("ERRO: marcador da próxima função não encontrado.")

new_function = '''def _handle_remote_login_popup_if_any(
    session,
    *,
    default_system: str,
    default_client: str,
    pause_s: float,
    debug_map: bool = False,
) -> bool:
    """
    Tratamento remoto STMS desativado.

    Regra temporária de validação:
      - não detectar janela de login remoto
      - não focar popup
      - não aguardar timeout
      - não preencher credenciais
      - não interagir com janela adicional

    Nesta etapa, o script apenas consulta wnd[0]/sbar para obter
    a mensagem real do SAP e segue o fluxo normal de validação.
    """
    _ = default_system
    _ = default_client
    _ = pause_s
    _ = debug_map

    msg_type, msg_text = _status_message(session)
    if msg_text:
        print(f"INFO: STMS statusbar | Tipo={msg_type or '-'} | MSG={msg_text}")

    return False
'''

content = content[:start] + new_function + content[end:]

if "SAP_STMS_REMOTE_LOGIN_APPEAR_TIMEOUT" in content:
    raise SystemExit("ERRO: ainda existe uso de SAP_STMS_REMOTE_LOGIN_APPEAR_TIMEOUT no bloco ativo.")

if "Timeout a aguardar preenchimento manual do login remoto STMS" in content:
    print("INFO: helper antigo ainda existe no ficheiro, mas não é chamado pela função principal.")

if "def _handle_remote_login_popup_if_any(" not in content:
    raise SystemExit("ERRO: função final não ficou no ficheiro.")

if "Tratamento remoto STMS desativado." not in content:
    raise SystemExit("ERRO: função nova não foi aplicada.")

if "return False" not in content:
    raise SystemExit("ERRO: função nova não retorna False.")

write_text(TARGET_PATH, content)

print("OK: tratamento de popup/login remoto STMS foi desativado em transportar_request.py.")
print("OK: a função agora apenas lê wnd[0]/sbar e não faz interação adicional.")
