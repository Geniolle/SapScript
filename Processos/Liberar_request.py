from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from sap_session import ensure_sap_access_from_env, load_dotenv_manual


REQUEST_REGEX = re.compile(r"^[A-Z0-9]{3,4}K\d{6,}$")
REQUEST_IN_TEXT_REGEX = re.compile(r"\b([A-Z0-9]{3,4}K\d{6,})\b")
BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in BOOL_TRUE


def _normalize_request(value: str) -> str:
    req = str(value or "").strip().upper().replace(" ", "")
    if not req:
        return ""
    if REQUEST_REGEX.match(req):
        return req
    return ""


def _extract_requests_from_text(value: str) -> list[str]:
    text = str(value or "").strip().upper()
    if not text:
        return []
    found = []
    for match in REQUEST_IN_TEXT_REGEX.findall(text):
        req = _normalize_request(match)
        if req and req not in found:
            found.append(req)
    return found


def _safe_find(session, sap_id: str):
    try:
        return session.findById(sap_id)
    except Exception:
        return None


def _get_status(session) -> tuple[str, str]:
    sbar = _safe_find(session, "wnd[0]/sbar")
    if not sbar:
        return "", ""
    msg_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
    msg_text = str(getattr(sbar, "Text", "") or "").strip()
    return msg_type, msg_text


def _try_get_text(obj) -> str:
    for attr in ("Text", "text", "Value", "value", "Key", "key"):
        try:
            value = getattr(obj, attr)
            if value is None:
                continue
            text = str(value).strip()
            if text:
                return text
        except Exception:
            continue
    return ""


def _iter_nodes(root, max_nodes: int = 9000):
    stack = [root]
    seen = 0
    while stack and seen < max_nodes:
        node = stack.pop()
        seen += 1
        yield node
        try:
            child_count = int(node.Children.Count)
        except Exception:
            child_count = 0
        for idx in range(child_count - 1, -1, -1):
            try:
                stack.append(node.Children(idx))
            except Exception:
                continue


def _send_vkey(session, vkey: int, pause_s: float = 0.35) -> None:
    session.findById("wnd[0]").sendVKey(vkey)
    time.sleep(pause_s)


def _reset_to_home(session, *, enter_delay_s: float = 0.35) -> None:
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        _send_vkey(session, 0, pause_s=enter_delay_s)
    except Exception:
        return


def _select_release_target(session, request_number: str, *, required: bool = True) -> bool:
    field = _safe_find(session, "wnd[0]/usr/ctxtP_KORR-LOW")
    if not field:
        if required:
            raise RuntimeError("Campo P_KORR-LOW nao encontrado na SE10.")
        return False
    field.text = request_number
    try:
        field.caretPosition = len(request_number)
    except Exception:
        pass

    btn_execute = _safe_find(session, "wnd[0]/tbar[1]/btn[8]")
    if not btn_execute:
        if required:
            raise RuntimeError("Botao Executar nao encontrado na SE10.")
        return False
    btn_execute.press()
    time.sleep(0.5)
    return True


def _focus_request_in_tree(session, request_number: str) -> bool:
    request = _normalize_request(request_number)
    if not request:
        return False

    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        root = _safe_find(session, "wnd[0]")
    if not root:
        return False

    for obj in _iter_nodes(root):
        text = _try_get_text(obj)
        if not text:
            continue
        found = _extract_requests_from_text(text)
        if request not in found:
            continue
        try:
            obj.setFocus()
        except Exception:
            continue
        try:
            obj.caretPosition = len(request)
        except Exception:
            pass
        return True

    return False


def _focus_any_order_or_task(session, *, preferred_prefix: str = "") -> bool:
    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        root = _safe_find(session, "wnd[0]")
    if not root:
        return False

    candidates = []
    for obj in _iter_nodes(root):
        text = _try_get_text(obj)
        if not text:
            continue
        found = _extract_requests_from_text(text)
        for req in found:
            candidates.append((req, obj))

    if preferred_prefix:
        preferred = [(req, obj) for req, obj in candidates if req.startswith(preferred_prefix)]
        if preferred:
            candidates = preferred

    for _req, obj in candidates:
        try:
            obj.setFocus()
            return True
        except Exception:
            continue

    # Fallback por labels comuns quando a leitura de texto da arvore falha.
    for sap_id in ("wnd[0]/usr/lbl[24,5]", "wnd[0]/usr/lbl[20,9]", "wnd[0]/usr/lbl[16,7]", "wnd[0]/usr/lbl[11,10]"):
        obj = _safe_find(session, sap_id)
        if not obj:
            continue
        try:
            obj.setFocus()
            return True
        except Exception:
            continue

    return False


def _release_current_focus(session, *, enter_delay_s: float = 0.35) -> tuple[str, str]:
    _send_vkey(session, 9, pause_s=enter_delay_s)
    return _get_status(session)


def _confirm_popup_if_any(session, *, enter_delay_s: float = 0.35) -> None:
    popup = _safe_find(session, "wnd[1]")
    if not popup:
        return
    for btn_id in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[11]", "wnd[1]/tbar[0]/btn[12]"):
        btn = _safe_find(session, btn_id)
        if not btn:
            continue
        try:
            btn.press()
            time.sleep(enter_delay_s)
            return
        except Exception:
            continue
    try:
        popup.sendVKey(0)
        time.sleep(enter_delay_s)
    except Exception:
        return


def _has_pending_subtask_message(message: str) -> bool:
    msg = str(message or "").strip().lower()
    if not msg:
        return False
    tokens = (
        "ainda não liberad",
        "ainda nao liberad",
        "not yet released",
        "still not released",
        "tarefas",
        "tasks",
    )
    return any(token in msg for token in tokens)


def _has_position_cursor_message(message: str) -> bool:
    msg = str(message or "").strip().lower()
    if not msg:
        return False
    return ("posicionar o cursor" in msg) or ("position the cursor" in msg)


def _request_sort_key(request_number: str) -> tuple[str, int]:
    req = _normalize_request(request_number)
    if not req:
        return ("", 0)
    m = re.search(r"(\d+)$", req)
    seq = int(m.group(1)) if m else 0
    return (req[:4], seq)


def _collect_related_requests(session, parent_request: str) -> list[str]:
    parent = _normalize_request(parent_request)
    if not parent:
        return []

    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        root = _safe_find(session, "wnd[0]")
    if not root:
        return [parent]

    prefix = parent[:4]
    found: list[str] = []
    for obj in _iter_nodes(root):
        text = _try_get_text(obj)
        if not text:
            continue
        for req in _extract_requests_from_text(text):
            if req.startswith(prefix) and req not in found:
                found.append(req)

    if parent not in found:
        found.append(parent)
    return found


def _ensure_release_list(session, parent_request: str, *, enter_delay_s: float) -> None:
    if _select_release_target(session, parent_request, required=False):
        return
    _open_se10_release_selector(session, enter_delay_s=enter_delay_s)
    _select_release_target(session, parent_request, required=True)


def _release_specific_request(
    session,
    target_request: str,
    *,
    parent_request: str,
    enter_delay_s: float,
) -> tuple[bool, str]:
    target = _normalize_request(target_request)
    parent = _normalize_request(parent_request)
    if not target:
        return False, "Request alvo invalida."

    if not _focus_request_in_tree(session, target):
        _focus_any_order_or_task(session, preferred_prefix=target[:4])
        if not _focus_request_in_tree(session, target):
            return True, f"Request {target} nao localizada para foco na SE10 (possivelmente ja liberada)."

    try:
        _release_current_focus(session, enter_delay_s=enter_delay_s)
    except Exception as exc:
        if "virtual key is not enabled" in str(exc).lower():
            _ensure_release_list(session, parent, enter_delay_s=enter_delay_s)
            if not _focus_request_in_tree(session, target):
                return False, f"Nao consegui refocar {target} apos retry."
            _release_current_focus(session, enter_delay_s=enter_delay_s)
        else:
            raise

    _confirm_popup_if_any(session, enter_delay_s=enter_delay_s)
    msg_type, msg_text = _get_status(session)

    if msg_type in {"E", "A"}:
        return False, msg_text or "Erro SAP sem detalhe."

    msg_norm = str(msg_text or "").strip().lower()
    mentioned = _extract_requests_from_text(msg_text)

    if target in mentioned:
        return True, msg_text
    if "ja liberad" in msg_norm or "já liberad" in msg_norm:
        return True, msg_text
    if "liberad" in msg_norm and target in (msg_text or ""):
        return True, msg_text
    if _has_position_cursor_message(msg_text):
        return False, msg_text
    if _has_pending_subtask_message(msg_text) and target == parent:
        return False, msg_text

    return True, msg_text


def _open_se10_release_selector(session, *, enter_delay_s: float = 0.35) -> None:
    okcd = _safe_find(session, "wnd[0]/tbar[0]/okcd")
    if not okcd:
        raise RuntimeError("Campo de comando SAP (okcd) nao encontrado.")

    okcd.text = "/nSE10"
    _send_vkey(session, 0, pause_s=enter_delay_s)

    menu = _safe_find(session, "wnd[0]/mbar/menu[0]/menu[3]")
    if not menu:
        raise RuntimeError("Menu de liberacao nao encontrado na SE10.")
    menu.select()
    time.sleep(0.35)


def _validate_no_blocking_error(session, action: str) -> None:
    msg_type, msg_text = _get_status(session)
    if msg_type in {"E", "A"}:
        raise RuntimeError(f"{action}: {msg_text or 'erro SAP sem detalhe na status bar'}")


def executar(
    request_number: str = "",
    *,
    system_name: str = "",
    client: str = "",
    chamado_pelo_main: bool = False,
    enter_delay_s: float = 0.35,
) -> bool:
    load_dotenv_manual()

    req = _normalize_request(request_number)
    if not req:
        print("INFO: Request nao informada/valida. Nada para liberar.")
        return True

    key = str(os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100") or "").strip().upper() or "S4DCLNT100"
    if system_name:
        os.environ["WORKFLOW_SAP_SYSTEM"] = str(system_name).strip().upper()
    if client:
        os.environ["WORKFLOW_SAP_CLIENT"] = str(client).strip()

    # Mantem modo nao interativo quando o fluxo veio do main/workflow.
    if chamado_pelo_main or _to_bool(os.getenv("SAP_CALLED_BY_MAIN", "false")):
        os.environ["SAP_CALLED_BY_MAIN"] = "1"

    session = ensure_sap_access_from_env(key=key, timeout_s=40, load_env=False)

    try:
        print(f"INFO: A liberar request {req} via SE10...")
        _open_se10_release_selector(session, enter_delay_s=enter_delay_s)
        _validate_no_blocking_error(session, "Falha ao abrir menu de liberacao")

        _select_release_target(session, req, required=True)
        _validate_no_blocking_error(session, "Falha ao selecionar request")

        related = _collect_related_requests(session, req)
        subtasks = sorted(
            [item for item in related if item != req],
            key=_request_sort_key,
            reverse=True,
        )
        sequence = subtasks + [req]
        print(f"INFO: Sequencia de liberacao (subtarefas -> principal): {', '.join(sequence)}")

        for target in subtasks:
            ok, message = _release_specific_request(
                session,
                target,
                parent_request=req,
                enter_delay_s=enter_delay_s,
            )
            if not ok:
                raise RuntimeError(f"Falha ao liberar subtarefa {target}: {message}")
            if message:
                print(f"INFO: Subtarefa {target} -> {message}")

        ok_parent, message_parent = _release_specific_request(
            session,
            req,
            parent_request=req,
            enter_delay_s=enter_delay_s,
        )
        if not ok_parent:
            raise RuntimeError(f"Falha ao liberar request principal {req}: {message_parent}")
        if message_parent:
            print(f"INFO: Principal {req} -> {message_parent}")

        print(f"REQUEST_RELEASED={req}")
        return True
    finally:
        _reset_to_home(session, enter_delay_s=enter_delay_s)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", default="", help="Numero da request a liberar (ex: S4DK952988).")
    parser.add_argument("--system-name", default="", help="Sistema SAP alvo (opcional).")
    parser.add_argument("--client", default="", help="Mandante SAP alvo (opcional).")
    parser.add_argument("--from-main", action="store_true")
    parser.add_argument(
        "--enter-delay-seconds",
        type=float,
        default=0.35,
        help="Delay em segundos apos cada Enter (sendVKey).",
    )
    args = parser.parse_args()

    try:
        ok = executar(
            request_number=args.request,
            system_name=args.system_name,
            client=args.client,
            chamado_pelo_main=bool(args.from_main),
            enter_delay_s=max(0.0, float(args.enter_delay_seconds)),
        )
        return 0 if ok else 1
    except Exception as exc:
        print(f"ERRO: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
