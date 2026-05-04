from __future__ import annotations

import argparse
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from sap_session import ensure_sap_access_from_env, load_dotenv_manual


REQUEST_REGEX = re.compile(r"^[A-Z0-9]{3,4}K\d{6,}$")
BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in BOOL_TRUE


def _resolve_remote_login_mode() -> str:
    # Por seguranca, o default e manual.
    mode = str(os.getenv("SAP_STMS_REMOTE_LOGIN_MODE", "manual") or "").strip().lower()
    if mode in {"auto", "automatico"}:
        return "auto"
    return "manual"


def _normalize_request(value: str) -> str:
    req = str(value or "").strip().upper().replace(" ", "")
    if not req:
        return ""
    if REQUEST_REGEX.match(req):
        return req
    return ""


def _safe_find(session, sap_id: str):
    try:
        return session.findById(sap_id)
    except Exception:
        return None


def _send_vkey(session, vkey: int, pause_s: float) -> None:
    session.findById("wnd[0]").sendVKey(vkey)
    _ = pause_s


def _status_message(session) -> tuple[str, str]:
    sbar = _safe_find(session, "wnd[0]/sbar")
    if not sbar:
        return "", ""
    msg_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
    msg_text = str(getattr(sbar, "Text", "") or "").strip()
    return msg_type, msg_text


def _force_show_and_focus(session, wnd_id: str) -> None:
    try:
        wnd = session.findById(wnd_id)
        wnd.setFocus()
    except Exception:
        pass


def _is_login_window(session, wnd_id: str) -> bool:
    client_field, user_field, pwd_field, _lang_field = _find_login_fields_dynamic(session, wnd_id)
    return bool(client_field or user_field or pwd_field)


def _has_any_login_window(session) -> bool:
    return any(_is_login_window(session, f"wnd[{i}]") for i in (0, 1, 2, 3))


def _wait_statusbar_after_login_close(
    session,
    *,
    timeout_s: int = 45,
    poll_s: float = 0.5,
) -> tuple[str, str]:
    deadline = time.time() + max(5, int(timeout_s))
    last_type, last_text = _status_message(session)

    while time.time() < deadline:
        if _has_any_login_window(session):
            time.sleep(max(0.2, float(poll_s)))
            continue

        msg_type, msg_text = _status_message(session)
        if msg_type in {"E", "A"}:
            return msg_type, msg_text
        if str(msg_text or "").strip():
            return msg_type, msg_text
        last_type, last_text = msg_type, msg_text
        time.sleep(max(0.2, float(poll_s)))

    return last_type, last_text


def _wait_manual_remote_login(
    session,
    *,
    default_system: str,
    timeout_s: int = 300,
    poll_s: float = 1.0,
) -> bool:
    # Mantem janela visivel e aguarda operador fechar o login remoto.
    deadline = time.time() + max(30, int(timeout_s))
    prompted = False

    while time.time() < deadline:
        login_wnd = ""
        for idx in (0, 1, 2, 3):
            wid = f"wnd[{idx}]"
            if _is_login_window(session, wid):
                login_wnd = wid
                break

        if not login_wnd:
            msg_type, msg_text = _status_message(session)
            if msg_type in {"E", "A"}:
                raise RuntimeError(f"Erro apos login remoto manual: {msg_text or 'sem detalhe na status bar'}")
            return True

        _force_show_and_focus(session, login_wnd)
        if not prompted:
            system = str(default_system or "").strip().upper() or "sistema remoto"
            print(
                "INFO: Login remoto STMS exige preenchimento manual. "
                f"Preencha a credencial para {system} e confirme no SAP."
            )
            prompted = True

        time.sleep(max(0.2, float(poll_s)))

    raise RuntimeError("Timeout a aguardar preenchimento manual do login remoto STMS.")


def _require_no_error(session, action: str) -> None:
    msg_type, msg_text = _status_message(session)
    if msg_type in {"E", "A"}:
        raise RuntimeError(f"{action}: {msg_text or 'erro SAP sem detalhe na status bar'}")


def _press(session, sap_id: str, action: str) -> None:
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"{action}: elemento nao encontrado ({sap_id})")
    obj.press()


def _set_text(session, sap_id: str, value: str, action: str) -> None:
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"{action}: campo nao encontrado ({sap_id})")
    obj.text = value


def _focus(session, sap_id: str) -> None:
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Elemento para foco nao encontrado: {sap_id}")
    obj.setFocus()


def _try_focus_first(session, candidates: list[str]) -> bool:
    for sap_id in candidates:
        obj = _safe_find(session, sap_id)
        if not obj:
            continue
        try:
            obj.setFocus()
            return True
        except Exception:
            continue
    return False


def _iter_nodes(root, max_nodes: int = 7000):
    stack = [root]
    seen = 0
    while stack and seen < max_nodes:
        obj = stack.pop()
        seen += 1
        yield obj
        try:
            child_count = int(obj.Children.Count)
        except Exception:
            child_count = 0
        for idx in range(child_count - 1, -1, -1):
            try:
                stack.append(obj.Children(idx))
            except Exception:
                continue


def _read_text(obj) -> str:
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


def _iter_window_nodes(window_obj, max_nodes: int = 2000):
    stack = [window_obj]
    seen = 0
    while stack and seen < max_nodes:
        obj = stack.pop()
        seen += 1
        yield obj
        try:
            child_count = int(obj.Children.Count)
        except Exception:
            child_count = 0
        for idx in range(child_count - 1, -1, -1):
            try:
                stack.append(obj.Children(idx))
            except Exception:
                continue


def _detect_system_from_login_window(window_obj, fallback_system: str) -> str:
    detected = str(fallback_system or "").strip().upper()
    pattern = re.compile(r"\b(S4[A-Z])\b", re.IGNORECASE)
    for node in _iter_window_nodes(window_obj):
        text = _read_text(node)
        if not text:
            continue
        match = pattern.search(text.upper())
        if match:
            detected = match.group(1).strip().upper()
            break
    return detected


def _dump_popup_map(session, *, reason: str = "") -> str:
    base_dir = BASE_DIR / "cache"
    base_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = base_dir / f"stms_popup_map_{stamp}.txt"

    lines = []
    lines.append(f"# STMS popup map | {datetime.now().isoformat(timespec='seconds')}")
    if reason:
        lines.append(f"# reason: {reason}")

    for idx in (0, 1, 2, 3):
        wnd_id = f"wnd[{idx}]"
        wnd = _safe_find(session, wnd_id)
        if not wnd:
            continue
        wnd_title = ""
        try:
            wnd_title = str(getattr(wnd, "Text", "") or "").strip()
        except Exception:
            wnd_title = ""
        lines.append("")
        lines.append(f"[{wnd_id}] title={wnd_title}")

        for node in _iter_window_nodes(wnd, max_nodes=3000):
            node_id = ""
            node_type = ""
            node_text = ""
            try:
                node_id = str(getattr(node, "Id", "") or "")
            except Exception:
                node_id = ""
            try:
                node_type = str(getattr(node, "Type", "") or "")
            except Exception:
                node_type = ""
            node_text = _read_text(node)
            if node_id:
                lines.append(f"{node_id} | {node_type} | {node_text}")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return str(out_path)


def _find_login_fields_dynamic(session, wnd_id: str):
    wnd = _safe_find(session, wnd_id)
    if not wnd:
        return None, None, None, None

    client_field = None
    user_field = None
    pwd_field = None
    lang_field = None

    for node in _iter_window_nodes(wnd, max_nodes=2500):
        node_id = ""
        try:
            node_id = str(getattr(node, "Id", "") or "").upper()
        except Exception:
            node_id = ""
        if not node_id:
            continue

        if (node_id.endswith("RSYST-MANDT") or "RSYST-MANDT" in node_id) and not client_field:
            client_field = node
        elif (node_id.endswith("RSYST-BNAME") or "RSYST-BNAME" in node_id) and not user_field:
            user_field = node
        elif (node_id.endswith("RSYST-BCODE") or "RSYST-BCODE" in node_id) and not pwd_field:
            pwd_field = node
        elif (node_id.endswith("RSYST-LANGU") or "RSYST-LANGU" in node_id) and not lang_field:
            lang_field = node

    return client_field, user_field, pwd_field, lang_field


def _resolve_login_credentials(system_name: str, client_hint: str) -> tuple[str, str, str, str]:
    system = str(system_name or "").strip().upper()
    client_hint = str(client_hint or "").strip()
    user = str(os.getenv("SAP_USER", "") or "").strip()
    language = str(os.getenv("SAP_LANGUAGE", "PT") or "PT").strip() or "PT"

    def _candidate_client_from_password_env():
        pattern = re.compile(rf"^SAP_PASSWORD_{re.escape(system)}CLNT(\d+)$", re.IGNORECASE)
        for key, value in os.environ.items():
            key_up = str(key or "").strip().upper()
            match = pattern.match(key_up)
            if not match:
                continue
            if str(value or "").strip():
                return match.group(1)
        return ""

    def _candidate_client_from_client_env():
        pattern = re.compile(rf"^SAP_CLIENT_{re.escape(system)}CLNT(\d+)$", re.IGNORECASE)
        for key, value in os.environ.items():
            key_up = str(key or "").strip().upper()
            match = pattern.match(key_up)
            if not match:
                continue
            env_client = str(value or "").strip()
            if env_client:
                return env_client
        return ""

    client = ""
    if client_hint:
        hinted_pwd_key = f"SAP_PASSWORD_{system}CLNT{client_hint}"
        hinted_pwd = str(os.getenv(hinted_pwd_key, "") or "").strip()
        if hinted_pwd:
            client = client_hint

    if not client:
        client = _candidate_client_from_password_env()
    if not client:
        client = _candidate_client_from_client_env()
    if not client:
        client = client_hint or "100"

    password_key = f"SAP_PASSWORD_{system}CLNT{client}"
    password = str(os.getenv(password_key, "") or "").strip()

    if not user:
        raise RuntimeError("Variavel SAP_USER nao definida para login remoto STMS.")
    if not password:
        raise RuntimeError(f"Variavel {password_key} nao definida para login remoto STMS.")

    return user, password, client, language


def _first_field(session, field_ids: list[str]):
    for fid in field_ids:
        obj = _safe_find(session, fid)
        if obj:
            return obj, fid
    return None, ""


def _handle_remote_login_popup_if_any(
    session,
    *,
    default_system: str,
    default_client: str,
    pause_s: float,
    debug_map: bool = False,
) -> bool:
    _ = pause_s
    _ = default_client
    _ = _resolve_remote_login_mode()

    appear_timeout = int(os.getenv("SAP_STMS_REMOTE_LOGIN_APPEAR_TIMEOUT", "25") or "25")
    appear_timeout = max(0, appear_timeout)

    has_login = any(_is_login_window(session, f"wnd[{i}]") for i in (0, 1, 2, 3))
    if not has_login and appear_timeout > 0:
        deadline = time.time() + appear_timeout
        while time.time() < deadline:
            has_login = any(_is_login_window(session, f"wnd[{i}]") for i in (0, 1, 2, 3))
            if has_login:
                break
            time.sleep(0.5)

    if not has_login:
        if debug_map:
            out = _dump_popup_map(session, reason="handle_remote_login_popup_if_any:no_action")
            print(f"DEBUG: mapa popup gerado em: {out}")
        return False

    if debug_map:
        out = _dump_popup_map(session, reason="manual_login_detected")
        print(f"DEBUG: mapa popup gerado em: {out}")

    _wait_manual_remote_login(
        session,
        default_system=default_system,
        timeout_s=int(os.getenv("SAP_STMS_REMOTE_LOGIN_TIMEOUT", "300") or "300"),
        poll_s=1.0,
    )
    msg_type, msg_text = _wait_statusbar_after_login_close(
        session,
        timeout_s=int(os.getenv("SAP_STMS_STATUSBAR_TIMEOUT", "45") or "45"),
        poll_s=0.5,
    )
    if msg_type in {"E", "A"}:
        raise RuntimeError(f"Erro apos login remoto manual: {msg_text or 'sem detalhe na status bar'}")
    return True


def _focus_request_visible_in_list(session, request_number: str) -> bool:
    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        return False
    target = _normalize_request(request_number)
    if not target:
        return False

    for obj in _iter_nodes(root):
        try:
            obj_id = str(getattr(obj, "Id", "") or "")
        except Exception:
            obj_id = ""
        if "/lbl[" not in obj_id:
            continue

        txt = _read_text(obj).upper().replace(" ", "")
        if txt != target:
            continue
        try:
            obj.setFocus()
            return True
        except Exception:
            continue

    return False


def _focus_selection_cell_for_request(session, request_number: str) -> bool:
    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        return False
    target = _normalize_request(request_number)
    if not target:
        return False

    req_obj_id = ""
    for obj in _iter_nodes(root):
        try:
            obj_id = str(getattr(obj, "Id", "") or "")
        except Exception:
            obj_id = ""
        if "/lbl[" not in obj_id:
            continue

        txt = _read_text(obj).upper().replace(" ", "")
        if txt != target:
            continue
        req_obj_id = obj_id
        break

    if not req_obj_id:
        return False

    match = re.search(r"/lbl\[(\d+),(\d+)\]$", req_obj_id)
    if not match:
        return False

    row = match.group(2)
    candidate_cols = ["11", "10", "12", "9"]
    for col in candidate_cols:
        candidate_id = f"wnd[0]/usr/lbl[{col},{row}]"
        obj = _safe_find(session, candidate_id)
        if not obj:
            continue
        try:
            obj.setFocus()
            return True
        except Exception:
            continue

    # Fallback no proprio label da request.
    req_candidate = _safe_find(session, req_obj_id.replace("/app/con[2]/ses[0]/", ""))
    if req_candidate:
        try:
            req_candidate.setFocus()
            return True
        except Exception:
            return False
    return False


def _is_no_order_selected_message(text: str) -> bool:
    msg = str(text or "").strip().lower()
    return (
        "nenhuma ordem de transporte selecionada" in msg
        or "no transport request selected" in msg
    )


def _is_transport_confirmation_message(text: str, request_number: str) -> bool:
    msg = str(text or "").strip().lower()
    req = str(request_number or "").strip().upper()
    if not msg:
        return False

    negative_tokens = (
        "nenhuma ordem de transporte selecionada",
        "no transport request selected",
        "erro",
        "error",
        "falha",
        "failed",
        "not authorized",
        "nao autorizado",
        "não autorizado",
    )
    if any(tok in msg for tok in negative_tokens):
        return False

    positive_tokens = (
        "import",
        "transport",
        "transfer",
        "iniciad",
        "agend",
        "colocad",
        "incluid",
        "queued",
        "scheduled",
        "execut",
    )
    has_positive = any(tok in msg for tok in positive_tokens)
    has_request = req and (req.lower() in msg)
    # Sucesso comum no STMS: "Importação para o sistema S4Q executada"
    explicit_import_success = ("import" in msg and "execut" in msg)
    if explicit_import_success:
        return True
    return bool(has_positive and (has_request or "request" in msg or "ordem" in msg or "solicit" in msg))


def _find_request_label_id(session, request_number: str) -> str:
    root = _safe_find(session, "wnd[0]/usr")
    if not root:
        return ""
    target = _normalize_request(request_number)
    if not target:
        return ""

    for obj in _iter_nodes(root):
        try:
            obj_id = str(getattr(obj, "Id", "") or "")
        except Exception:
            obj_id = ""
        if "/lbl[" not in obj_id:
            continue
        txt = _read_text(obj).upper().replace(" ", "")
        if txt == target:
            return obj_id
    return ""


def _select_request_for_import(session, request_number: str, *, enter_delay_s: float) -> bool:
    req = _normalize_request(request_number)
    if not req:
        return False

    label_id = _find_request_label_id(session, req)
    if not label_id:
        return False

    row_match = re.search(r"/lbl\[(\d+),(\d+)\]$", label_id)
    if not row_match:
        return False
    row = row_match.group(2)

    # Tenta focar a coluna de seleÃ§Ã£o da mesma linha e marcar com vKey 9.
    candidate_cols = ["23", "24", "11", "10", "2", "3", "1", "0"]
    for col in candidate_cols:
        cid = f"wnd[0]/usr/lbl[{col},{row}]"
        obj = _safe_find(session, cid)
        if not obj:
            continue
        try:
            obj.setFocus()
        except Exception:
            continue
        try:
            _send_vkey(session, 9, pause_s=enter_delay_s)
        except Exception:
            continue
        _msg_type, msg = _status_message(session)
        if not _is_no_order_selected_message(msg):
            return True

    # Fallback por menu "Marcar solicitacao +/-"
    try:
        menu_mark = _safe_find(session, "wnd[0]/mbar/menu[1]/menu[1]/menu[0]")
        if menu_mark:
            menu_mark.select()
            pass
        _msg_type, msg = _status_message(session)
        if not _is_no_order_selected_message(msg):
            return True
    except Exception:
        pass

    # Fallback por botÃ£o da toolbar que normalmente marca a linha corrente.
    try:
        btn_mark = _safe_find(session, "wnd[0]/tbar[1]/btn[9]")
        if btn_mark:
            btn_mark.press()
            pass
            _msg_type, msg = _status_message(session)
            if not _is_no_order_selected_message(msg):
                return True
    except Exception:
        pass

    return False


def _try_transfer_via_menu_fallback(
    session,
    *,
    request_number: str,
    target_client: str,
    system_name: str,
    pause_s: float,
    debug_popup_map: bool,
) -> bool:
    # Ordem > Transferir > Mandante
    menu_path = "wnd[0]/mbar/menu[3]/menu[2]/menu[1]"
    menu = _safe_find(session, menu_path)
    if not menu:
        return False
    try:
        menu.select()
    except Exception:
        return False
    _ = pause_s

    # Alguns layouts pedem request aqui.
    request_fields = [
        "wnd[1]/usr/ctxtSO_TRKOR-LOW",
        "wnd[1]/usr/ctxtTRKORR-LOW",
        "wnd[1]/usr/ctxtTRKORR",
    ]
    for fld in request_fields:
        obj = _safe_find(session, fld)
        if not obj:
            continue
        obj.text = request_number
        ok = _safe_find(session, "wnd[1]/tbar[0]/btn[0]")
        if ok:
            ok.press()
            pass
        break

    _handle_remote_login_popup_if_any(
        session,
        default_system=system_name,
        default_client=target_client,
        pause_s=pause_s,
        debug_map=debug_popup_map,
    )

    client_fields = [
        "wnd[1]/usr/ctxtWTMSU-CLIENT",
        "wnd[1]/usr/txtWTMSU-CLIENT",
        "wnd[1]/usr/ctxtWTMSU-MANDT",
        "wnd[1]/usr/txtWTMSU-MANDT",
    ]
    for fld in client_fields:
        obj = _safe_find(session, fld)
        if not obj:
            continue
        obj.text = target_client
        ok = _safe_find(session, "wnd[1]/tbar[0]/btn[0]")
        if ok:
            ok.press()
            pass
        break

    confirm_btn = _safe_find(session, "wnd[2]/usr/btnBUTTON_1")
    if confirm_btn:
        confirm_btn.press()
        pass

    _handle_remote_login_popup_if_any(
        session,
        default_system=system_name,
        default_client=target_client,
        pause_s=pause_s,
        debug_map=debug_popup_map,
    )

    _msg_type, msg = _status_message(session)
    return (not _is_no_order_selected_message(msg)) and _is_transport_confirmation_message(msg, request_number)


def _reset_home(session, *, enter_delay_s: float) -> None:
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        _send_vkey(session, 0, pause_s=enter_delay_s)
    except Exception:
        return


def executar(
    request_number: str = "",
    *,
    system_name: str = "",
    client: str = "",
    target_client: str = "",
    chamado_pelo_main: bool = False,
    enter_delay_s: float = 0.35,
    debug_popup_map: bool = False,
) -> bool:
    load_dotenv_manual()

    req = _normalize_request(request_number)
    if not req:
        print("INFO: Request nao informada/valida. Nada para transportar.")
        return True

    key = str(os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100") or "").strip().upper() or "S4DCLNT100"
    if system_name:
        os.environ["WORKFLOW_SAP_SYSTEM"] = str(system_name).strip().upper()
    if client:
        os.environ["WORKFLOW_SAP_CLIENT"] = str(client).strip()
    if chamado_pelo_main or _to_bool(os.getenv("SAP_CALLED_BY_MAIN", "false")):
        os.environ["SAP_CALLED_BY_MAIN"] = "1"

    target = str(target_client or client or "").strip()
    if not target:
        target = "100"

    session = ensure_sap_access_from_env(key=key, timeout_s=40, load_env=False)

    try:
        print(f"INFO: A transportar request {req} via STMS para cliente {target}...")

        session.findById("wnd[0]").resizeWorkingPane(92, 26, False)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nSTMS"
        _send_vkey(session, 0, pause_s=enter_delay_s)
        _require_no_error(session, "Falha ao abrir STMS")

        _press(session, "wnd[0]/tbar[1]/btn[5]", "Falha ao abrir fila de importacao")
        pass

        _try_focus_first(
            session,
            [
                "wnd[0]/usr/lbl[11,8]",
                "wnd[0]/usr/lbl[10,8]",
                "wnd[0]/usr/lbl[12,8]",
                "wnd[0]/usr/lbl[11,7]",
                "wnd[0]/usr/lbl[11,9]",
            ],
        )
        _send_vkey(session, 2, pause_s=enter_delay_s)

        _try_focus_first(
            session,
            [
                "wnd[0]/usr/lbl[11,10]",
                "wnd[0]/usr/lbl[10,10]",
                "wnd[0]/usr/lbl[12,10]",
                "wnd[0]/usr/lbl[11,9]",
                "wnd[0]/usr/lbl[11,11]",
            ],
        )
        if not _focus_request_visible_in_list(session, req):
            _send_vkey(session, 34, pause_s=enter_delay_s)
            popup_fields = [
                "wnd[1]/usr/ctxtSO_TRKOR-LOW",
                "wnd[1]/usr/ctxtTRKORR-LOW",
                "wnd[1]/usr/ctxtTRKORR",
            ]
            filled = False
            for fld in popup_fields:
                obj = _safe_find(session, fld)
                if not obj:
                    continue
                obj.text = req
                filled = True
                break
            if not filled:
                raise RuntimeError("Falha ao preencher request na pesquisa: campo de request nao encontrado.")
            _press(session, "wnd[1]/tbar[0]/btn[0]", "Falha ao confirmar pesquisa da request")
            pass
            _focus_request_visible_in_list(session, req)

        marked = _select_request_for_import(session, req, enter_delay_s=enter_delay_s)
        if not marked:
            _focus_selection_cell_for_request(session, req)
            try:
                _send_vkey(session, 9, pause_s=enter_delay_s)
            except Exception:
                pass

        try:
            _send_vkey(session, 35, pause_s=enter_delay_s)
        except Exception:
            pass
        _handle_remote_login_popup_if_any(
            session,
            default_system=system_name or str(getattr(session.Info, "SystemName", "") or ""),
            default_client=target,
            pause_s=enter_delay_s,
            debug_map=debug_popup_map,
        )

        client_fields = [
            "wnd[1]/usr/ctxtWTMSU-CLIENT",
            "wnd[1]/usr/txtWTMSU-CLIENT",
            "wnd[1]/usr/ctxtWTMSU-MANDT",
            "wnd[1]/usr/txtWTMSU-MANDT",
        ]
        client_popup_present = _safe_find(session, "wnd[1]") is not None
        if client_popup_present:
            filled_client = False
            for fld in client_fields:
                obj = _safe_find(session, fld)
                if not obj:
                    continue
                obj.text = target
                filled_client = True
                break
            if not filled_client:
                raise RuntimeError("Falha ao preencher cliente destino: campo de cliente nao encontrado.")
            _press(session, "wnd[1]/tbar[0]/btn[0]", "Falha ao confirmar cliente destino")
            pass
            _handle_remote_login_popup_if_any(
                session,
                default_system=system_name or str(getattr(session.Info, "SystemName", "") or ""),
                default_client=target,
                pause_s=enter_delay_s,
                debug_map=debug_popup_map,
            )

        confirm_btn = _safe_find(session, "wnd[2]/usr/btnBUTTON_1")
        if confirm_btn:
            confirm_btn.press()
            pass
            _handle_remote_login_popup_if_any(
                session,
                default_system=system_name or str(getattr(session.Info, "SystemName", "") or ""),
                default_client=target,
                pause_s=enter_delay_s,
                debug_map=debug_popup_map,
            )

        _require_no_error(session, "Erro apos transporte da request")
        msg_type, msg_text = _status_message(session)
        if not str(msg_text or "").strip():
            msg_type, msg_text = _wait_statusbar_after_login_close(
                session,
                timeout_s=int(os.getenv("SAP_STMS_STATUSBAR_TIMEOUT", "45") or "45"),
                poll_s=0.5,
            )
        msg_norm = (msg_text or "").strip().lower()
        confirmed = _is_transport_confirmation_message(msg_text, req)
        if "nenhuma ordem de transporte selecionada" in msg_norm:
            ok_fallback = _try_transfer_via_menu_fallback(
                session,
                request_number=req,
                target_client=target,
                system_name=system_name or str(getattr(session.Info, "SystemName", "") or ""),
                pause_s=enter_delay_s,
                debug_popup_map=debug_popup_map,
            )
            _require_no_error(session, "Erro apos fallback de transporte por menu")
            msg_type, msg_text = _status_message(session)
            confirmed = _is_transport_confirmation_message(msg_text, req) or bool(ok_fallback)
            if not ok_fallback and _is_no_order_selected_message(msg_text):
                raise RuntimeError(msg_text)
        if not confirmed:
            raise RuntimeError(
                "Transporte sem confirmacao explicita na status bar. "
                f"Mensagem SAP='{msg_text or '(vazia)'}'"
            )
        if msg_text:
            print(f"INFO: Status SAP apos transporte ({msg_type or '-'}): {msg_text}")
        print(f"REQUEST_TRANSPORTED={req}")
        return True
    finally:
        _reset_home(session, enter_delay_s=enter_delay_s)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", default="", help="Numero da request a transportar.")
    parser.add_argument("--system-name", default="", help="Sistema SAP alvo (opcional).")
    parser.add_argument("--client", default="", help="Mandante SAP da sessao atual (opcional).")
    parser.add_argument("--target-client", default="", help="Mandante destino na importacao STMS.")
    parser.add_argument("--from-main", action="store_true")
    parser.add_argument(
        "--enter-delay-seconds",
        type=float,
        default=0.0,
        help="(Legado) sem efeito para passos SAP; espera ocorre apenas no popup manual de credenciais.",
    )
    parser.add_argument(
        "--debug-popup-map",
        action="store_true",
        help="Gera dump de mapeamento dos popups/telas STMS quando nao agir no login remoto.",
    )
    args = parser.parse_args()

    try:
        ok = executar(
            request_number=args.request,
            system_name=args.system_name,
            client=args.client,
            target_client=args.target_client,
            chamado_pelo_main=bool(args.from_main),
            enter_delay_s=max(0.0, float(args.enter_delay_seconds)),
            debug_popup_map=bool(args.debug_popup_map),
        )
        return 0 if ok else 1
    except Exception as exc:
        print(f"ERRO: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
