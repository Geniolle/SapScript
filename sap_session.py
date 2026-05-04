from __future__ import annotations

import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional, Tuple


DEFAULT_SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}
SECOND_ENTER_DELAY_S = 2.0
WINDOW_MODE_MINIMIZE = {"minimize", "minimizar", "hidden", "hide", "ocultar", "quiet"}
WINDOW_MODE_SHOW = {"show", "mostrar", "visible", "visivel", "exibir"}


@dataclass
class SapTarget:
    key: str
    system_name: str
    connection_name: str
    client: str
    user: str
    password: str
    language: str
    saplogon_path: str


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in BOOL_TRUE


def _resolve_window_mode() -> str:
    mode = str(os.getenv("SAP_WINDOW_MODE", "") or "").strip().lower()
    if mode in WINDOW_MODE_MINIMIZE:
        return "minimize"
    if mode in WINDOW_MODE_SHOW:
        return "show"

    # Compatibilidade retro para um switch booleano.
    if _to_bool(os.getenv("SAP_WINDOW_MINIMIZE", "false")):
        return "minimize"
    return "show"


def apply_window_mode(session, *, mode: str | None = None) -> None:
    selected = (mode or _resolve_window_mode()).strip().lower()

    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return

    def _dock_left_half() -> bool:
        try:
            import win32api  # type: ignore
            import win32con  # type: ignore
            import win32gui  # type: ignore
        except Exception:
            return False

        try:
            hwnd = int(getattr(wnd0, "Handle"))
        except Exception:
            return False

        if not hwnd:
            return False

        try:
            screen_w = int(win32api.GetSystemMetrics(0))
            screen_h = int(win32api.GetSystemMetrics(1))
            target_w = max(900, screen_w // 2)
            target_h = max(700, screen_h)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.MoveWindow(hwnd, 0, 0, target_w, target_h, True)
            return True
        except Exception:
            return False

    try:
        if selected == "minimize":
            wnd0.iconify()
        elif selected == "show":
            # Mantem o SAP visivel no lado esquerdo (metade da tela), sem maximizar.
            _dock_left_half()
    except Exception:
        return


def _parse_env_line(raw: str) -> Tuple[Optional[str], Optional[str]]:
    line = (raw or "").strip()
    if not line or line.startswith("#") or "=" not in line:
        return None, None

    key, value = line.split("=", 1)
    key = key.strip()
    value = value.strip()
    if not key:
        return None, None

    if len(value) >= 2 and (
        (value.startswith('"') and value.endswith('"'))
        or (value.startswith("'") and value.endswith("'"))
    ):
        value = value[1:-1]

    return key, value


def load_dotenv_manual() -> Optional[str]:
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent / ".env",
        Path(__file__).resolve().parent.parent / ".env",
    ]

    seen = set()
    for path in candidates:
        resolved = str(path.resolve())
        if resolved in seen:
            continue
        seen.add(resolved)

        if not path.exists():
            continue

        with open(path, "r", encoding="utf-8-sig") as file_obj:
            for raw in file_obj:
                key, value = _parse_env_line(raw)
                if key and key not in os.environ:
                    os.environ[key] = value or ""

        return resolved

    return None


def _derive_system_from_key(key: str) -> str:
    upper = str(key or "").strip().upper()
    if "CLNT" in upper:
        return upper.split("CLNT", 1)[0].strip().upper()
    return upper


def resolve_sap_target_from_env(key: str | None = None) -> SapTarget:
    workflow_key = (key or os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100")).strip().upper()
    if not workflow_key:
        workflow_key = "S4DCLNT100"

    system_name = os.getenv("WORKFLOW_SAP_SYSTEM", "").strip().upper()
    if not system_name:
        system_name = _derive_system_from_key(workflow_key)

    connection_name = os.getenv(f"SAP_CONNECTION_{workflow_key}", "").strip()
    client = os.getenv("WORKFLOW_SAP_CLIENT", "").strip()
    if not client:
        client = os.getenv(f"SAP_CLIENT_{workflow_key}", "").strip()
    if not client:
        client = os.getenv("SAP_CLIENT", "").strip()

    user = os.getenv("SAP_USER", "").strip()
    password = os.getenv(f"SAP_PASSWORD_{workflow_key}", "").strip()
    language = os.getenv("SAP_LANGUAGE", "PT").strip() or "PT"
    saplogon_path = os.getenv("SAPLOGON_PATH", DEFAULT_SAPLOGON_PATH).strip() or DEFAULT_SAPLOGON_PATH

    return SapTarget(
        key=workflow_key,
        system_name=system_name,
        connection_name=connection_name,
        client=client,
        user=user,
        password=password,
        language=language,
        saplogon_path=saplogon_path,
    )


def _import_sap_com():
    if os.name != "nt":
        raise RuntimeError("A automacao SAP GUI requer Windows.")
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("Dependencias pywin32 nao estao disponiveis.") from exc
    return pythoncom, win32com.client


def _iter_sessions(application) -> Iterator[Tuple[object, object]]:
    for i in range(application.Children.Count):
        conn = application.Children(i)
        for j in range(conn.Children.Count):
            sess = conn.Children(j)
            yield conn, sess


def _find_logged_session(application, *, system_name: str, client: str):
    expected_system = str(system_name or "").strip().upper()
    expected_client = str(client or "").strip()

    for _conn, sess in _iter_sessions(application):
        try:
            sess_system = str(sess.Info.SystemName or "").strip().upper()
            sess_client = str(sess.Info.Client or "").strip()
            sess_user = str(sess.Info.User or "").strip()
        except Exception:
            continue

        if expected_system and sess_system != expected_system:
            continue
        if expected_client and sess_client != expected_client:
            continue
        if not sess_user:
            continue
        return sess

    return None


def _find_login_session(application):
    for _conn, sess in _iter_sessions(application):
        try:
            sess.findById("wnd[0]/usr/txtRSYST-BNAME")
            sess.findById("wnd[0]/usr/pwdRSYST-BCODE")
            sess.findById("wnd[0]/usr/txtRSYST-MANDT")
            return sess
        except Exception:
            continue
    return None


def _wait_for_connection_session(connection, timeout_s: int):
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            if connection.Children.Count > 0:
                return connection.Children(0)
        except Exception:
            pass
        time.sleep(0.5)
    return None


def _dismiss_popup_if_any(session) -> None:
    try:
        session.findById("wnd[1]")
    except Exception:
        return

    for btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[11]", "wnd[1]/tbar[0]/btn[12]"):
        try:
            session.findById(btn).press()
            return
        except Exception:
            continue

    try:
        session.findById("wnd[1]").sendVKey(0)
    except Exception:
        pass


def _wait_login_ok(session, target: SapTarget, timeout_s: int) -> Tuple[bool, str]:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            user = str(session.Info.User or "").strip()
            system_name = str(session.Info.SystemName or "").strip().upper()
            client = str(session.Info.Client or "").strip()
            if user and system_name == target.system_name and (not target.client or client == target.client):
                return True, ""
        except Exception:
            pass

        _dismiss_popup_if_any(session)

        try:
            sbar = session.findById("wnd[0]/sbar")
            msg_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
            msg_text = str(getattr(sbar, "Text", "") or "").strip()
            if msg_type in ("E", "A") and msg_text:
                return False, msg_text
        except Exception:
            pass

        time.sleep(0.5)

    return False, "Timeout a aguardar confirmacao de login."


def _try_minimize_saplogon_windows() -> None:
    try:
        import win32con  # type: ignore
        import win32gui  # type: ignore
    except Exception:
        return

    handles = []

    def _enum_window(hwnd, _extra):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = str(win32gui.GetWindowText(hwnd) or "").strip().lower()
            if not title:
                return
            if "sap logon" in title or "saplogon" in title:
                handles.append(hwnd)
        except Exception:
            return

    try:
        win32gui.EnumWindows(_enum_window, None)
        for hwnd in handles:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            except Exception:
                continue
    except Exception:
        return


def _submit_login(session, target: SapTarget) -> None:
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = target.client
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = target.user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = target.password
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = target.language
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(SECOND_ENTER_DELAY_S)
    session.findById("wnd[0]").sendVKey(0)


def _validate_target(target: SapTarget) -> None:
    missing = []
    if not target.connection_name:
        missing.append(f"SAP_CONNECTION_{target.key}")
    if not target.client:
        missing.append(f"SAP_CLIENT_{target.key} (ou SAP_CLIENT)")
    if not target.user:
        missing.append("SAP_USER")
    if not target.password:
        missing.append(f"SAP_PASSWORD_{target.key}")
    if missing:
        raise RuntimeError("Variaveis de ambiente em falta: " + ", ".join(missing))


def _get_scripting_engine(target: SapTarget, win32_client):
    try:
        sap = win32_client.GetObject("SAPGUI")
    except Exception:
        saplogon = Path(target.saplogon_path)
        if not saplogon.exists():
            raise RuntimeError(f"SAP Logon nao encontrado em: {target.saplogon_path}")
        subprocess.Popen([str(saplogon)], shell=False)
        time.sleep(5)
        sap = win32_client.GetObject("SAPGUI")

    _try_minimize_saplogon_windows()
    application = sap.GetScriptingEngine
    if not application:
        raise RuntimeError("SAP GUI Scripting indisponivel.")
    return application


def session_info(session) -> dict:
    return {
        "system_name": str(getattr(session.Info, "SystemName", "")).strip(),
        "client": str(getattr(session.Info, "Client", "")).strip(),
        "user": str(getattr(session.Info, "User", "")).strip(),
    }


def ensure_sap_access(target: SapTarget, timeout_s: int = 40):
    _validate_target(target)

    pythoncom, win32_client = _import_sap_com()
    pythoncom.CoInitialize()

    application = _get_scripting_engine(target, win32_client)

    already = _find_logged_session(
        application,
        system_name=target.system_name,
        client=target.client,
    )
    if already:
        apply_window_mode(already)
        return already

    login_session = _find_login_session(application)
    if not login_session:
        connection = application.OpenConnection(target.connection_name, True)
        _try_minimize_saplogon_windows()
        login_session = _wait_for_connection_session(connection, timeout_s=30)
        if not login_session:
            raise RuntimeError("Nao foi possivel abrir sessao de login SAP.")

    _submit_login(login_session, target)
    ok, error = _wait_login_ok(login_session, target=target, timeout_s=timeout_s)
    if not ok:
        raise RuntimeError(f"Login SAP nao confirmado: {error}")

    apply_window_mode(login_session)
    return login_session


def ensure_sap_access_from_env(
    *,
    key: str | None = None,
    timeout_s: int = 40,
    load_env: bool = True,
):
    if load_env:
        load_dotenv_manual()
    target = resolve_sap_target_from_env(key=key)
    return ensure_sap_access(target=target, timeout_s=timeout_s)


if __name__ == "__main__":
    load_dotenv_manual()
    if not _to_bool(os.getenv("SAP_LOGIN_RUN", "true")):
        raise SystemExit(0)

    session = ensure_sap_access_from_env()
    info = session_info(session)
    print(
        f"OK: Login efetuado | Sistema={info['system_name']} | "
        f"Cliente={info['client']} | User={info['user']}"
    )
