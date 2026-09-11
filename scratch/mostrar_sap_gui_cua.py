from __future__ import annotations

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ["WORKFLOW_SAP_KEY"] = "SPACLNT001"
os.environ["WORKFLOW_SAP_SYSTEM"] = "SPA"
os.environ["WORKFLOW_SAP_CLIENT"] = "001"

import win32con  # type: ignore  # noqa: E402
import win32api  # type: ignore  # noqa: E402
import win32gui  # type: ignore  # noqa: E402
from sap_session import ensure_sap_access_from_env, session_info  # noqa: E402


session = ensure_sap_access_from_env(key="SPACLNT001")
window = session.findById("wnd[0]")
handle = int(window.Handle)
win32gui.ShowWindow(handle, win32con.SW_RESTORE)
monitor = win32api.MonitorFromWindow(handle, win32con.MONITOR_DEFAULTTONEAREST)
work_left, work_top, work_right, work_bottom = win32api.GetMonitorInfo(monitor)["Work"]
work_width = work_right - work_left
work_height = work_bottom - work_top
win32gui.MoveWindow(handle, work_left, work_top, work_width // 2, work_height, True)
win32gui.BringWindowToTop(handle)
win32gui.SetForegroundWindow(handle)
info = session_info(session)
print(
    f"OK: janela visivel | Sistema={info['system_name']} | "
    f"Cliente={info['client']} | User={info['user']}"
)
