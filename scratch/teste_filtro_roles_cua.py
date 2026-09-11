from __future__ import annotations

import importlib.util
import os
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass
os.environ["WORKFLOW_SAP_KEY"] = "SPACLNT001"
os.environ["WORKFLOW_SAP_SYSTEM"] = "SPA"
os.environ["WORKFLOW_SAP_CLIENT"] = "001"

from sap_session import apply_window_mode, ensure_sap_access_from_env, session_info  # noqa: E402


def carregar_fluxo_remocao():
    caminho = ROOT / "Processos" / "Funções PFCG" / "J. CUA_REMOVE.py"
    spec = importlib.util.spec_from_file_location("cua_remove_gui", caminho)
    if spec is None or spec.loader is None:
        raise RuntimeError("Não foi possível carregar o fluxo CUA_REMOVE")
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    modulo.MODO_DEBUG_PASSO_A_PASSO = False
    return modulo


def main() -> None:
    username = sys.argv[1].strip().upper() if len(sys.argv) > 1 else "S170"
    subsystem = sys.argv[2].strip().upper() if len(sys.argv) > 2 else "S4DCLNT100"
    fluxo = carregar_fluxo_remocao()
    session = ensure_sap_access_from_env(key="SPACLNT001")

    try:
        session.findById("wnd[1]").sendVKey(12)
        time.sleep(0.3)
    except Exception:
        pass

    command = session.findById("wnd[0]/tbar[0]/okcd")
    command.text = "/nSU01"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(0.5)

    user_field = session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME")
    user_field.text = username
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(0.5)

    message_type, message = fluxo.obter_status_bar(session)
    if message_type in ("E", "A"):
        raise RuntimeError(message or f"Não foi possível abrir o utilizador {username}")

    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    time.sleep(0.5)
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
    time.sleep(0.5)

    grid = fluxo.obter_grid_roles(session)
    grid.currentCellColumn = "SUBSYSTEM"
    grid.contextMenu()
    grid.selectContextMenuItem("&FILTER")
    fluxo.preencher_popup_filtro(session, subsystem)
    time.sleep(0.5)

    count = fluxo.obter_row_count_grid(grid)
    encontrados = {
        fluxo.obter_valor_celula_grid(grid, row, "SUBSYSTEM")
        for row in range(count)
    }
    encontrados.discard("")
    if encontrados and encontrados != {subsystem}:
        raise RuntimeError(
            f"Filtro inesperado: esperado={subsystem}, encontrados={sorted(encontrados)}"
        )

    apply_window_mode(session, mode="show")
    info = session_info(session)
    print(
        f"OK: filtro aplicado | Sistema={info['system_name']} | Cliente={info['client']} | "
        f"Utilizador={username} | SUBSYSTEM={subsystem} | Linhas={count} | SEM_GRAVAR"
    )


if __name__ == "__main__":
    main()
