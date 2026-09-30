# -*- coding: utf-8 -*-
"""
zterm_copy_gui.py - Automação SAP GUI Scripting para criar/copiar Condições de Pagamento (OBB8)
"""
import os
import sys
import time
import re
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from Processos.criar_request import get_sap_session_by_system_client, criar_nova_request_auto, press, set_text, send_vkey, safe_find


def copy_zterm_gui(system_name="S4D", client="100", zterm_source="0001", zterm_target="Z031", zterm_name="Pagamento 30 dias", request_number=None):
    """
    Cria uma nova condição de pagamento por cópia na OBB8 via SAP GUI Scripting.
    Se request_number não for fornecido, cria automaticamente uma Customizing Request.
    """
    session = get_sap_session_by_system_client(system_name=system_name, client=client)

    # 1. Se não tiver request_number, gera automaticamente usando código + nome
    if not request_number:
        desc_req = f"{zterm_target} - {zterm_name}"[:60]
        request_number = criar_nova_request_auto(session, tipo="1", desc=desc_req)

    # 2. Entrar na OBB8
    okcd = safe_find(session, "wnd[0]/tbar[0]/okcd")
    if not okcd:
        raise RuntimeError("Campo de comando okcd não encontrado no SAP GUI.")
    okcd.text = "/nOBB8"
    send_vkey(session, 0)
    time.sleep(1.0)

    # 3. Posicionar na zterm_source
    press(session, "wnd[0]/tbar[1]/btn[23]")  # Posicionar / Position...
    set_text(session, "wnd[1]/usr/txtV_T052-ZTERM", zterm_source)
    send_vkey(session, 0)  # Enter
    time.sleep(0.5)

    # 4. Marcar a linha e Clicar Copiar Como (btn[14])
    try:
        tbl = safe_find(session, "wnd[0]/usr/tblSAPL0F30TCTRL_V_T052")
        if tbl:
            tbl.getAbsoluteRow(0).selected = True
    except Exception:
        pass

    press(session, "wnd[0]/tbar[1]/btn[14]")  # Copiar como...
    time.sleep(0.5)

    # 5. Preencher a nova ZTERM e Descrição na tela de detalhes
    set_text(session, "wnd[0]/usr/txtV_T052-ZTERM", zterm_target)
    set_text(session, "wnd[0]/usr/txtV_T052-TEXT1", zterm_name)
    send_vkey(session, 0)  # Enter para validar a cópia
    time.sleep(0.5)

    # 6. Salvar (Ctrl+S / btn[11])
    press(session, "wnd[0]/tbar[0]/btn[11]")
    time.sleep(0.8)

    # 7. Tratar Popup da Order de Transporte (se aparecer wnd[1])
    wnd1 = safe_find(session, "wnd[1]")
    if wnd1:
        set_text(session, "wnd[1]/usr/ctxtKO008-TRKORR", request_number)
        press(session, "wnd[1]/tbar[0]/btn[0]")  # Enter / OK
        time.sleep(0.8)

    # 8. Verificar mensagem na barra de status
    sbar = safe_find(session, "wnd[0]/sbar")
    sbar_txt = sbar.text if sbar else ""
    sbar_type = sbar.messageType if sbar else ""

    try:
        okcd = safe_find(session, "wnd[0]/tbar[0]/okcd")
        if okcd:
            okcd.text = "/n"
            send_vkey(session, 0)
    except Exception:
        pass

    return {
        "success": True,
        "request_number": request_number,
        "zterm_source": zterm_source,
        "zterm_target": zterm_target,
        "zterm_name": zterm_name,
        "message": sbar_txt or "Condição de pagamento criada com sucesso por cópia.",
        "message_type": sbar_type,
    }
