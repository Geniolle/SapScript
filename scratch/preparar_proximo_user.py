import sys, time, threading, win32service, pythoncom, win32com.client, win32clipboard
sys.stdout.reconfigure(encoding="utf-8")

USER_ALVO = "S75"
ROLES_A_REMOVER = [
    "ZFIN_DADOS_MESTRE_BASIC",
    "ZFIN_MM_COORDINATOR",
    "ZFIN_PS_BASIC",
    "ZFIN_PS_SPEC",
    "ZFI_REAL_ESTATE_SIMPLES",
    "ZMM_SUPPLY_CHAIN_BASIC",
    "ZORG_TODAS_EMPRESAS",
]

def worker():
    try:
        hwinsta = win32service.OpenWindowStation("WinSta0", True, 0x037F)
        hwinsta.SetProcessWindowStation()
        hdesk = win32service.OpenDesktop("default", 0, True, 0x01FF)
        hdesk.SetThreadDesktop()
        
        pythoncom.CoInitialize()
        rot = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
        sap = rot.GetROTEntry("SAPGUI")
        session = sap.GetScriptingEngine.Children(0).Children(0)
        
        # 1. Abrir SU01 em modo de alteração
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)
        
        session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = USER_ALVO
        session.findById("wnd[0]/tbar[1]/btn[18]").press() # Alterar
        time.sleep(0.4)
        
        # 2. Selecionar aba Funções (tabpACTG)
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
        time.sleep(0.3)
        
        grid = session.findById(
            "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/"
            "ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/"
            "cntlG_ROLES_CONTAINER/shellcont/shell"
        )
        
        # 3. Copiar lista para clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText("\r\n".join(ROLES_A_REMOVER))
        win32clipboard.CloseClipboard()
        
        # 4. Selecionar colunas SUBSYSTEM e AGR_NAME e abrir filtro
        grid.selectColumn("SUBSYSTEM")
        grid.selectColumn("AGR_NAME")
        grid.pressToolbarButton("&MB_FILTER")
        time.sleep(0.3)
        
        wnd1 = session.findById("wnd[1]")
        wnd1.findById("usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "S4PCLNT100"
        
        # Abrir seleção múltipla de Role
        wnd1.findById("usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
        time.sleep(0.3)
        
        wnd2 = session.findById("wnd[2]")
        # Upload from Clipboard
        wnd2.findById("tbar[0]/btn[24]").press()
        time.sleep(0.3)
        
        # Confirmar wnd[2] (Copy / F8)
        wnd2.findById("tbar[0]/btn[8]").press()
        time.sleep(0.3)
        
        # Confirmar wnd[1] (Enter / btn[0])
        wnd1.findById("tbar[0]/btn[0]").press()
        time.sleep(0.4)
        
        rc = grid.RowCount
        print(f"Utilizador {USER_ALVO} filtrado com sucesso! Linhas a eliminar visíveis: {rc}")
        for r in range(rc):
            sys_val = grid.GetCellValue(r, "SUBSYSTEM")
            agr_val = grid.GetCellValue(r, "AGR_NAME")
            print(f"  Linha {r}: {sys_val} | {agr_val}")

        if rc > 0:
            grid.setCurrentCell(0, "AGR_NAME")
            grid.selectedRows = "0"
            print(f"Linha 0 selecionada: '{grid.GetCellValue(0, 'AGR_NAME')}'")

    except Exception as exc:
        print("Worker error:", exc)

t = threading.Thread(target=worker)
t.start()
t.join()
