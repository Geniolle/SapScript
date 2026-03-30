# -*- coding: utf-8 -*-
###################################################################################
# SAP Cockpit - Gestão de Requests (SE10) | Opções 1/2/3
# - Atualizado para S/4HANA: Extração robusta via Status Bar e Try/Except
###################################################################################

import os
import re
import time
import subprocess

import pythoncom
import win32com.client
import pywintypes


# =========================
# CONFIG
# =========================
SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

SLEEP_UI = 0.25
SLEEP_ACTION = 0.40

# Quantos níveis/objetos varrer no ecrã para procurar a request
SCAN_MAX_DEPTH = 6
SCAN_MAX_NODES = 2500


# =========================
# UTIL
# =========================
def _sleep(t=SLEEP_UI):
    time.sleep(t)


def start_saplogon_if_needed():
    try:
        win32com.client.GetObject("SAPGUI")
        return
    except Exception:
        pass

    if os.path.exists(SAPLOGON_PATH):
        subprocess.Popen([SAPLOGON_PATH], shell=False)
        time.sleep(2.0)


def get_sap_session(connection_index=0, session_index=0):
    pythoncom.CoInitialize()

    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
    except Exception:
        raise RuntimeError("SAP GUI não está aberto ou Scripting não está ativo.")

    if application.Children.Count == 0:
        raise RuntimeError("Nenhuma conexão SAP encontrada. Abra o SAP e faça login primeiro.")

    connection = application.Children(connection_index)

    if connection.Children.Count == 0:
        raise RuntimeError("Conexão SAP sem sessões. Abra uma sessão e faça login.")

    session = connection.Children(session_index)
    return session


def safe_find(session, sap_id):
    try:
        return session.findById(sap_id)
    except Exception:
        return None


def press(session, sap_id):
    obj = safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Elemento SAP não encontrado: {sap_id}")
    obj.press()
    _sleep(SLEEP_ACTION)


def send_vkey(session, vkey):
    try:
        session.findById("wnd[0]").sendVKey(vkey)
        _sleep(SLEEP_ACTION)
    except:
        pass


def set_text(session, sap_id, text, caret_pos=None):
    obj = safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Campo SAP não encontrado: {sap_id}")
    obj.text = text
    if caret_pos is not None:
        try:
            obj.caretPosition = caret_pos
        except Exception:
            pass
    _sleep(SLEEP_UI)


def select_radio_if_exists(session, sap_id):
    obj = safe_find(session, sap_id)
    if not obj:
        return False
    try:
        obj.select()
        obj.setFocus()
        _sleep(SLEEP_UI)
        return True
    except Exception:
        return False


def ensure_se10(session):
    okcd = safe_find(session, "wnd[0]/tbar[0]/okcd")
    if not okcd:
        raise RuntimeError("Não foi possível localizar o campo de comando (okcd).")

    okcd.text = "/nSE10"
    send_vkey(session, 0)
    time.sleep(0.8)


# =========================
# EXTRAÇÃO DO NÚMERO DA REQUEST
# =========================
_REQ_PATTERNS = [
    re.compile(r"\b[A-Z0-9]{3,4}K\d{6,}\b"),
    re.compile(r"\b[A-Z0-9]{3,4}[A-Z]\d{6,}\b"),
]


def extract_request_number_from_text(text):
    if not text:
        return None
    t = str(text).strip()
    for rgx in _REQ_PATTERNS:
        m = rgx.search(t)
        if m:
            return m.group(0)
    return None


def try_get_obj_text(obj):
    try: return obj.Text
    except: pass
    try: return obj.text
    except: pass
    try: return obj.Value
    except: pass
    return None


def extract_request_from_known_ids(session):
    # IDs onde a request costuma aparecer após o OK da descrição
    candidates = [
        "wnd[0]/sbar",         # Barra de status (Mais comum no S4: 'Ordem DEVK900xxx foi criada')
        "wnd[0]/usr/lbl[20,9]",
        "wnd[0]/usr/lbl[1,1]",
    ]
    for sap_id in candidates:
        obj = safe_find(session, sap_id)
        if obj:
            txt = try_get_obj_text(obj)
            req = extract_request_number_from_text(txt)
            if req:
                return req
    return None


def extract_request_by_scanning_usr(session):
    area = safe_find(session, "wnd[0]/usr")
    if not area: return None

    stack = [(area, 0)]
    seen = 0
    while stack:
        node, depth = stack.pop()
        seen += 1
        if seen > SCAN_MAX_NODES: break

        txt = try_get_obj_text(node)
        req = extract_request_number_from_text(txt)
        if req: return req

        if depth < SCAN_MAX_DEPTH:
            try:
                for i in range(int(node.Children.Count) - 1, -1, -1):
                    stack.append((node.Children.Item(int(i)), depth + 1))
            except: pass
    return None


def get_created_request_number(session):
    # Aguarda o SAP processar a criação no banco de dados
    time.sleep(0.8)

    # 1. Tenta IDs conhecidos e Barra de Status (mais rápido)
    req = extract_request_from_known_ids(session)
    if req: return req

    # 2. Varre o ecrã completo como último recurso
    req = extract_request_by_scanning_usr(session)
    return req


# =========================
# OPTION 3 - CRIAR NOVA REQUEST
# =========================
def criar_nova_request(session):
    ensure_se10(session)

    print("\nTipo da ordem:")
    print('1 - Ordem customizing')
    print('2 - Ordem workbench')
    tipo = ask_choice("Digite a opção (1/2): ", ["1", "2"])

    desc = input("Descrição da request (máx 60): ").strip()
    if not desc:
        desc = "REQUEST CRIADA VIA SCRIPT"
    desc = desc[:60]

    # Clicar em Criar (F6)
    press(session, "wnd[0]/tbar[1]/btn[6]")

    # Popup de Tipo de Ordem
    if tipo == "2":
        select_radio_if_exists(session, "wnd[1]/usr/radKO042-REQ_CONS_K")
    press(session, "wnd[1]/tbar[0]/btn[0]") # OK Tipo

    # Popup de Descrição
    set_text(session, "wnd[1]/usr/txtKO013-AS4TEXT", desc, caret_pos=len(desc))
    press(session, "wnd[1]/tbar[0]/btn[0]") # OK Descrição (Aqui a request é gerada)

    # --- EXTRAÇÃO SEGURA ---
    req = get_created_request_number(session)

    # Voltar ao início (/n) de forma segura sem depender de labels fixos
    try:
        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        okcd.text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        _sleep(0.5)
    except:
        pass

    print("\n✅ Processo de criação finalizado.")
    print(f"Tipo: {'Customizing' if tipo == '1' else 'Workbench'}")
    print(f"Descrição: {desc}")

    if not req:
        print("⚠️  Não consegui extrair o número automaticamente.")
        req = input("👉 Por favor, digite o número da Request criada: ").strip().upper()

    print(f"🚀 Request final: {req}")
    return req


# =========================
# OPTION 1/2/3 - MENU
# =========================
def ask_choice(prompt, allowed):
    allowed_set = set(allowed)
    while True:
        v = input(prompt).strip()
        if v in allowed_set:
            return v


def tratar_request(session):
    print("\n❓ Como deseja tratar a Request?")
    print("1: Numero da Request:")
    print("2: Listar Requests (tarefas) e escolher pela linha")
    print("3: Criar nova request:")
    opt = ask_choice("Digite a opção (1/2/3): ", ["1", "2", "3"])

    if opt == "1":
        req = input("Digite o número da Request (ex.: DEVK900123): ").strip().upper()
        return req

    if opt == "2":
        print("⚠️  Opção 2: (Implemente aqui a listagem de tarefas via SE16H se necessário)")
        req = input("Digite o número da Request: ").strip().upper()
        return req

    return criar_nova_request(session)


# =========================
# MAIN
# =========================
def main():
    start_saplogon_if_needed()
    try:
        session = get_sap_session(connection_index=0, session_index=0)
    except Exception as e:
        print(f"❌ Erro ao conectar ao SAP: {e}")
        return

    req = tratar_request(session)
    print(f"\n➡️ Fluxo segue com a request: {req}")


if __name__ == "__main__":
    main()