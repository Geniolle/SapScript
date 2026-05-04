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
MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}
CLIENTES_POR_AMBIENTE = {"DEV": "100", "QAD": "100", "PRD": "100", "CUA": "001"}

SLEEP_UI = 0.25
SLEEP_ACTION = 0.40

# Quantos níveis/objetos varrer no ecrã para procurar a request
SCAN_MAX_DEPTH = 6
SCAN_MAX_NODES = 2500


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on", "sim", "s"}


def _called_from_main_env() -> bool:
    return _to_bool(os.getenv("SAP_CALLED_BY_MAIN", "false"))


def _apply_sap_window_mode(session) -> None:
    mode = str(os.getenv("SAP_WINDOW_MODE", "") or "").strip().lower()
    if not mode and _to_bool(os.getenv("SAP_WINDOW_MINIMIZE", "false")):
        mode = "minimize"
    if not mode:
        mode = "show"

    try:
        wnd0 = session.findById("wnd[0]")
    except Exception:
        return

    try:
        if mode in {"minimize", "minimizar", "hidden", "hide", "ocultar", "quiet"}:
            wnd0.iconify()
        elif mode in {"show", "mostrar", "visible", "visivel", "exibir"}:
            wnd0.maximize()
    except Exception:
        return


def _parse_env_line(line: str):
    line = (line or "").strip()
    if not line or line.startswith("#") or "=" not in line:
        return None, None
    key, value = line.split("=", 1)
    key = key.strip()
    value = value.strip()
    if not key:
        return None, None
    if len(value) >= 2 and (
        (value.startswith('"') and value.endswith('"')) or
        (value.startswith("'") and value.endswith("'"))
    ):
        value = value[1:-1]
    return key, value


def _load_dotenv_manual():
    candidates = [
        os.path.join(os.getcwd(), ".env"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"),
    ]
    seen = set()
    for path in candidates:
        path_abs = os.path.abspath(path)
        if path_abs in seen:
            continue
        seen.add(path_abs)
        if not os.path.exists(path_abs):
            continue
        with open(path_abs, "r", encoding="utf-8-sig") as file_obj:
            for raw in file_obj:
                key, value = _parse_env_line(raw)
                if key and key not in os.environ:
                    os.environ[key] = value
        return path_abs
    return None


def _resolve_target_from_ambiente(ambiente_cockpit="", system_name="", client=""):
    ambiente = str(ambiente_cockpit or "").strip().upper()
    resolved_system = str(system_name or "").strip().upper()
    resolved_client = str(client or "").strip()

    if not resolved_system and ambiente in MAPA_SISTEMA:
        resolved_system = MAPA_SISTEMA[ambiente]
    if not resolved_client and ambiente in CLIENTES_POR_AMBIENTE:
        resolved_client = CLIENTES_POR_AMBIENTE[ambiente]

    return resolved_system, resolved_client


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

    saplogon_path = os.getenv("SAPLOGON_PATH", SAPLOGON_PATH).strip() or SAPLOGON_PATH
    if os.path.exists(saplogon_path):
        subprocess.Popen([saplogon_path], shell=False)
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
    _apply_sap_window_mode(session)
    return session


def get_sap_session_by_system_client(system_name="", client=""):
    pythoncom.CoInitialize()

    expected_system = str(system_name or "").strip().upper()
    expected_client = str(client or "").strip()

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except Exception:
        raise RuntimeError("SAP GUI nao esta aberto ou Scripting nao esta ativo.")

    if application.Children.Count == 0:
        raise RuntimeError("Nenhuma conexao SAP encontrada. Abra o SAP e faca login primeiro.")

    for connection in application.Children:
        for session in connection.Children:
            sess_system = str(getattr(session.Info, "SystemName", "")).strip().upper()
            sess_client = str(getattr(session.Info, "Client", "")).strip()

            if expected_system and sess_system != expected_system:
                continue
            if expected_client and sess_client != expected_client:
                continue
            _apply_sap_window_mode(session)
            return session

    raise RuntimeError(
        f"Nenhuma sessao encontrada para sistema='{expected_system or '*'}' "
        f"e mandante='{expected_client or '*'}'."
    )


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
# AUTO CREATE (SEM INPUT)
# =========================
def criar_nova_request_auto(session, tipo="1", desc=""):
    tipo = str(tipo).strip()
    if tipo not in {"1", "2"}:
        tipo = "1"

    descricao = (desc or "").strip()[:60]
    if not descricao:
        descricao = "REQUEST CRIADA VIA SCRIPT"

    ensure_se10(session)

    press(session, "wnd[0]/tbar[1]/btn[6]")
    if tipo == "2":
        select_radio_if_exists(session, "wnd[1]/usr/radKO042-REQ_CONS_K")
    press(session, "wnd[1]/tbar[0]/btn[0]")

    set_text(session, "wnd[1]/usr/txtKO013-AS4TEXT", descricao, caret_pos=len(descricao))
    press(session, "wnd[1]/tbar[0]/btn[0]")

    req = get_created_request_number(session)
    if not req:
        raise RuntimeError("Nao consegui extrair o numero da request criada automaticamente.")

    try:
        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        okcd.text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        _sleep(0.5)
    except Exception:
        pass

    print(f"REQUEST_NUMBER={req}")
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
# EXECUTAR (COMPATÍVEL COM SAP COCKPIT)
# =========================
def executar(
    ambiente_cockpit=None,
    request_ctx=None,
    request_transporte=None,
    modo_nao_interativo=False,
    tipo_ordem="customizing",
    descricao_request="",
    connection_index=0,
    session_index=0,
    system_name="",
    client="",
    chamado_pelo_main=False
):
    _load_dotenv_manual()
    called_by_main = bool(chamado_pelo_main) or _called_from_main_env()
    if called_by_main:
        modo_nao_interativo = True

    tipo_map = {
        "1": "1",
        "2": "2",
        "customizing": "1",
        "workbench": "2",
    }
    tipo_forcado = tipo_map.get(str(tipo_ordem or "").strip().lower(), "1")

    req_recebida = ""
    if isinstance(request_ctx, dict):
        req_recebida = str(request_ctx.get("request_number", "")).strip().upper()
    if not req_recebida:
        req_recebida = str(request_transporte or "").strip().upper()

    if req_recebida:
        print(f"REQUEST_NUMBER={req_recebida}")
        return req_recebida

    resolved_system, resolved_client = _resolve_target_from_ambiente(
        ambiente_cockpit=ambiente_cockpit,
        system_name=system_name,
        client=client
    )

    start_saplogon_if_needed()
    if resolved_system or resolved_client:
        session = get_sap_session_by_system_client(
            system_name=resolved_system,
            client=resolved_client
        )
    else:
        session = get_sap_session(
            connection_index=connection_index,
            session_index=session_index
        )

    if modo_nao_interativo:
        req = criar_nova_request_auto(
            session=session,
            tipo=tipo_forcado,
            desc=descricao_request
        )
    else:
        req = tratar_request(session)

    return req


# =========================
# MAIN
# =========================
def main():
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--auto-create", action="store_true")
    parser.add_argument(
        "--order-type",
        default="customizing",
        help="customizing|workbench|1|2",
    )
    parser.add_argument("--description", default="")
    parser.add_argument("--connection-index", type=int, default=0)
    parser.add_argument("--session-index", type=int, default=0)
    parser.add_argument("--system-name", default="")
    parser.add_argument("--client", default="")
    parser.add_argument("--from-main", action="store_true")
    args = parser.parse_args()

    try:
        req = executar(
            ambiente_cockpit=None,
            request_ctx=None,
            request_transporte=None,
            modo_nao_interativo=bool(args.auto_create),
            tipo_ordem=args.order_type,
            descricao_request=args.description,
            connection_index=args.connection_index,
            session_index=args.session_index,
            system_name=args.system_name,
            client=args.client,
            chamado_pelo_main=bool(args.from_main),
        )
    except Exception as e:
        print(f"Erro ao conectar ao SAP: {e}")
        return

    print(f"\nFluxo segue com a request: {req}")


if __name__ == "__main__":
    main()
