###################################################################################
# SAP Cockpit - Inicialização e Gestão de Sessão SAP
###################################################################################

import os
import sys
import win32com.client
import subprocess
import importlib.util
import time
import msvcrt
import inspect
import re

import pythoncom
import pywintypes

# Caminhos
base_dir = r"C:\SAP Script"
processos_dir = os.path.join(base_dir, "Processos")
saplogon_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Texto exigido (sem ícone). O ícone é aplicado no logger.
MSG_RZ11_SCRIPTING = 'Ativar na transação RZ11 o nome do parametro "sapgui/user_scripting" alterar para "TRUE"'

# Módulo de lista/pesquisa de requests
pesquisar_request_path = os.path.join(processos_dir, "pesquisar_request.py")

# Ambientes SAP
ambientes = {
    "1": ("DEV", "DESENVOLVIMENTO (S4H)"),
    "2": ("QAD", "QUALIDADE (S4H)"),
    "3": ("PRD", "PRODUÇÃO (S4H)"),
    "4": ("CUA", "CUA (PRD)")
}

mapa_sistema = {
    "DEV": "S4D",
    "QAD": "S4Q",
    "PRD": "S4P",
    "CUA": "SPA"
}

clientes_por_ambiente = {
    "DEV": "100",
    "QAD": "100",
    "PRD": "100",
    "CUA": "001"
}

# Sleeps (ajusta se o SAP “se perder” nos inputs)
SLEEP_UI = 0.25
SLEEP_ACTION = 0.40

# Scan para extrair TRKORR do ecrã (sem status bar)
SCAN_MAX_DEPTH = 6
SCAN_MAX_NODES = 2500

_REQ_PATTERNS = [
    re.compile(r"\b[A-Z0-9]{3,4}K\d{6,}\b"),      # ex.: S4QK900416 / S4DK951499
    re.compile(r"\b[A-Z0-9]{3,4}[A-Z]\d{6,}\b"),  # fallback
]


###################################################################################
# Funções auxiliares (I/O e menus)
###################################################################################

def limpar_buffer_teclado():
    try:
        while msvcrt.kbhit():
            msvcrt.getch()
    except Exception:
        pass


def ler_texto(prompt: str, oculto: bool = False) -> str:
    """
    Leitura tecla-a-tecla (aceita colar).
    - oculto=True imprime "*" para cada caractere.
    """
    limpar_buffer_teclado()
    print(prompt, end="", flush=True)

    buf = []
    while True:
        ch = msvcrt.getwch()

        # ENTER
        if ch in ("\r", "\n"):
            print()
            return "".join(buf)

        # BACKSPACE
        if ch in ("\b", "\x7f"):
            if buf:
                buf.pop()
                print("\b \b", end="", flush=True)
            continue

        # CTRL+C
        if ch == "\x03":
            raise KeyboardInterrupt

        # ignora outros control chars
        if ord(ch) < 32:
            continue

        buf.append(ch)
        print("*" if oculto else ch, end="", flush=True)


def selecionar_ambiente():
    if len(sys.argv) > 1:
        ambiente_recebido = sys.argv[1].upper()
        if ambiente_recebido in mapa_sistema:
            return ambiente_recebido
        print(f"❌ Ambiente '{ambiente_recebido}' inválido via argumento.")
        sys.exit(1)

    print("\n🌐 Ambientes disponíveis:")
    for k, (sigla, nome) in ambientes.items():
        print(f"{k}: {sigla} → {nome}")

    opcao = input("\nDigite o número do ambiente que deseja atualizar: ").strip()
    if opcao not in ambientes:
        print("❌ Opção inválida.")
        sys.exit(1)

    return ambientes[opcao][0]


def selecionar_pasta_processo():
    print("\n📂 Processos disponíveis:")
    pastas = [
        p for p in os.listdir(processos_dir)
        if os.path.isdir(os.path.join(processos_dir, p)) and p != "__pycache__"
    ]

    for i, pasta in enumerate(pastas, 1):
        print(f"{i}: {pasta}")

    try:
        escolha = int(input("\nDigite o número do processo que deseja abrir: ")) - 1
        pasta_escolhida = pastas[escolha]
        return os.path.join(processos_dir, pasta_escolhida)
    except (ValueError, IndexError):
        print("❌ Seleção inválida.")
        sys.exit(1)


def validar_request(valor: str) -> str:
    """
    Exemplo: S4QK900396 / S4DK951499 (3-4 chars + K + 6+ dígitos)
    """
    v = (valor or "").strip().upper().replace(" ", "")
    if not v:
        return ""
    if re.match(r"^[A-Z0-9]{3,4}K\d{6,}$", v):
        return v
    return ""


def carregar_pesquisar_request():
    """
    Carrega dinamicamente o ficheiro pesquisar_request.py
    """
    if not os.path.exists(pesquisar_request_path):
        print(f"❌ Ficheiro não encontrado: {pesquisar_request_path}")
        return None

    spec = importlib.util.spec_from_file_location("pesquisar_request", pesquisar_request_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def escolher_request_por_linha(lista_resultados):
    """
    lista_resultados: list[(TRKORR, AS4TEXT)]
    Retorna (trkorr, as4text) escolhido.
    """
    if not lista_resultados:
        return ("", "")

    while True:
        raw = input(f"\nDigite o número da linha da request (1-{len(lista_resultados)}): ").strip()
        if not raw.isdigit():
            print("❌ Digite apenas números.")
            continue

        idx = int(raw)
        if idx < 1 or idx > len(lista_resultados):
            print("❌ Número fora do intervalo.")
            continue

        trkorr, as4text = lista_resultados[idx - 1]
        return (trkorr, as4text)


def _resetar_env_request():
    os.environ["SAP_REQUEST_OPTION"] = ""
    os.environ["SAP_REQUEST_NUMBER"] = ""
    os.environ["SAP_REQUEST_DESC"] = ""
    os.environ["SAP_SEARCH_TEXT"] = ""


###################################################################################
# SAP GUI helpers (SE10 - criar request sem status bar)
###################################################################################

def _sleep(t=SLEEP_UI):
    time.sleep(t)


def _safe_find(session, sap_id: str):
    try:
        return session.findById(sap_id)
    except pywintypes.com_error:
        return None
    except Exception:
        return None


def _press(session, sap_id: str):
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Elemento SAP não encontrado: {sap_id}")
    obj.press()
    _sleep(SLEEP_ACTION)


def _send_vkey(session, vkey: int):
    session.findById("wnd[0]").sendVKey(vkey)
    _sleep(SLEEP_ACTION)


def _set_text(session, sap_id: str, text: str, caret_pos: int | None = None):
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Campo SAP não encontrado: {sap_id}")
    obj.text = text
    if caret_pos is not None:
        try:
            obj.caretPosition = caret_pos
        except Exception:
            pass
    _sleep(SLEEP_UI)


def _ensure_se10(session):
    okcd = _safe_find(session, "wnd[0]/tbar[0]/okcd")
    if not okcd:
        raise RuntimeError("Não foi possível localizar o campo de comando (okcd).")

    okcd.text = "/nSE10"
    _send_vkey(session, 0)
    time.sleep(0.8)


def _extract_request_number_from_text(text):
    if not text:
        return None
    t = str(text).strip()
    for rgx in _REQ_PATTERNS:
        m = rgx.search(t)
        if m:
            return m.group(0)
    return None


def _try_get_obj_text(obj):
    try:
        return obj.Text
    except Exception:
        pass
    try:
        return obj.text
    except Exception:
        pass
    try:
        return obj.Value
    except Exception:
        pass
    return None


def _extract_request_from_known_ids(session):
    for sap_id in ["wnd[0]/usr/lbl[20,9]", "wnd[0]/usr/lbl[1,1]"]:
        obj = _safe_find(session, sap_id)
        if not obj:
            continue
        req = _extract_request_number_from_text(_try_get_obj_text(obj))
        if req:
            return req
    return None


def _extract_request_by_scanning_usr(session):
    area = _safe_find(session, "wnd[0]/usr")
    if not area:
        return None

    stack = [(area, 0)]
    seen = 0

    while stack:
        node, depth = stack.pop()
        seen += 1
        if seen > SCAN_MAX_NODES:
            break

        req = _extract_request_number_from_text(_try_get_obj_text(node))
        if req:
            return req

        if depth >= SCAN_MAX_DEPTH:
            continue

        try:
            cnt = int(node.Children.Count)
        except Exception:
            cnt = 0

        for i in range(cnt - 1, -1, -1):
            try:
                child = node.Children.Item(i)
                stack.append((child, depth + 1))
            except Exception:
                continue

    return None


def _get_created_request_number(session):
    time.sleep(0.6)
    return _extract_request_from_known_ids(session) or _extract_request_by_scanning_usr(session)


def _select_radio_if_exists(session, sap_id: str) -> bool:
    obj = _safe_find(session, sap_id)
    if not obj:
        return False
    try:
        obj.select()
        obj.setFocus()
        _sleep(SLEEP_UI)
        return True
    except Exception:
        return False


def _criar_nova_request_no_sap(session) -> tuple[str, str, str]:
    _ensure_se10(session)

    print("\nTipo da ordem:")
    print('1 - Ordem customizing')
    print('2 - Ordem workbench')

    while True:
        tipo = input("Digite a opção (1/2): ").strip()
        if tipo in ("1", "2"):
            break
        print("❌ Opção inválida. Use apenas 1 ou 2.")

    desc = input("Descrição da request (máx 60): ").strip()
    if not desc:
        desc = "REQUEST CRIADA VIA SCRIPT"
    desc = desc[:60]

    _press(session, "wnd[0]/tbar[1]/btn[6]")  # criar

    if tipo == "2":
        _select_radio_if_exists(session, "wnd[1]/usr/radKO042-REQ_CONS_K")

    _press(session, "wnd[1]/tbar[0]/btn[0]")  # OK tipo

    _set_text(session, "wnd[1]/usr/txtKO013-AS4TEXT", desc, caret_pos=len(desc))
    _press(session, "wnd[1]/tbar[0]/btn[0]")  # OK desc

    trkorr = _get_created_request_number(session)

    # voltar (/n)
    okcd = _safe_find(session, "wnd[0]/tbar[0]/okcd")
    if okcd:
        okcd.text = "/n"
        _send_vkey(session, 0)

    tipo_txt = "Customizing" if tipo == "1" else "Workbench"

    print("\n✅ Request criada.")
    print(f"Tipo: {tipo_txt}")
    print(f"Descrição: {desc}")

    if not trkorr:
        trkorr = input("Não consegui extrair a request automaticamente. Cole aqui (ex.: S4QK900416): ").strip().upper()

    print(f"Request: {trkorr}")
    return (trkorr, desc, tipo_txt)


###################################################################################
# Request menu (apenas quando o processo realmente precisa)
###################################################################################

def perguntar_opcao_request(sistema_desejado: str, session) -> dict:
    print("\n============================================================")
    print("🚚 Opções de configuração de Transporte.\n")
    print("   1 - Escreva o número da Request")
    print("   2 - Criar nova ordem de transporte")
    print("   3 - Pesquisar suas request criadas.")
    print("   4 - Prima [Enter] vazio para NÃO transportar")
    print("============================================================")

    while True:
        opc = input("\n👉 Opção: ").strip()
        
        if opc in ("1", "2", "3", "4", ""):
            if opc == "":
                opc = "4"
            break
        print("❌ Opção inválida. Use 1, 2, 3, 4 ou apenas pressione Enter.")

    ctx = {
        "request_option": opc,
        "request_number": "",
        "request_desc": "",
        "search_text": ""
    }

    if opc == "1":
        while True:
            num_raw = input("🔢 Numero da Request (ex: S4QK900396): ").strip()
            num = validar_request(num_raw)
            if num:
                ctx["request_number"] = num
                break
            print("❌ Request inválida. Exemplo válido: S4QK900396")

    elif opc == "2":
        trkorr, desc, _tipo_txt = _criar_nova_request_no_sap(session)
        ctx["request_number"] = validar_request(trkorr) or trkorr.strip().upper()
        ctx["request_desc"] = desc

    elif opc == "3":
        mod_pesq = carregar_pesquisar_request()
        if not mod_pesq or not hasattr(mod_pesq, "listar_requests"):
            print("❌ Módulo pesquisar_request.py não carregado ou função listar_requests não encontrada.")
            return ctx

        try:
            lista = mod_pesq.listar_requests(
                system_name=sistema_desejado,
                max_rows="5000",
                include_requests=False,
                use_new_mode=True,
                minimize=True,
                close_after=True
            )
        except TypeError:
            lista = mod_pesq.listar_requests(system_name=sistema_desejado, max_rows="5000")
        except Exception as e:
            print(f"❌ Falha ao gerar lista (pesquisar_request.py): {e}")
            return ctx

        if not lista:
            print("❌ Nenhuma request encontrada na lista.")
            return ctx

        trkorr, as4text = escolher_request_por_linha(lista)
        if trkorr:
            ctx["request_number"] = trkorr
            ctx["request_desc"] = as4text
            print(f"\n✅ Request selecionada: {trkorr} | {as4text}")

    elif opc == "4":
        print("⏭️  Nenhuma request selecionada (Transporte ignorado).")
        ctx["request_number"] = ""

    os.environ["SAP_REQUEST_OPTION"] = ctx["request_option"]
    os.environ["SAP_REQUEST_NUMBER"] = ctx["request_number"]
    os.environ["SAP_REQUEST_DESC"] = ctx["request_desc"]
    os.environ["SAP_SEARCH_TEXT"] = ctx["search_text"]

    return ctx


###################################################################################
# Execução de processos (compatível + “file first” + dinâmico)
###################################################################################

def _analisar_exec_signature(func):
    sig = inspect.signature(func)
    params = list(sig.parameters.values())

    def find_param(*names_lower):
        for p in params:
            if p.name.lower() in names_lower:
                return p
        return None

    p_request_ctx = find_param("request_ctx", "ctx", "context")
    p_request_transp = find_param("request_transporte", "request_number", "trkorr")
    p_file = find_param("caminho_ficheiro", "xlsx", "excel_path", "file_path", "caminho_excel", "caminho_arquivo")
    p_pfcg = find_param("pfcg_object", "objeto_pfcg", "sheet_name")

    return {
        "sig": sig,
        "params": params,
        "has_kwargs": any(p.kind == inspect.Parameter.VAR_KEYWORD for p in params),
        "p_request_ctx": p_request_ctx,
        "p_request_transp": p_request_transp,
        "p_file": p_file,
        "p_pfcg": p_pfcg,
        "file_is_optional": bool(p_file and p_file.default is not inspect._empty),
        "request_transp_is_optional": bool(p_request_transp and p_request_transp.default is not inspect._empty),
        "request_ctx_is_optional": bool(p_request_ctx and p_request_ctx.default is not inspect._empty),
        "pfcg_is_optional": bool(p_pfcg and p_pfcg.default is not inspect._empty),
    }


def executar_processo(ambiente_cockpit, caminho_pasta, sistema_desejado, session):
    while True:
        scripts_py = [f for f in os.listdir(caminho_pasta) if f.endswith(".py") and not f.startswith("~$")]
        if not scripts_py:
            print("❌ Nenhum script .py encontrado na pasta selecionada.")
            return

        print("\n📂 Sub-Processos disponíveis:")
        for i, script in enumerate(scripts_py, 1):
            print(f"{i}: {script}")
        print(f"{len(scripts_py)+1}: 🔙 Voltar ao menu de Processos")

        try:
            escolha = int(input("\nDigite o número do sub-processo que deseja executar: ")) - 1
            if escolha == len(scripts_py):
                return "voltar"
            processo_escolhido = scripts_py[escolha]
        except (ValueError, IndexError):
            print("❌ Seleção inválida. Tente novamente.\n")
            continue

        _resetar_env_request()

        print(f"\n✅ Processo selecionado: {processo_escolhido}")

        caminho_script = os.path.join(caminho_pasta, processo_escolhido)
        spec = importlib.util.spec_from_file_location("modulo_processo", caminho_script)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)

        if not hasattr(modulo, "executar"):
            print(f"⚠️ O ficheiro '{processo_escolhido}' não contém a função 'executar()'.")
            continue

        exec_fn = modulo.executar
        info = _analisar_exec_signature(exec_fn)

        precisa_request_agora = False
        if info["p_request_ctx"] and not info["request_ctx_is_optional"]:
            precisa_request_agora = True
        elif info["p_request_transp"] and not info["request_transp_is_optional"] and not info["file_is_optional"]:
            precisa_request_agora = True

        request_ctx = {"request_option": "", "request_number": "", "request_desc": "", "search_text": ""}
        if precisa_request_agora:
            request_ctx = perguntar_opcao_request(sistema_desejado, session)

        kwargs = {}

        # ---> Nome da Aba (Sheet) dinâmico baseado no nome do script <---
        if info["p_pfcg"]:
            # 1. Remove a extensão .py
            nome_sem_ext = processo_escolhido.replace('.py', '').strip()
            
            # 2. Divide a string no primeiro ponto e fica com a parte da direita
            if '.' in nome_sem_ext:
                aba_calculada = nome_sem_ext.split('.', 1)[1].strip()
            else:
                aba_calculada = nome_sem_ext
            
            print(f"👉 Aba do Excel detetada automaticamente: '{aba_calculada}'")
            kwargs[info["p_pfcg"].name] = aba_calculada

        # Se for preciso abrir ficheiro via popup
        if info["p_file"] and not info["file_is_optional"]:
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                path = filedialog.askopenfilename(
                    title="Selecione o ficheiro Excel",
                    filetypes=(("Ficheiros Excel", "*.xlsx"), ("Todos os ficheiros", "*.*"))
                )
                root.destroy()
                if not path:
                    print("❌ Operação cancelada (ficheiro não selecionado).")
                    continue
                kwargs[info["p_file"].name] = path
            except Exception as e:
                print(f"❌ Falha ao abrir popup de ficheiro: {e}")
                continue

        if info["p_request_ctx"]:
            kwargs[info["p_request_ctx"].name] = request_ctx

        if info["p_request_transp"] and request_ctx.get("request_number"):
            if precisa_request_agora:
                kwargs[info["p_request_transp"].name] = request_ctx["request_number"]

        try:
            exec_fn(ambiente_cockpit, **kwargs)
        except TypeError:
            exec_fn(ambiente_cockpit)
        except Exception as e:
            print(f"❌ Erro ao executar '{processo_escolhido}': {e}")
            continue


###################################################################################
# Scripting / sessão SAP
###################################################################################

def _is_scripting_disabled_error(exc: Exception) -> bool:
    parts = [str(exc)]
    try:
        if hasattr(exc, "excepinfo") and exc.excepinfo:
            parts.append(" ".join([str(p) for p in exc.excepinfo if p]))
    except Exception:
        pass

    hay = " ".join(parts).lower()
    has_scripting = any(k in hay for k in ["scripting", "sapgui", "script"])
    has_disabled = any(k in hay for k in [
        "disabled", "not enabled", "not active", "inactive",
        "desativ", "inativ", "não está ativo", "nao esta ativo"
    ])
    return has_scripting and has_disabled


def _log_alerta_rz11():
    print(f"⚠️  {MSG_RZ11_SCRIPTING}")


def _erro_scripting_inativo(exc: Exception | None = None):
    print("❌ O scripting do SAP GUI não está ativo ou não foi possível inicializar o objeto SAPGUI.")
    _log_alerta_rz11()
    if exc:
        print(f"🔧 Detalhes técnicos: {exc}")
    sys.exit(1)


def _is_sap_logado(session, cliente_esperado: str) -> bool:
    try:
        return bool(session.Info.User) and str(session.Info.Client) == str(cliente_esperado)
    except Exception:
        return False


def _aguardar_login(session, cliente_esperado: str, timeout_s: int = 20) -> bool:
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        if _is_sap_logado(session, cliente_esperado):
            return True
        time.sleep(0.5)
    return False


def _log_scripting_status_apenas_quando_logado(session, cliente_esperado: str):
    if _is_sap_logado(session, cliente_esperado):
        print("🔍 A verificar disponibilidade do SAP GUI Scripting...")
        print("✅ SAP GUI Scripting está ativo.")


def _tem_alguma_sessao_ativa(application) -> bool:
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            try:
                if conn.Children.Count > 0:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _encontrar_sessao_do_sistema(application, sistema_desejado: str):
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            try:
                for j in range(conn.Children.Count):
                    sess = conn.Children(j)
                    try:
                        if str(sess.Info.SystemName).upper() == str(sistema_desejado).upper():
                            return sess, conn
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception:
        pass

    return None, None


###################################################################################
# Execução Principal - Com tratamento de erro de conexão inicial
###################################################################################

ambiente_cockpit = selecionar_ambiente()
sistema_desejado = mapa_sistema.get(ambiente_cockpit)
cliente_esperado = clientes_por_ambiente.get(ambiente_cockpit, "100")
nome_logon = dict((v[0], v[1]) for v in ambientes.values()).get(ambiente_cockpit, ambiente_cockpit)

print(f"\n🌍 Ambiente selecionado: {ambiente_cockpit} ({nome_logon})")
print("\n🔄 A verificar se já existe uma sessão aberta no ambiente desejado...")

try:
    pythoncom.CoInitialize()
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        print("🚀 SAP Logon não detectado. Iniciando executável...")
        subprocess.Popen(saplogon_path)
        time.sleep(5)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")

    application = SapGuiAuto.GetScriptingEngine
    if not application:
        raise RuntimeError("Scripting Engine desativado (RZ11).")

except Exception as e:
    _erro_scripting_inativo(e)

session, connection = _encontrar_sessao_do_sistema(application, sistema_desejado)

if session is None:
    if not _tem_alguma_sessao_ativa(application):
        usuario = input("👤 Utilizador: ").strip()
        senha = ler_texto("🔒 Senha: ", oculto=True).strip()

        try:
            print(f"📡 Abrindo conexão: {nome_logon}...")
            connection = application.OpenConnection(nome_logon, True)

            tentativas = 0
            while connection.Children.Count == 0 and tentativas < 20:
                time.sleep(0.5)
                tentativas += 1

            if connection.Children.Count == 0:
                print("\n❌ A sessão SAP não foi iniciada corretamente.")
                _log_alerta_rz11()
                sys.exit(1)

            session = connection.Children(0)

            try:
                if str(session.Info.SystemName).upper() != str(sistema_desejado).upper():
                    print(f"\n❌ A sessão SAP aberta não pertence ao ambiente '{ambiente_cockpit}'.")
                    sys.exit(1)
            except Exception:
                pass

            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = cliente_esperado
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
            session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            session.findById("wnd[0]").sendVKey(0)

        except Exception as e:
            if _is_scripting_disabled_error(e):
                _erro_scripting_inativo(e)
            else:
                print("\n❌ Erro ao tentar abrir a ligação SAP.")
                print("📡 Verifique se o servidor SAP está ligado e disponível.")
                print(f"🔧 Detalhes técnicos: {e}")
                sys.exit(1)

        if not _aguardar_login(session, cliente_esperado, timeout_s=20):
            print("\n❌ Login não foi confirmado (User/Client não disponíveis).")
            print("ℹ️  Verifique se o SAP pediu pop-up, senha extra, ou se o login falhou.")
            sys.exit(1)

    else:
        print("\n⚠️ Existe uma sessão SAP ativa, mas não é do ambiente selecionado.")
        print(f"ℹ️  Abra manualmente a ligação do ambiente '{nome_logon}' no SAP Logon e faça login (se necessário).")
        input("🕑 Pressione ENTER quando a sessão do ambiente estiver aberta...")

        session, connection = _encontrar_sessao_do_sistema(application, sistema_desejado)
        if session is None:
            print("❌ Não foi encontrada uma sessão do ambiente selecionado após confirmação. A encerrar.")
            sys.exit(1)

if not _is_sap_logado(session, cliente_esperado):
    print(f"\n⚠️ Sessão do ambiente '{ambiente_cockpit}' encontrada, mas sem login ou client incorreto (esperado: {cliente_esperado}).")
    print("🔁 Faça o login manualmente no SAP GUI (e ajuste o client se necessário).")
    input("🕑 Pressione ENTER assim que tiver terminado o login...")

    if not _aguardar_login(session, cliente_esperado, timeout_s=20):
        print("❌ O login ainda não foi detectado corretamente após confirmação. A encerrar.")
        sys.exit(1)

_log_scripting_status_apenas_quando_logado(session, cliente_esperado)

try:
    print(f"✅ Sessão SAP pronta no sistema: {session.Info.SystemName}")
    print(f"👤 Utilizador SAP: {session.Info.User} | Cliente: {session.Info.Client}")
except Exception:
    pass


###################################################################################
# Loop principal do Cockpit
###################################################################################

while True:
    caminho_processo = selecionar_pasta_processo()
    resultado = executar_processo(
        ambiente_cockpit,
        caminho_pasta=caminho_processo,
        sistema_desejado=sistema_desejado,
        session=session
    )
    if resultado != "voltar":
        continue