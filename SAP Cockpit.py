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

from app.config import (
    PROCESSOS_DIR,
    SAPLOGON_PATH,
    MSG_RZ11_SCRIPTING,
    PESQUISAR_REQUEST_PATH,
    AMBIENTES,
    MAPA_SISTEMA,
    CLIENTES_POR_AMBIENTE,
    SLEEP_UI,
    SLEEP_ACTION,
    SCAN_MAX_DEPTH,
    SCAN_MAX_NODES,
    REQ_PATTERNS,
)

from app.ui import (
    console,
    mostrar_titulo,
    mostrar_ambientes,
    mostrar_processos,
    mostrar_subprocessos,
    info,
    ok,
    warn,
    erro,
    destaque,
    linha,
)


###################################################################################
# Funções auxiliares (I/O, .env e menus)
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

        if ch in ("\r", "\n"):
            print()
            return "".join(buf)

        if ch in ("\b", "\x7f"):
            if buf:
                buf.pop()
                print("\b \b", end="", flush=True)
            continue

        if ch == "\x03":
            raise KeyboardInterrupt

        if ord(ch) < 32:
            continue

        buf.append(ch)
        print("*" if oculto else ch, end="", flush=True)


def _parse_env_line(linha: str):
    """
    Faz parse simples de uma linha KEY=VALUE.
    Ignora linhas vazias e comentários.
    Mantém o valor como texto literal, removendo apenas aspas externas simples/duplas.
    """
    if not linha:
        return None, None

    linha = linha.strip()
    if not linha or linha.startswith("#"):
        return None, None

    if "=" not in linha:
        return None, None

    chave, valor = linha.split("=", 1)
    chave = chave.strip()
    valor = valor.strip()

    if not chave:
        return None, None

    if len(valor) >= 2 and (
        (valor.startswith('"') and valor.endswith('"')) or
        (valor.startswith("'") and valor.endswith("'"))
    ):
        valor = valor[1:-1]

    return chave, valor


def _carregar_dotenv_manual():
    """
    Carrega manualmente variáveis de um ficheiro .env sem depender de python-dotenv.
    Procura nesta ordem:
      1) diretório atual
      2) diretório do ficheiro atual
      3) diretório pai do ficheiro atual
    Não sobrescreve variáveis já existentes no ambiente do processo.
    """
    base_atual = os.getcwd()
    base_script = os.path.dirname(os.path.abspath(__file__))
    base_script_pai = os.path.dirname(base_script)

    candidatos = [
        os.path.join(base_atual, ".env"),
        os.path.join(base_script, ".env"),
        os.path.join(base_script_pai, ".env"),
    ]

    vistos = set()
    for caminho in candidatos:
        caminho_norm = os.path.abspath(caminho)
        if caminho_norm in vistos:
            continue
        vistos.add(caminho_norm)

        if not os.path.exists(caminho_norm):
            continue

        with open(caminho_norm, "r", encoding="utf-8-sig") as f:
            for linha in f:
                chave, valor = _parse_env_line(linha)
                if not chave:
                    continue

                if chave not in os.environ:
                    os.environ[chave] = valor
        return caminho_norm

    return None


def _carregar_dotenv():
    caminho = _carregar_dotenv_manual()
    if caminho:
        info(f"Ficheiro .env carregado: {caminho}")
    else:
        warn("Ficheiro .env não encontrado nos caminhos esperados.")


def _obter_credenciais_env(sistema_desejado: str, cliente_esperado: str) -> tuple[str, str, str, str]:
    """
    Lê as credenciais do .env usando:
      SAP_USER
      SAP_LANGUAGE (opcional, default PT)
      SAP_PASSWORD_{SISTEMA}CLNT{CLIENTE}

    Exemplo:
      SAP_PASSWORD_S4QCLNT100
    """
    _carregar_dotenv()

    sistema = str(sistema_desejado or "").strip().upper()
    cliente = str(cliente_esperado or "").strip()
    usuario = os.getenv("SAP_USER", "").strip()
    idioma = os.getenv("SAP_LANGUAGE", "PT").strip() or "PT"

    chave_password = f"SAP_PASSWORD_{sistema}CLNT{cliente}"
    senha = os.getenv(chave_password, "").strip()

    if not usuario:
        raise RuntimeError("Variável SAP_USER não encontrada ou vazia no ficheiro .env.")

    if not senha:
        raise RuntimeError(
            f"Variável '{chave_password}' não encontrada ou vazia no ficheiro .env."
        )

    return usuario, senha, idioma, chave_password


def selecionar_ambiente():
    if len(sys.argv) > 1:
        ambiente_recebido = sys.argv[1].upper()
        if ambiente_recebido in MAPA_SISTEMA:
            return ambiente_recebido
        erro(f"Ambiente '{ambiente_recebido}' inválido via argumento.")
        sys.exit(1)

    mostrar_titulo()
    mostrar_ambientes(AMBIENTES)

    opcao = input("\nDigite o número do ambiente que deseja atualizar: ").strip()
    if opcao not in AMBIENTES:
        erro("Opção inválida.")
        sys.exit(1)

    return AMBIENTES[opcao][0]


def selecionar_pasta_processo():
    pastas = sorted(
        [
            p for p in os.listdir(PROCESSOS_DIR)
            if os.path.isdir(os.path.join(PROCESSOS_DIR, p)) and p != "__pycache__"
        ]
    )

    mostrar_processos(pastas)

    try:
        escolha = int(input("\nDigite o número do processo que deseja abrir: ")) - 1
        pasta_escolhida = pastas[escolha]
        return os.path.join(PROCESSOS_DIR, pasta_escolhida)
    except (ValueError, IndexError):
        erro("Seleção inválida.")
        sys.exit(1)


def validar_request(valor: str) -> str:
    v = (valor or "").strip().upper().replace(" ", "")
    if not v:
        return ""
    if re.match(r"^[A-Z0-9]{3,4}K\d{6,}$", v):
        return v
    return ""


def carregar_pesquisar_request():
    if not os.path.exists(PESQUISAR_REQUEST_PATH):
        erro(f"Ficheiro não encontrado: {PESQUISAR_REQUEST_PATH}")
        return None

    spec = importlib.util.spec_from_file_location("pesquisar_request", PESQUISAR_REQUEST_PATH)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def escolher_request_por_linha(lista_resultados):
    if not lista_resultados:
        return ("", "")

    while True:
        raw = input(f"\nDigite o número da linha da request (1-{len(lista_resultados)}): ").strip()
        if not raw.isdigit():
            erro("Digite apenas números.")
            continue

        idx = int(raw)
        if idx < 1 or idx > len(lista_resultados):
            erro("Número fora do intervalo.")
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
    for rgx in REQ_PATTERNS:
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

    linha()
    destaque("CRIAR NOVA REQUEST")
    print("\nTipo da ordem:")
    print("1 - Ordem customizing")
    print("2 - Ordem workbench")

    while True:
        tipo = input("Digite a opção (1/2): ").strip()
        if tipo in ("1", "2"):
            break
        erro("Opção inválida. Use apenas 1 ou 2.")

    desc = input("Descrição da request (máx 60): ").strip()
    if not desc:
        desc = "REQUEST CRIADA VIA SCRIPT"
    desc = desc[:60]

    _press(session, "wnd[0]/tbar[1]/btn[6]")

    if tipo == "2":
        _select_radio_if_exists(session, "wnd[1]/usr/radKO042-REQ_CONS_K")

    _press(session, "wnd[1]/tbar[0]/btn[0]")
    _set_text(session, "wnd[1]/usr/txtKO013-AS4TEXT", desc, caret_pos=len(desc))
    _press(session, "wnd[1]/tbar[0]/btn[0]")

    trkorr = _get_created_request_number(session)

    okcd = _safe_find(session, "wnd[0]/tbar[0]/okcd")
    if okcd:
        okcd.text = "/n"
        _send_vkey(session, 0)

    tipo_txt = "Customizing" if tipo == "1" else "Workbench"

    ok("Request criada.")
    info(f"Tipo: {tipo_txt}")
    info(f"Descrição: {desc}")

    if not trkorr:
        trkorr = input("Não consegui extrair a request automaticamente. Cole aqui (ex.: S4QK900416): ").strip().upper()

    info(f"Request: {trkorr}")
    return (trkorr, desc, tipo_txt)


###################################################################################
# Request menu
###################################################################################

def perguntar_opcao_request(sistema_desejado: str, session) -> dict:
    linha()
    destaque("OPÇÕES DE TRANSPORTE")
    print("1 - Escreva o número da Request")
    print("2 - Criar nova ordem de transporte")
    print("3 - Pesquisar suas request criadas")
    print("4 - Prima [Enter] vazio para NÃO transportar")
    linha()

    while True:
        opc = input("\nOpção: ").strip()

        if opc in ("1", "2", "3", "4", ""):
            if opc == "":
                opc = "4"
            break
        erro("Opção inválida. Use 1, 2, 3, 4 ou apenas pressione Enter.")

    ctx = {
        "request_option": opc,
        "request_number": "",
        "request_desc": "",
        "search_text": ""
    }

    if opc == "1":
        while True:
            num_raw = input("Numero da Request (ex: S4QK900396): ").strip()
            num = validar_request(num_raw)
            if num:
                ctx["request_number"] = num
                break
            erro("Request inválida. Exemplo válido: S4QK900396")

    elif opc == "2":
        trkorr, desc, _tipo_txt = _criar_nova_request_no_sap(session)
        ctx["request_number"] = validar_request(trkorr) or trkorr.strip().upper()
        ctx["request_desc"] = desc

    elif opc == "3":
        mod_pesq = carregar_pesquisar_request()
        if not mod_pesq or not hasattr(mod_pesq, "listar_requests"):
            erro("Módulo pesquisar_request.py não carregado ou função listar_requests não encontrada.")
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
            erro(f"Falha ao gerar lista (pesquisar_request.py): {e}")
            return ctx

        if not lista:
            warn("Nenhuma request encontrada na lista.")
            return ctx

        trkorr, as4text = escolher_request_por_linha(lista)
        if trkorr:
            ctx["request_number"] = trkorr
            ctx["request_desc"] = as4text
            ok(f"Request selecionada: {trkorr} | {as4text}")

    elif opc == "4":
        warn("Nenhuma request selecionada (Transporte ignorado).")
        ctx["request_number"] = ""

    os.environ["SAP_REQUEST_OPTION"] = ctx["request_option"]
    os.environ["SAP_REQUEST_NUMBER"] = ctx["request_number"]
    os.environ["SAP_REQUEST_DESC"] = ctx["request_desc"]
    os.environ["SAP_SEARCH_TEXT"] = ctx["search_text"]

    return ctx


###################################################################################
# Execução de processos
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
        scripts_py = sorted([f for f in os.listdir(caminho_pasta) if f.endswith(".py") and not f.startswith("~$")])
        if not scripts_py:
            erro("Nenhum script .py encontrado na pasta selecionada.")
            return

        mostrar_subprocessos(scripts_py)

        try:
            escolha = int(input("\nDigite o número do sub-processo que deseja executar: ")) - 1
            if escolha == len(scripts_py):
                return "voltar"
            processo_escolhido = scripts_py[escolha]
        except (ValueError, IndexError):
            erro("Seleção inválida. Tente novamente.")
            continue

        _resetar_env_request()

        ok(f"Processo selecionado: {processo_escolhido}")

        caminho_script = os.path.join(caminho_pasta, processo_escolhido)
        spec = importlib.util.spec_from_file_location("modulo_processo", caminho_script)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)

        if not hasattr(modulo, "executar"):
            warn(f"O ficheiro '{processo_escolhido}' não contém a função 'executar()'.")
            continue

        exec_fn = modulo.executar
        info_exec = _analisar_exec_signature(exec_fn)

        precisa_request_agora = False
        if info_exec["p_request_ctx"] and not info_exec["request_ctx_is_optional"]:
            precisa_request_agora = True
        elif info_exec["p_request_transp"] and not info_exec["request_transp_is_optional"] and not info_exec["file_is_optional"]:
            precisa_request_agora = True

        request_ctx = {"request_option": "", "request_number": "", "request_desc": "", "search_text": ""}
        if precisa_request_agora:
            request_ctx = perguntar_opcao_request(sistema_desejado, session)

        kwargs = {}

        if info_exec["p_pfcg"]:
            nome_sem_ext = processo_escolhido.replace(".py", "").strip()
            if "." in nome_sem_ext:
                aba_calculada = nome_sem_ext.split(".", 1)[1].strip()
            else:
                aba_calculada = nome_sem_ext

            info(f"Aba do Excel detetada automaticamente: '{aba_calculada}'")
            kwargs[info_exec["p_pfcg"].name] = aba_calculada

        if info_exec["p_file"] and not info_exec["file_is_optional"]:
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
                    erro("Operação cancelada (ficheiro não selecionado).")
                    continue
                kwargs[info_exec["p_file"].name] = path
            except Exception as e:
                erro(f"Falha ao abrir popup de ficheiro: {e}")
                continue

        if info_exec["p_request_ctx"]:
            kwargs[info_exec["p_request_ctx"].name] = request_ctx

        if info_exec["p_request_transp"] and request_ctx.get("request_number"):
            if precisa_request_agora:
                kwargs[info_exec["p_request_transp"].name] = request_ctx["request_number"]

        try:
            exec_fn(ambiente_cockpit, **kwargs)
        except TypeError:
            exec_fn(ambiente_cockpit)
        except Exception as e:
            erro(f"Erro ao executar '{processo_escolhido}': {e}")
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
    warn(MSG_RZ11_SCRIPTING)


def _erro_scripting_inativo(exc: Exception | None = None):
    erro("O scripting do SAP GUI não está ativo ou não foi possível inicializar o objeto SAPGUI.")
    _log_alerta_rz11()
    if exc:
        erro(f"Detalhes técnicos: {exc}")
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
        info("A verificar disponibilidade do SAP GUI Scripting...")
        ok("SAP GUI Scripting está ativo.")


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
# Execução Principal
###################################################################################

ambiente_cockpit = selecionar_ambiente()
sistema_desejado = MAPA_SISTEMA.get(ambiente_cockpit)
cliente_esperado = CLIENTES_POR_AMBIENTE.get(ambiente_cockpit, "100")
nome_logon = dict((v[0], v[1]) for v in AMBIENTES.values()).get(ambiente_cockpit, ambiente_cockpit)

mostrar_titulo(
    ambiente=f"{ambiente_cockpit} ({nome_logon})",
    sistema=sistema_desejado,
    cliente=cliente_esperado,
)

info("A verificar se já existe uma sessão aberta no ambiente desejado...")

try:
    pythoncom.CoInitialize()
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        warn("SAP Logon não detectado. Iniciando executável...")
        subprocess.Popen(SAPLOGON_PATH)
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
        try:
            usuario, senha, idioma, chave_password = _obter_credenciais_env(
                sistema_desejado=sistema_desejado,
                cliente_esperado=cliente_esperado
            )
            info(f"Credenciais carregadas do .env | CHAVE_PASSWORD={chave_password}")
        except Exception as e:
            erro(f"Falha ao carregar credenciais do .env: {e}")
            sys.exit(1)

        try:
            info(f"Abrindo conexão: {nome_logon}...")
            connection = application.OpenConnection(nome_logon, True)

            tentativas = 0
            while connection.Children.Count == 0 and tentativas < 20:
                time.sleep(0.5)
                tentativas += 1

            if connection.Children.Count == 0:
                erro("A sessão SAP não foi iniciada corretamente.")
                _log_alerta_rz11()
                sys.exit(1)

            session = connection.Children(0)

            try:
                if str(session.Info.SystemName).upper() != str(sistema_desejado).upper():
                    erro(f"A sessão SAP aberta não pertence ao ambiente '{ambiente_cockpit}'.")
                    sys.exit(1)
            except Exception:
                pass

            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = cliente_esperado
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
            session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
            session.findById("wnd[0]").sendVKey(0)

        except Exception as e:
            if _is_scripting_disabled_error(e):
                _erro_scripting_inativo(e)
            else:
                erro("Erro ao tentar abrir a ligação SAP.")
                erro("Verifique se o servidor SAP está ligado e disponível.")
                erro(f"Detalhes técnicos: {e}")
                sys.exit(1)

        if not _aguardar_login(session, cliente_esperado, timeout_s=20):
            erro("Login não foi confirmado (User/Client não disponíveis).")
            warn("Verifique se o SAP pediu pop-up, senha extra, ou se o login falhou.")
            sys.exit(1)

    else:
        warn("Existe uma sessão SAP ativa, mas não é do ambiente selecionado.")
        info(f"Abra manualmente a ligação do ambiente '{nome_logon}' no SAP Logon e faça login (se necessário).")
        input("Pressione ENTER quando a sessão do ambiente estiver aberta...")

        session, connection = _encontrar_sessao_do_sistema(application, sistema_desejado)
        if session is None:
            erro("Não foi encontrada uma sessão do ambiente selecionado após confirmação. A encerrar.")
            sys.exit(1)

if not _is_sap_logado(session, cliente_esperado):
    warn(f"Sessão do ambiente '{ambiente_cockpit}' encontrada, mas sem login ou client incorreto (esperado: {cliente_esperado}).")
    info("Faça o login manualmente no SAP GUI (e ajuste o client se necessário).")
    input("Pressione ENTER assim que tiver terminado o login...")

    if not _aguardar_login(session, cliente_esperado, timeout_s=20):
        erro("O login ainda não foi detectado corretamente após confirmação. A encerrar.")
        sys.exit(1)

_log_scripting_status_apenas_quando_logado(session, cliente_esperado)

try:
    mostrar_titulo(
        ambiente=f"{ambiente_cockpit} ({nome_logon})",
        sistema=session.Info.SystemName,
        cliente=session.Info.Client,
        utilizador=session.Info.User,
    )
    ok(f"Sessão SAP pronta no sistema: {session.Info.SystemName}")
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