###################################################################################
# SAP Cockpit - versao pronta para Terminal e Web Worker
###################################################################################

from __future__ import annotations

import importlib.util
import inspect
import msvcrt
import os
import re
import subprocess
import sys
import time
import traceback
from typing import Any

import pythoncom
import pywintypes
import win32com.client

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


class SapCockpitError(RuntimeError):
    pass


###################################################################################
# Regra central do projeto: STATUS sempre vem de wnd[0]/sbar
###################################################################################


def read_sbar_status(session) -> str:
    try:
        return str(session.findById("wnd[0]/sbar").Text).strip()
    except Exception as exc:
        return f"Nao foi possivel ler STATUS em wnd[0]/sbar: {exc}"


def _raise_or_exit(message: str, interactive: bool, exc: Exception | None = None):
    if exc:
        message = f"{message}\nDetalhes tecnicos: {exc}"
    if interactive:
        erro(message)
        sys.exit(1)
    raise SapCockpitError(message)


###################################################################################
# Funcoes auxiliares (I/O, .env e menus)
###################################################################################


def limpar_buffer_teclado():
    try:
        while msvcrt.kbhit():
            msvcrt.getch()
    except Exception:
        pass


def ler_texto(prompt: str, oculto: bool = False) -> str:
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


def _parse_env_line(linha_txt: str):
    if not linha_txt:
        return None, None

    linha_txt = linha_txt.strip()
    if not linha_txt or linha_txt.startswith("#"):
        return None, None

    if "=" not in linha_txt:
        return None, None

    chave, valor = linha_txt.split("=", 1)
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
            for linha_txt in f:
                chave, valor = _parse_env_line(linha_txt)
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
        warn("Ficheiro .env nao encontrado nos caminhos esperados.")


def _obter_credenciais_env(sistema_desejado: str, cliente_esperado: str) -> tuple[str, str, str, str]:
    _carregar_dotenv()

    sistema = str(sistema_desejado or "").strip().upper()
    cliente = str(cliente_esperado or "").strip()
    usuario = os.getenv("SAP_USER", "").strip()
    idioma = os.getenv("SAP_LANGUAGE", "PT").strip() or "PT"

    chave_password = f"SAP_PASSWORD_{sistema}CLNT{cliente}"
    senha = os.getenv(chave_password, "").strip()

    if not usuario:
        raise RuntimeError("Variavel SAP_USER nao encontrada ou vazia no ficheiro .env.")

    if not senha:
        raise RuntimeError(f"Variavel '{chave_password}' nao encontrada ou vazia no ficheiro .env.")

    return usuario, senha, idioma, chave_password


def selecionar_ambiente(payload: dict[str, Any] | None = None, interactive: bool = True) -> str:
    payload = payload or {}
    ambiente_payload = str(payload.get("ambiente") or "").strip().upper()
    if ambiente_payload:
        if ambiente_payload in MAPA_SISTEMA:
            return ambiente_payload
        raise SapCockpitError(f"Ambiente '{ambiente_payload}' invalido recebido pela web.")

    if len(sys.argv) > 1:
        ambiente_recebido = sys.argv[1].upper()
        if ambiente_recebido in MAPA_SISTEMA:
            return ambiente_recebido
        if interactive:
            erro(f"Ambiente '{ambiente_recebido}' invalido via argumento.")
            sys.exit(1)
        raise SapCockpitError(f"Ambiente '{ambiente_recebido}' invalido via argumento.")

    if not interactive:
        raise SapCockpitError("Ambiente nao informado no payload. Ex.: {'ambiente': 'S4Q'}")

    mostrar_titulo()
    mostrar_ambientes(AMBIENTES)

    opcao = input("\nDigite o numero do ambiente que deseja atualizar: ").strip()
    if opcao not in AMBIENTES:
        erro("Opcao invalida.")
        sys.exit(1)

    return AMBIENTES[opcao][0]


def _resolve_processo_path(processo: str) -> str:
    processo = str(processo or "").strip()
    if not processo:
        raise SapCockpitError("Processo/pasta nao informado no payload.")

    if os.path.isabs(processo):
        caminho = processo
    else:
        caminho = os.path.join(PROCESSOS_DIR, processo)

    caminho = os.path.abspath(caminho)
    processos_abs = os.path.abspath(PROCESSOS_DIR)

    if not caminho.startswith(processos_abs):
        raise SapCockpitError("Processo fora da pasta PROCESSOS_DIR nao e permitido.")

    if not os.path.isdir(caminho):
        raise SapCockpitError(f"Pasta de processo nao encontrada: {caminho}")

    return caminho


def selecionar_pasta_processo(payload: dict[str, Any] | None = None, interactive: bool = True):
    payload = payload or {}
    processo_payload = str(payload.get("processo") or "").strip()
    if processo_payload:
        return _resolve_processo_path(processo_payload)

    if not interactive:
        raise SapCockpitError("Processo nao informado no payload.")

    pastas = sorted(
        [
            p for p in os.listdir(PROCESSOS_DIR)
            if os.path.isdir(os.path.join(PROCESSOS_DIR, p)) and p != "__pycache__"
        ]
    )

    mostrar_processos(pastas)

    try:
        escolha = int(input("\nDigite o numero do processo que deseja abrir: ")) - 1
        pasta_escolhida = pastas[escolha]
        return os.path.join(PROCESSOS_DIR, pasta_escolhida)
    except (ValueError, IndexError):
        erro("Selecao invalida.")
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
        erro(f"Ficheiro nao encontrado: {PESQUISAR_REQUEST_PATH}")
        return None

    spec = importlib.util.spec_from_file_location("pesquisar_request", PESQUISAR_REQUEST_PATH)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def escolher_request_por_linha(lista_resultados):
    if not lista_resultados:
        return ("", "")

    while True:
        raw = input(f"\nDigite o numero da linha da request (1-{len(lista_resultados)}): ").strip()
        if not raw.isdigit():
            erro("Digite apenas numeros.")
            continue

        idx = int(raw)
        if idx < 1 or idx > len(lista_resultados):
            erro("Numero fora do intervalo.")
            continue

        trkorr, as4text = lista_resultados[idx - 1]
        return (trkorr, as4text)


def _resetar_env_request():
    os.environ["SAP_REQUEST_OPTION"] = ""
    os.environ["SAP_REQUEST_NUMBER"] = ""
    os.environ["SAP_REQUEST_DESC"] = ""
    os.environ["SAP_SEARCH_TEXT"] = ""


###################################################################################
# SAP GUI helpers (SE10 - criar request)
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
        raise RuntimeError(f"Elemento SAP nao encontrado: {sap_id}")
    obj.press()
    _sleep(SLEEP_ACTION)


def _send_vkey(session, vkey: int):
    session.findById("wnd[0]").sendVKey(vkey)
    _sleep(SLEEP_ACTION)


def _set_text(session, sap_id: str, text: str, caret_pos: int | None = None):
    obj = _safe_find(session, sap_id)
    if not obj:
        raise RuntimeError(f"Campo SAP nao encontrado: {sap_id}")
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
        raise RuntimeError("Nao foi possivel localizar o campo de comando (okcd).")

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


def _criar_nova_request_no_sap(
    session,
    tipo: str | None = None,
    desc: str | None = None,
    interactive: bool = True,
) -> tuple[str, str, str]:
    _ensure_se10(session)

    if interactive:
        linha()
        destaque("CRIAR NOVA REQUEST")
        print("\nTipo da ordem:")
        print("1 - Ordem customizing")
        print("2 - Ordem workbench")

    while True:
        if tipo is None and interactive:
            tipo = input("Digite a opcao (1/2): ").strip()
        tipo = str(tipo or "1").strip()
        if tipo in ("1", "2"):
            break
        if not interactive:
            raise SapCockpitError("Tipo de request invalido. Use '1' ou '2'.")
        erro("Opcao invalida. Use apenas 1 ou 2.")
        tipo = None

    if desc is None and interactive:
        desc = input("Descricao da request (max 60): ").strip()

    desc = str(desc or "REQUEST CRIADA VIA SCRIPT").strip()
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
    info(f"Descricao: {desc}")

    if not trkorr:
        if not interactive:
            raise SapCockpitError(
                "A request foi criada, mas nao consegui extrair o numero automaticamente. "
                f"STATUS: {read_sbar_status(session)}"
            )
        trkorr = input("Nao consegui extrair a request automaticamente. Cole aqui (ex.: S4QK900416): ").strip().upper()

    info(f"Request: {trkorr}")
    return (trkorr, desc, tipo_txt)


###################################################################################
# Request menu
###################################################################################


def perguntar_opcao_request(
    sistema_desejado: str,
    session,
    payload: dict[str, Any] | None = None,
    interactive: bool = True,
) -> dict:
    payload = payload or {}

    if not interactive:
        opc = str(payload.get("request_option") or "4").strip()
        if opc == "":
            opc = "4"
        if opc not in ("1", "2", "4"):
            raise SapCockpitError("Na web, request_option suportado: 1, 2 ou 4.")

        ctx = {
            "request_option": opc,
            "request_number": "",
            "request_desc": "",
            "search_text": "",
        }

        if opc == "1":
            num = validar_request(str(payload.get("request_number") or ""))
            if not num:
                raise SapCockpitError("Request invalida ou vazia para request_option=1.")
            ctx["request_number"] = num

        elif opc == "2":
            trkorr, desc, _tipo_txt = _criar_nova_request_no_sap(
                session,
                tipo=str(payload.get("request_type") or "1"),
                desc=str(payload.get("request_desc") or "REQUEST CRIADA VIA WEB"),
                interactive=False,
            )
            ctx["request_number"] = validar_request(trkorr) or trkorr.strip().upper()
            ctx["request_desc"] = desc

        elif opc == "4":
            ctx["request_number"] = ""

        os.environ["SAP_REQUEST_OPTION"] = ctx["request_option"]
        os.environ["SAP_REQUEST_NUMBER"] = ctx["request_number"]
        os.environ["SAP_REQUEST_DESC"] = ctx["request_desc"]
        os.environ["SAP_SEARCH_TEXT"] = ctx["search_text"]

        return ctx

    linha()
    destaque("OPCOES DE TRANSPORTE")
    print("1 - Escreva o numero da Request")
    print("2 - Criar nova ordem de transporte")
    print("3 - Pesquisar suas request criadas")
    print("4 - Prima [Enter] vazio para NAO transportar")
    linha()

    while True:
        opc = input("\nOpcao: ").strip()

        if opc in ("1", "2", "3", "4", ""):
            if opc == "":
                opc = "4"
            break
        erro("Opcao invalida. Use 1, 2, 3, 4 ou apenas pressione Enter.")

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
            erro("Request invalida. Exemplo valido: S4QK900396")

    elif opc == "2":
        trkorr, desc, _tipo_txt = _criar_nova_request_no_sap(session, interactive=True)
        ctx["request_number"] = validar_request(trkorr) or trkorr.strip().upper()
        ctx["request_desc"] = desc

    elif opc == "3":
        mod_pesq = carregar_pesquisar_request()
        if not mod_pesq or not hasattr(mod_pesq, "listar_requests"):
            erro("Modulo pesquisar_request.py nao carregado ou funcao listar_requests nao encontrada.")
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
# Execucao de processos
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


def _build_exec_kwargs(
    processo_escolhido: str,
    exec_fn,
    sistema_desejado: str,
    session,
    payload: dict[str, Any] | None = None,
    interactive: bool = True,
) -> dict[str, Any]:
    payload = payload or {}
    info_exec = _analisar_exec_signature(exec_fn)

    precisa_request_agora = False
    if info_exec["p_request_ctx"] and not info_exec["request_ctx_is_optional"]:
        precisa_request_agora = True
    elif info_exec["p_request_transp"] and not info_exec["request_transp_is_optional"] and not info_exec["file_is_optional"]:
        precisa_request_agora = True

    request_ctx = {"request_option": "", "request_number": "", "request_desc": "", "search_text": ""}
    if precisa_request_agora:
        request_ctx = perguntar_opcao_request(sistema_desejado, session, payload=payload, interactive=interactive)
    elif not interactive and str(payload.get("request_option") or "").strip() in ("1", "2", "4"):
        request_ctx = perguntar_opcao_request(sistema_desejado, session, payload=payload, interactive=False)

    kwargs: dict[str, Any] = {}

    if info_exec["p_pfcg"]:
        nome_sem_ext = processo_escolhido.replace(".py", "").strip()
        if "." in nome_sem_ext:
            aba_calculada = nome_sem_ext.split(".", 1)[1].strip()
        else:
            aba_calculada = nome_sem_ext

        info(f"Aba do Excel detetada automaticamente: '{aba_calculada}'")
        kwargs[info_exec["p_pfcg"].name] = aba_calculada

    if info_exec["p_file"] and not info_exec["file_is_optional"]:
        if not interactive:
            path = str(payload.get("caminho_ficheiro") or payload.get("file_path") or "").strip()
            if not path:
                raise SapCockpitError(
                    f"O subprocesso '{processo_escolhido}' pede ficheiro, mas caminho_ficheiro nao foi enviado pela web."
                )
            if not os.path.exists(path):
                raise SapCockpitError(f"Ficheiro nao encontrado no worker Windows: {path}")
            kwargs[info_exec["p_file"].name] = path
        else:
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
                    erro("Operacao cancelada (ficheiro nao selecionado).")
                    raise SapCockpitError("Ficheiro nao selecionado.")
                kwargs[info_exec["p_file"].name] = path
            except Exception as e:
                erro(f"Falha ao abrir popup de ficheiro: {e}")
                raise

    if info_exec["p_request_ctx"]:
        kwargs[info_exec["p_request_ctx"].name] = request_ctx

    if info_exec["p_request_transp"] and request_ctx.get("request_number"):
        kwargs[info_exec["p_request_transp"].name] = request_ctx["request_number"]

    return kwargs


def _load_process_module(caminho_script: str):
    spec = importlib.util.spec_from_file_location("modulo_processo", caminho_script)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _executar_um_script(
    ambiente_cockpit,
    caminho_pasta,
    processo_escolhido,
    sistema_desejado,
    session,
    payload: dict[str, Any] | None = None,
    interactive: bool = True,
):
    _resetar_env_request()
    ok(f"Processo selecionado: {processo_escolhido}")

    caminho_script = os.path.join(caminho_pasta, processo_escolhido)
    if not os.path.exists(caminho_script):
        raise SapCockpitError(f"Subprocesso nao encontrado: {caminho_script}")

    modulo = _load_process_module(caminho_script)

    if not hasattr(modulo, "executar"):
        raise SapCockpitError(f"O ficheiro '{processo_escolhido}' nao contem a funcao executar().")

    exec_fn = modulo.executar
    kwargs = _build_exec_kwargs(
        processo_escolhido=processo_escolhido,
        exec_fn=exec_fn,
        sistema_desejado=sistema_desejado,
        session=session,
        payload=payload,
        interactive=interactive,
    )

    try:
        exec_fn(ambiente_cockpit, **kwargs)
    except TypeError:
        exec_fn(ambiente_cockpit)

    return read_sbar_status(session)


def executar_processo(
    ambiente_cockpit,
    caminho_pasta,
    sistema_desejado,
    session,
    payload: dict[str, Any] | None = None,
    interactive: bool = True,
):
    payload = payload or {}

    if not interactive:
        subprocesso = str(payload.get("subprocesso") or payload.get("script") or "").strip()
        if not subprocesso:
            raise SapCockpitError("Subprocesso/ficheiro .py nao informado no payload.")
        if not subprocesso.lower().endswith(".py"):
            subprocesso = f"{subprocesso}.py"
        return _executar_um_script(
            ambiente_cockpit=ambiente_cockpit,
            caminho_pasta=caminho_pasta,
            processo_escolhido=subprocesso,
            sistema_desejado=sistema_desejado,
            session=session,
            payload=payload,
            interactive=False,
        )

    while True:
        scripts_py = sorted([f for f in os.listdir(caminho_pasta) if f.endswith(".py") and not f.startswith("~$")])
        if not scripts_py:
            erro("Nenhum script .py encontrado na pasta selecionada.")
            return

        mostrar_subprocessos(scripts_py)

        try:
            escolha = int(input("\nDigite o numero do sub-processo que deseja executar: ")) - 1
            if escolha == len(scripts_py):
                return "voltar"
            processo_escolhido = scripts_py[escolha]
        except (ValueError, IndexError):
            erro("Selecao invalida. Tente novamente.")
            continue

        try:
            _executar_um_script(
                ambiente_cockpit=ambiente_cockpit,
                caminho_pasta=caminho_pasta,
                processo_escolhido=processo_escolhido,
                sistema_desejado=sistema_desejado,
                session=session,
                payload=payload,
                interactive=True,
            )
        except Exception as e:
            erro(f"Erro ao executar '{processo_escolhido}': {e}")
            continue


###################################################################################
# Scripting / sessao SAP
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
        "desativ", "inativ", "nao esta ativo", "não está ativo"
    ])
    return has_scripting and has_disabled


def _log_alerta_rz11():
    warn(MSG_RZ11_SCRIPTING)


def _erro_scripting_inativo(exc: Exception | None = None, interactive: bool = True):
    msg = "O scripting do SAP GUI nao esta ativo ou nao foi possivel inicializar o objeto SAPGUI."
    if interactive:
        erro(msg)
        _log_alerta_rz11()
        if exc:
            erro(f"Detalhes tecnicos: {exc}")
        sys.exit(1)
    raise SapCockpitError(f"{msg}\n{MSG_RZ11_SCRIPTING}\nDetalhes tecnicos: {exc}")


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
        ok("SAP GUI Scripting esta ativo.")


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


def obter_sessao_sap(ambiente_cockpit: str, interactive: bool = True):
    sistema_desejado = MAPA_SISTEMA.get(ambiente_cockpit)
    cliente_esperado = CLIENTES_POR_AMBIENTE.get(ambiente_cockpit, "100")
    nome_logon = dict((v[0], v[1]) for v in AMBIENTES.values()).get(ambiente_cockpit, ambiente_cockpit)

    if interactive:
        mostrar_titulo(
            ambiente=f"{ambiente_cockpit} ({nome_logon})",
            sistema=sistema_desejado,
            cliente=cliente_esperado,
        )
        info("A verificar se ja existe uma sessao aberta no ambiente desejado...")

    try:
        pythoncom.CoInitialize()
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            if not interactive:
                raise
            warn("SAP Logon nao detectado. Iniciando executavel...")
            subprocess.Popen(SAPLOGON_PATH)
            time.sleep(5)
            SapGuiAuto = win32com.client.GetObject("SAPGUI")

        application = SapGuiAuto.GetScriptingEngine
        if not application:
            raise RuntimeError("Scripting Engine desativado (RZ11).")

    except Exception as e:
        _erro_scripting_inativo(e, interactive=interactive)

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
                _raise_or_exit(f"Falha ao carregar credenciais do .env: {e}", interactive)

            try:
                info(f"Abrindo conexao: {nome_logon}...")
                connection = application.OpenConnection(nome_logon, True)

                tentativas = 0
                while connection.Children.Count == 0 and tentativas < 20:
                    time.sleep(0.5)
                    tentativas += 1

                if connection.Children.Count == 0:
                    _raise_or_exit("A sessao SAP nao foi iniciada corretamente.", interactive)

                session = connection.Children(0)

                try:
                    if str(session.Info.SystemName).upper() != str(sistema_desejado).upper():
                        _raise_or_exit(f"A sessao SAP aberta nao pertence ao ambiente '{ambiente_cockpit}'.", interactive)
                except SapCockpitError:
                    raise
                except Exception:
                    pass

                session.findById("wnd[0]/usr/txtRSYST-MANDT").text = cliente_esperado
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
                session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
                session.findById("wnd[0]").sendVKey(0)

            except Exception as e:
                if _is_scripting_disabled_error(e):
                    _erro_scripting_inativo(e, interactive=interactive)
                else:
                    _raise_or_exit(
                        "Erro ao tentar abrir a ligacao SAP. Verifique se o servidor SAP esta ligado e disponivel.",
                        interactive,
                        e,
                    )

            if not _aguardar_login(session, cliente_esperado, timeout_s=20):
                _raise_or_exit(
                    "Login nao foi confirmado (User/Client nao disponiveis). Verifique pop-up, senha extra, ou falha de login.",
                    interactive,
                )

        else:
            if not interactive:
                raise SapCockpitError(
                    f"Existe uma sessao SAP ativa, mas nao e do ambiente selecionado '{nome_logon}'. "
                    "Abra esse ambiente no SAP GUI ou feche as sessoes incorretas."
                )

            warn("Existe uma sessao SAP ativa, mas nao e do ambiente selecionado.")
            info(f"Abra manualmente a ligacao do ambiente '{nome_logon}' no SAP Logon e faca login (se necessario).")
            input("Pressione ENTER quando a sessao do ambiente estiver aberta...")

            session, connection = _encontrar_sessao_do_sistema(application, sistema_desejado)
            if session is None:
                _raise_or_exit("Nao foi encontrada uma sessao do ambiente selecionado apos confirmacao.", interactive)

    if not _is_sap_logado(session, cliente_esperado):
        if not interactive:
            raise SapCockpitError(
                f"Sessao do ambiente '{ambiente_cockpit}' encontrada, mas sem login ou client incorreto. "
                f"Client esperado: {cliente_esperado}."
            )

        warn(f"Sessao do ambiente '{ambiente_cockpit}' encontrada, mas sem login ou client incorreto (esperado: {cliente_esperado}).")
        info("Faca o login manualmente no SAP GUI (e ajuste o client se necessario).")
        input("Pressione ENTER assim que tiver terminado o login...")

        if not _aguardar_login(session, cliente_esperado, timeout_s=20):
            _raise_or_exit("O login ainda nao foi detectado corretamente apos confirmacao.", interactive)

    _log_scripting_status_apenas_quando_logado(session, cliente_esperado)

    try:
        if interactive:
            mostrar_titulo(
                ambiente=f"{ambiente_cockpit} ({nome_logon})",
                sistema=session.Info.SystemName,
                cliente=session.Info.Client,
                utilizador=session.Info.User,
            )
            ok(f"Sessao SAP pronta no sistema: {session.Info.SystemName}")
    except Exception:
        pass

    return session, connection


###################################################################################
# Entradas: Web Worker e Terminal
###################################################################################


def run_sap_cockpit(payload: dict[str, Any] | None = None) -> dict[str, str]:
    payload = payload or {}
    log_lines: list[str] = []
    session = None

    try:
        ambiente_cockpit = selecionar_ambiente(payload=payload, interactive=False)
        sistema_desejado = MAPA_SISTEMA.get(ambiente_cockpit)

        log_lines.append(f"Ambiente: {ambiente_cockpit}")
        log_lines.append(f"Sistema: {sistema_desejado}")

        session, _connection = obter_sessao_sap(ambiente_cockpit, interactive=False)
        log_lines.append(f"Sessao SAP: {session.Info.SystemName} / client {session.Info.Client} / user {session.Info.User}")

        caminho_processo = selecionar_pasta_processo(payload=payload, interactive=False)
        log_lines.append(f"Processo: {caminho_processo}")

        status = executar_processo(
            ambiente_cockpit,
            caminho_pasta=caminho_processo,
            sistema_desejado=sistema_desejado,
            session=session,
            payload=payload,
            interactive=False,
        )

        status = status or read_sbar_status(session)
        log_lines.append(f"STATUS: {status}")

        return {
            "status": status,
            "log": "\n".join(log_lines),
        }

    except Exception as exc:
        status = read_sbar_status(session) if session is not None else str(exc)
        log_lines.append(traceback.format_exc())
        return {
            "status": status or str(exc),
            "log": "\n".join(log_lines),
        }


def main_terminal():
    ambiente_cockpit = selecionar_ambiente(interactive=True)
    sistema_desejado = MAPA_SISTEMA.get(ambiente_cockpit)

    session, _connection = obter_sessao_sap(ambiente_cockpit, interactive=True)

    while True:
        caminho_processo = selecionar_pasta_processo(interactive=True)
        resultado = executar_processo(
            ambiente_cockpit,
            caminho_pasta=caminho_processo,
            sistema_desejado=sistema_desejado,
            session=session,
            interactive=True,
        )
        if resultado != "voltar":
            continue


if __name__ == "__main__":
    main_terminal()
