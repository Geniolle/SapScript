# -*- coding: utf-8 -*-

###################################################################################
# PROCESSO: Adicionar Função SU01/SU10  (sheet = nome do .py SEM o prefixo)
# Ex.: "H. CUA_ADICIONAR.py"  →  Sheet "CUA_ADICIONAR"
#
# ESTRUTURA ESPERADA DA SHEET:
# ID | UTILIZADOR | SISTEMA | AGR_NAME | STATUS | MSG | TIMESTEMP
#
# PADRÃO APLICADO:
# - STATUS final decidido com base no wnd[0]/sbar
# - captura do sbar em cada passo crítico
# - guarda a última mensagem relevante do SAP
# - grava no Excel atualizando APENAS STATUS / MSG / TIMESTEMP
# - preserva formatação, fórmulas, filtros e restantes colunas
# - popup sem diretório fixo (abre no último local usado)
###################################################################################

###################################################################################
# BLOCO 1: IMPORTAÇÕES
###################################################################################

import os
import time
import unicodedata
from datetime import datetime

import pandas as pd
import win32com.client
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

###################################################################################
# BLOCO 2: NOME DO SCRIPT / SHEET / MAPA DE SISTEMAS
###################################################################################

try:
    NOME_SCRIPT = os.path.splitext(os.path.basename(__file__))[0]
except NameError:
    NOME_SCRIPT = "H. CUA_ADICIONAR"

NOME_SHEET = NOME_SCRIPT.split(".", 1)[-1].strip() if "." in NOME_SCRIPT else NOME_SCRIPT

MAPA_SISTEMA = {
    "DEV": "S4D",
    "QAD": "S4Q",
    "PRD": "S4P",
    "CUA": "SPA",
}

###################################################################################
# BLOCO 3: UTILITÁRIOS
###################################################################################

def formatar_tempo(segundos):
    m = int(segundos // 60)
    s = int(segundos % 60)
    return f"{m:02d}m {s:02d}s"


def agora_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalizar_coluna(valor):
    return (
        unicodedata.normalize("NFKD", str(valor))
        .encode("ASCII", "ignore")
        .decode("utf-8")
        .strip()
        .upper()
    )


def normalizar_valor(valor):
    return (
        unicodedata.normalize("NFKD", str(valor))
        .encode("ASCII", "ignore")
        .decode("utf-8")
        .strip()
        .upper()
    )


def texto_limpo(valor):
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    if txt.lower() in ("nan", "none", "<na>"):
        return ""
    return txt


def valor_vazio(valor):
    return texto_limpo(valor) == ""


def chave_id(valor):
    if pd.isna(valor):
        return ""

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        return str(int(valor)) if valor.is_integer() else str(valor).strip()

    txt = str(valor).strip()

    if txt.endswith(".0"):
        base = txt[:-2]
        if base.isdigit():
            return base

    return txt


def mapear_cabecalhos_openpyxl(ws):
    mapa = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if val is None:
            continue
        mapa[normalizar_coluna(val)] = c
    return mapa


###################################################################################
# BLOCO 4: SELEÇÃO DO FICHEIRO
###################################################################################

def selecionar_ficheiro_excel():
    """
    Popup sem diretório fixo.
    O Windows tende a abrir no último local utilizado.
    """
    try:
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        root.attributes("-topmost", True)

        caminho = filedialog.askopenfilename(
            title="Selecione o ficheiro Excel",
            filetypes=[
                ("Ficheiros Excel", "*.xlsx *.xlsm"),
                ("Todos os ficheiros", "*.*"),
            ],
        )

        root.destroy()

        if not caminho:
            print("⚠️ Seleção cancelada pelo utilizador.")
            return None

        ext = os.path.splitext(caminho)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            print("❌ Apenas ficheiros .xlsx e .xlsm são suportados neste processo.")
            return None

        print(f"✅ Ficheiro a processar: {caminho}")
        return caminho

    except Exception as e:
        print(f"❌ Erro ao abrir o popup: {e}")
        return None


###################################################################################
# BLOCO 5: LEITURA DO EXCEL
###################################################################################

def ler_ficheiro(caminho_ficheiro, nome_sheet):
    """
    Lê a sheet alvo, normaliza cabeçalhos e valida estrutura obrigatória.
    Harmoniza variantes para:
    ID | UTILIZADOR | SISTEMA | AGR_NAME | STATUS | MSG | TIMESTEMP
    """
    if not caminho_ficheiro or not os.path.exists(caminho_ficheiro):
        print("❌ Caminho inválido ou ficheiro inexistente.")
        return None

    try:
        ext = os.path.splitext(caminho_ficheiro)[1].lower()
        keep_vba = ext == ".xlsm"

        # 1. Carregar em openpyxl para verificar e criar cabeçalhos em falta
        wb = load_workbook(
            caminho_ficheiro,
            read_only=False,
            data_only=False,
            keep_vba=keep_vba,
        )
        sheets = wb.sheetnames
        if nome_sheet not in sheets:
            print(f"❌ Sheet '{nome_sheet}' não encontrada. Disponíveis: {', '.join(sheets)}")
            wb.close()
            return None

        ws = wb[nome_sheet]
        
        # Obter os cabeçalhos reais da linha 1
        headers_linha_1 = [c.value for c in ws[1]]
        headers_norm = [normalizar_coluna(h) if h is not None else "" for h in headers_linha_1]

        renames = {
            "USER": "UTILIZADOR",
            "USERNAME": "UTILIZADOR",
            "SYSTEM": "SISTEMA",
            "FUNCAO": "AGR_NAME",
            "FUNÇÃO": "AGR_NAME",
            "ROLE": "AGR_NAME",
            "NOME FUNCAO": "AGR_NAME",
            "NOME FUNÇAO": "AGR_NAME",
            "NOME FUNÇÂO": "AGR_NAME",
            "AGRNAME": "AGR_NAME",
            "TIMESTAMP": "TIMESTEMP",
        }

        resolved_headers = []
        for h in headers_norm:
            resolved_headers.append(renames.get(h, h))

        # Colunas de entrada obrigatórias
        colunas_entrada = {"ID", "UTILIZADOR", "SISTEMA", "AGR_NAME"}
        # Colunas de saída que podem ser criadas
        colunas_saida = {"STATUS": "STATUS", "MSG": "MSG", "TIMESTEMP": "TIMESTEMP"}

        # Verificar se as colunas de entrada existem
        falta_entrada = [c for c in colunas_entrada if c not in resolved_headers]
        if falta_entrada:
            print(f"❌ Colunas obrigatórias de entrada em falta: {', '.join(falta_entrada)}")
            wb.close()
            return None

        # Se as colunas de entrada existem, podemos prosseguir e criar as colunas de saída em falta
        modificou = False
        last_col = len(headers_norm)
        for col_key, col_name in colunas_saida.items():
            if col_key not in resolved_headers:
                last_col += 1
                ws.cell(row=1, column=last_col, value=col_name)
                print(f"➕ Coluna '{col_name}' criada automaticamente na coluna {last_col}.")
                modificou = True

        # Validar e remover espaços do nome da Role (AGR_NAME) diretamente no Excel
        col_idx_agr = resolved_headers.index("AGR_NAME") + 1
        modificou_roles = False
        for r in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=col_idx_agr).value
            if cell_val is not None:
                agr = str(cell_val).strip()
                if " " in agr:
                    orig = agr
                    agr = agr.replace(" ", "")
                    ws.cell(row=r, column=col_idx_agr, value=agr)
                    print(f"⚠️ Espaços detetados e removidos da Role: '{orig}' -> '{agr}' (atualizado no Excel)")
                    modificou_roles = True

        if modificou or modificou_roles:
            try:
                wb.save(caminho_ficheiro)
                print("💾 Ficheiro Excel guardado com as correções de cabeçalhos/roles.")
            except Exception as e:
                wb.close()
                print(f"❌ Erro ao guardar o ficheiro Excel com as novas colunas/roles: {e}")
                print("💡 Certifique-se de que o ficheiro Excel não está aberto no Excel e tente novamente.")
                return None

        wb.close()

        # 2. Carregar o DataFrame em pandas para uso no script
        df = pd.read_excel(caminho_ficheiro, sheet_name=nome_sheet, dtype=object)
        df.columns = [normalizar_coluna(c) for c in df.columns]

        df.rename(
            columns={
                "USER": "UTILIZADOR",
                "USERNAME": "UTILIZADOR",
                "SYSTEM": "SISTEMA",
                "FUNCAO": "AGR_NAME",
                "FUNÇÃO": "AGR_NAME",
                "ROLE": "AGR_NAME",
                "NOME FUNCAO": "AGR_NAME",
                "NOME FUNÇAO": "AGR_NAME",
                "NOME FUNÇÂO": "AGR_NAME",
                "AGRNAME": "AGR_NAME",
                "TIMESTAMP": "TIMESTEMP",
            },
            inplace=True,
        )

        obrigatorias = ["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP"]
        falta = [c for c in obrigatorias if c not in df.columns]
        if falta:
            print(f"❌ Colunas obrigatórias em falta: {', '.join(falta)}")
            return None

        for c in ["UTILIZADOR", "SISTEMA", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP"]:
            df[c] = df[c].apply(texto_limpo)

        df["CHAVE_ID"] = df["ID"].apply(chave_id)

        print(f"📄 Sheet carregada: '{nome_sheet}' | Registos: {len(df)}")
        return df

    except Exception as e:
        print(f"❌ Erro ao ler a sheet: {e}")
        return None


###################################################################################
# BLOCO 6: SAP GUI / STATUS BAR / POPUPS
###################################################################################

def conectar_sap(sistema_desejado):
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        for conn in application.Children:
            for sess in conn.Children:
                try:
                    if texto_limpo(sess.Info.SystemName).upper() == sistema_desejado:
                        print(
                            f"✅ Conectado: {sess.Info.SystemName} "
                            f"| User: {sess.Info.User} "
                            f"| Cliente: {sess.Info.Client}"
                        )
                        return sess
                except Exception:
                    continue

        print(f"❌ Sessão SAP não encontrada para o sistema {sistema_desejado}.")
        return None

    except Exception as e:
        print(f"❌ Erro ao conectar SAP GUI: {e}")
        return None


def esperar_elemento(session, element_id, tentativas=20, espera=0.5):
    for _ in range(tentativas):
        try:
            return session.findById(element_id)
        except Exception:
            time.sleep(espera)
    return None


def existe_elemento(session, element_id):
    try:
        session.findById(element_id)
        return True
    except Exception:
        return False


def ir_para_transacao(session, tcode):
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/N{tcode}"
    session.findById("wnd[0]").sendVKey(0)


def voltar_para_inicio(session):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        session.findById("wnd[0]").sendVKey(0)
    except Exception:
        pass


def ler_status_bar_once(session):
    """
    Lê uma vez o wnd[0]/sbar.
    """
    try:
        sbar = session.findById("wnd[0]/sbar")
        tipo = texto_limpo(getattr(sbar, "MessageType", ""))
        texto = texto_limpo(getattr(sbar, "Text", ""))
        return tipo, texto
    except Exception:
        return "", ""


def registar_evento_status(eventos, origem, tipo="", texto=""):
    tipo = texto_limpo(tipo)
    texto = texto_limpo(texto)

    if not tipo and not texto:
        return

    eventos.append(
        {
            "origem": texto_limpo(origem),
            "tipo": tipo,
            "texto": texto,
        }
    )


def capturar_status_bar(session, eventos=None, origem="SBAR", tentativas=8, espera=0.25):
    """
    Faz várias tentativas curtas para apanhar a mensagem do wnd[0]/sbar
    no momento certo.
    """
    ultimo_tipo = ""
    ultimo_texto = ""

    for _ in range(tentativas):
        tipo, texto = ler_status_bar_once(session)
        if tipo or texto:
            ultimo_tipo = tipo
            ultimo_texto = texto
            break
        time.sleep(espera)

    if eventos is not None and (ultimo_tipo or ultimo_texto):
        registar_evento_status(eventos, origem, ultimo_tipo, ultimo_texto)

    combinado = (
        f"{ultimo_tipo} - {ultimo_texto}"
        if ultimo_tipo and ultimo_texto
        else (ultimo_texto or ultimo_tipo or "")
    )

    return ultimo_tipo, ultimo_texto, combinado


def obter_titulo_popup(session):
    try:
        return texto_limpo(session.findById("wnd[1]").text)
    except Exception:
        return ""


def tratar_popups_pos_save(session, eventos, max_popups=5):
    """
    Confirma popups após Save.
    Regista o título do popup e volta a tentar capturar o sbar após cada confirmação.
    """
    historico = []

    for n in range(1, max_popups + 1):
        if not existe_elemento(session, "wnd[1]"):
            break

        titulo = obter_titulo_popup(session) or "POPUP"
        historico.append(f"POPUP: {titulo}")
        registar_evento_status(eventos, f"POPUP_{n}", "", titulo)

        try:
            if existe_elemento(session, "wnd[1]/tbar[0]/btn[0]"):
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            elif existe_elemento(session, "wnd[1]/tbar[0]/btn[11]"):
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
            else:
                session.findById("wnd[1]").sendVKey(0)
        except Exception:
            try:
                session.findById("wnd[1]").sendVKey(0)
            except Exception:
                break

        time.sleep(0.35)
        capturar_status_bar(session, eventos, origem=f"SBAR_APOS_POPUP_{n}", tentativas=5, espera=0.20)

    return historico


def obter_ultimo_status_relevante(eventos):
    """
    Procura do fim para o início a última mensagem relevante do sbar.
    """
    for ev in reversed(eventos):
        tipo = texto_limpo(ev.get("tipo", ""))
        texto = texto_limpo(ev.get("texto", ""))
        if tipo or texto:
            combinado = f"{tipo} - {texto}" if tipo and texto else (texto or tipo)
            return tipo, texto, combinado
    return "", "", ""


def resumir_eventos_status(eventos, limite=5):
    """
    Monta uma trilha curta e útil das últimas mensagens.
    """
    itens = []

    for ev in eventos:
        origem = texto_limpo(ev.get("origem", ""))
        tipo = texto_limpo(ev.get("tipo", ""))
        texto = texto_limpo(ev.get("texto", ""))

        if tipo and texto:
            desc = f"{origem}: {tipo} - {texto}"
        elif texto:
            desc = f"{origem}: {texto}"
        elif tipo:
            desc = f"{origem}: {tipo}"
        else:
            continue

        if desc not in itens:
            itens.append(desc)

    if not itens:
        return ""

    return " | ".join(itens[-limite:])


def decidir_status_pelo_historico(eventos):
    """
    Decide STATUS com base no histórico de leituras do wnd[0]/sbar.
    Prioridade:
    - E/A/X => ERRO
    - S/W => CONCLUÍDO
    - fallback por texto
    """
    for ev in reversed(eventos):
        tipo = normalizar_valor(ev.get("tipo", ""))
        texto = normalizar_valor(ev.get("texto", ""))

        if tipo in ("E", "A", "X"):
            return "ERRO"

        if tipo in ("S", "W"):
            return "CONCLUÍDO"

        if any(ch in texto for ch in ["ERRO", "ERROR", "INVALID", "NAO EXIST", "NÃO EXIST", "OBRIGATOR", "INCONSIST"]):
            return "ERRO"

        if any(ch in texto for ch in ["GRAV", "GUARD", "SAVE", "SALV", "ATRIBU", "ATUALIZ", "ALTERACAO EFETUADA", "ALTERAÇÃO EFETUADA"]):
            return "CONCLUÍDO"

    return "ERRO"


def montar_msg_final(eventos):
    """
    A MSG final fica:
    - última mensagem relevante do sbar
    - + trilha curta dos passos, quando útil
    """
    _, _, ultima = obter_ultimo_status_relevante(eventos)
    trilha = resumir_eventos_status(eventos, limite=5)

    if ultima and trilha:
        if trilha.startswith(ultima):
            return trilha
        return f"{ultima} | {trilha}"

    if ultima:
        return ultima

    if trilha:
        return trilha

    return "Sem mensagem relevante do SAP"


def converter_data_sap(data_str) -> datetime:
    """
    Converte uma string de data do SAP (DD.MM.YYYY, YYYYMMDD ou similar) em um objeto datetime.
    """
    cleaned = str(data_str).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%Y%m%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    # Se for a data infinita do SAP (ex: 31.12.9999 ou 99991231 ou sem data), retorna data futura distante
    if "9999" in cleaned or not cleaned:
        return datetime(9999, 12, 31)
    raise ValueError(f"Formato de data invalido: {data_str}")


def _dismiss_popup(session) -> bool:
    """Fecha popups ou avisos se existirem na sessao SAP."""
    for btn_id in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[11]"):
        try:
            session.findById(btn_id).press()
            return True
        except Exception:
            pass
    try:
        session.findById("wnd[1]").sendVKey(12)  # ESC
        return True
    except Exception:
        pass
    return False


def formatar_data_para_sap(data_obj) -> str:
    """
    Formata uma data para o formato padrão aceito pelo SAP (DD.MM.YYYY).
    """
    return data_obj.strftime("%d.%m.%Y")


def descobrir_campos_se16(session) -> dict:
    """
    Descobre dinamicamente os IDs dos campos LOW para BNAME, SUBSYSTEM e TO_DAT
    na tela de seleção da SE16.
    Retorna um dicionário { "BNAME": id, "SUBSYSTEM": id, "TO_DAT": id }

    A SE16 clássica apresenta os labels como GuiTextField com nomes no padrão
    %_In_%_APP_%-TEXT e o campo correspondente como GuiCTextField com nome In-LOW.
    Esta função suporta tanto esse padrão como o padrão GuiLabel convencional.
    """
    import re as _re

    usr = session.findById("wnd[0]/usr")

    # Obter todos os elementos recursivamente
    todos_elementos = []
    stack = [usr]
    while stack:
        curr = stack.pop()
        todos_elementos.append(curr)
        try:
            for idx in range(curr.Children.Count - 1, -1, -1):
                stack.append(curr.Children(idx))
        except Exception:
            pass

    # Inputs: apenas campos -LOW (GuiTextField ou GuiCTextField)
    inputs_low = [
        e for e in todos_elementos
        if getattr(e, "Type", "") in ("GuiTextField", "GuiCTextField")
        and str(getattr(e, "Name", "")).upper().endswith("-LOW")
    ]

    mapeamento = {"BNAME": None, "SUBSYSTEM": None, "TO_DAT": None}

    # Termos de identificação por campo
    termos = {
        "BNAME":     ["BNAME", "USER", "USER NAME", "NOME DO UTILIZADOR", "NOME UTILIZADOR", "CÓDIGO"],
        "SUBSYSTEM": ["SUBSYSTEM", "RECEIVING SYSTEM", "LOGICAL SYSTEM", "SISTEMA RECETOR", "SISTEMA RECEPTOR"],
        "TO_DAT":    ["TO_DAT", "VALID TO", "VALIDADE", "DATA FINAL", "VÁLIDO ATÉ", "VALIDO ATE", "END DATE"],
    }

    # ─────────────────────────────────────────────────────────────────────────
    # Estratégia 1 (mais fiável): padrão SE16 clássica
    # Label: GuiTextField com Name = %_In_%_APP_%-TEXT  → Text contém o nome técnico do campo
    # Input: GuiCTextField ou GuiTextField com Name = In-LOW
    # ─────────────────────────────────────────────────────────────────────────
    for e in todos_elementos:
        if getattr(e, "Type", "") not in ("GuiTextField", "GuiLabel"):
            continue
        name = str(getattr(e, "Name", ""))
        mat = _re.match(r'^%_I(\d+)_%_APP_%', name)
        if not mat:
            continue
        idx_str = mat.group(1)
        e_text = str(getattr(e, "Text", "")).strip().upper()
        e_tooltip = str(getattr(e, "Tooltip", "")).strip().upper()
        e_id = str(getattr(e, "Id", "")).upper()

        # Encontrar o input correspondente (In-LOW)
        inp_name_target = f"I{idx_str}-LOW"
        inp_match = next(
            (i for i in todos_elementos if str(getattr(i, "Name", "")) == inp_name_target),
            None
        )
        if not inp_match:
            continue

        for chave, key_terms in termos.items():
            if mapeamento[chave] is not None:
                continue
            for term in key_terms:
                if term in e_text or term in e_tooltip or term in e_id:
                    mapeamento[chave] = inp_match.Id
                    break

    # ─────────────────────────────────────────────────────────────────────────
    # Estratégia 2: proximidade Top/Left (GuiLabel clássico ou GuiTextField)
    # ─────────────────────────────────────────────────────────────────────────
    if any(v is None for v in mapeamento.values()):
        # Labels: GuiLabel OU GuiTextField que não sejam campos de entrada
        _excluir_names = {"MAX_SEL", "LIST_BRE", "GD-MAXROWS", "MAX_HITS"}
        labels = [
            e for e in todos_elementos
            if getattr(e, "Type", "") == "GuiLabel"
            or (
                getattr(e, "Type", "") == "GuiTextField"
                and not str(getattr(e, "Name", "")).upper().endswith(("-LOW", "-HIGH"))
                and str(getattr(e, "Name", "")).strip() not in _excluir_names
            )
        ]

        def matches_label(lbl, key_terms) -> bool:
            text    = str(getattr(lbl, "Text", "")).strip().upper()
            tooltip = str(getattr(lbl, "Tooltip", "")).strip().upper()
            name    = str(getattr(lbl, "Name", "")).strip().upper()
            lbl_id  = str(getattr(lbl, "Id", "")).upper()
            for term in key_terms:
                if term in text or term in tooltip or term in name or term in lbl_id:
                    return True
            return False

        for chave, key_terms in termos.items():
            if mapeamento[chave] is not None:
                continue

            target_label = None
            for lbl in labels:
                if matches_label(lbl, key_terms):
                    target_label = lbl
                    break

            if not target_label:
                continue

            # Associar por proximidade Top/Left
            lbl_top = lbl_left = None
            try:
                lbl_top  = int(target_label.Top)
                lbl_left = int(target_label.Left)
            except Exception:
                pass

            input_associado = None
            if lbl_top is not None and lbl_left is not None:
                candidatos = []
                for inp in inputs_low:
                    try:
                        inp_top  = int(inp.Top)
                        inp_left = int(inp.Left)
                        if abs(inp_top - lbl_top) <= 10 and inp_left > lbl_left:
                            candidatos.append((inp_left, inp))
                    except Exception:
                        pass
                if candidatos:
                    candidatos.sort(key=lambda x: x[0])
                    input_associado = candidatos[0][1]

            # Fallback por ordem na árvore (próximo -LOW após o label)
            if not input_associado:
                try:
                    idx_lbl = todos_elementos.index(target_label)
                    for j in range(idx_lbl + 1, min(idx_lbl + 8, len(todos_elementos))):
                        cand = todos_elementos[j]
                        if getattr(cand, "Type", "") in ("GuiTextField", "GuiCTextField"):
                            if str(getattr(cand, "Name", "")).upper().endswith("-LOW"):
                                input_associado = cand
                                break
                except Exception:
                    pass

            if input_associado:
                mapeamento[chave] = input_associado.Id

    # ─────────────────────────────────────────────────────────────────────────
    # Estratégia 3 (legado): mapeamento direto por ID técnico do input
    # ─────────────────────────────────────────────────────────────────────────
    for e in inputs_low:
        inp_id = str(getattr(e, "Id", "")).upper()
        if mapeamento["BNAME"] is None and "BNAME" in inp_id:
            mapeamento["BNAME"] = e.Id
        elif mapeamento["SUBSYSTEM"] is None and ("SUBSYSTEM" in inp_id or "SUBSYS" in inp_id):
            mapeamento["SUBSYSTEM"] = e.Id
        elif mapeamento["TO_DAT"] is None and ("TO_DAT" in inp_id or "TODAT" in inp_id):
            mapeamento["TO_DAT"] = e.Id

    return mapeamento


def ler_alv_grid(grid) -> list[dict]:
    rows_count = int(grid.RowCount)
    cols_count = int(grid.ColumnCount)
    col_agr = None
    col_todat = None
    
    for c in range(cols_count):
        try:
            col_key = str(grid.GetColumnKey(c)).strip().upper()
            if col_key == "AGR_NAME":
                col_agr = col_key
            elif col_key == "TO_DAT":
                col_todat = col_key
        except Exception:
            pass
            
    if not col_agr or not col_todat:
        # Fallback para títulos das colunas
        for c in range(cols_count):
            try:
                title = str(grid.GetColumnTitle(c)).strip().upper()
                if "FUNÇÃO" in title or "ROLE" in title or "AGR_NAME" in title or "NOME DO" in title:
                    col_agr = grid.GetColumnKey(c)
                elif "VÁLIDO" in title or "VALID" in title or "TO_DAT" in title or "TO_DATE" in title:
                    col_todat = grid.GetColumnKey(c)
            except Exception:
                pass
                
    if not col_agr:
        col_agr = "AGR_NAME"
    if not col_todat:
        col_todat = "TO_DAT"
        
    results = []
    for r in range(rows_count):
        try:
            val_agr = str(grid.GetCellValue(r, col_agr)).strip()
            val_todat = str(grid.GetCellValue(r, col_todat)).strip()
            if val_agr or val_todat:
                results.append({
                    "AGR_NAME": val_agr,
                    "TO_DAT": val_todat
                })
        except Exception:
            pass
    return results


def ler_table_control(table_ctrl) -> list[dict]:
    cols = table_ctrl.Columns
    col_agr_idx = -1
    col_todat_idx = -1
    
    for idx in range(cols.Count):
        col_name = str(cols.ElementAt(idx).Name).upper()
        col_title = str(cols.ElementAt(idx).Title).upper()
        if "AGR_NAME" in col_name or "FUNÇÃO" in col_title or "ROLE" in col_title or "AGR_NAME" in col_title:
            col_agr_idx = idx
        elif "TO_DAT" in col_name or "VÁLIDO" in col_title or "VALID" in col_title or "TO_DAT" in col_title:
            col_todat_idx = idx
            
    if col_agr_idx == -1 or col_todat_idx == -1:
        col_agr_idx = 3
        col_todat_idx = 5
        
    results = []
    row_count = int(table_ctrl.RowCount)
    visible_row_count = int(table_ctrl.VisibleRowCount)
    
    if visible_row_count <= 0:
        return results
        
    scrollbar = None
    try:
        scrollbar = table_ctrl.verticalScrollbar
    except Exception:
        pass
        
    for r in range(row_count):
        row_in_screen = r
        if scrollbar is not None and r >= visible_row_count:
            try:
                scrollbar.position = r
                row_in_screen = 0
            except Exception:
                pass
                
        try:
            val_agr = str(table_ctrl.getCell(row_in_screen, col_agr_idx).Text).strip()
            val_todat = str(table_ctrl.getCell(row_in_screen, col_todat_idx).Text).strip()
            if val_agr or val_todat:
                results.append({
                    "AGR_NAME": val_agr,
                    "TO_DAT": val_todat
                })
        except Exception:
            pass
            
    return results


def ler_lista_standard(session) -> list[dict]:
    usr = session.findById("wnd[0]/usr")
    labels = []
    for child in usr.Children:
        if child.Type == "GuiLabel":
            labels.append(child)
            
    rows_data = {}
    for lbl in labels:
        lbl_id = lbl.Id
        if "[" in lbl_id and "]" in lbl_id:
            bracket_part = lbl_id.rsplit("[", 1)[-1].split("]")[0]
            parts = bracket_part.split(",")
            if len(parts) == 2:
                r_idx = int(parts[0])
                c_idx = int(parts[1])
                if r_idx not in rows_data:
                    rows_data[r_idx] = {}
                rows_data[r_idx][c_idx] = str(lbl.Text).strip()
                
    header_row_idx = -1
    col_agr_c = -1
    col_todat_c = -1
    
    for r_idx, row_cols in rows_data.items():
        for c_idx, val in row_cols.items():
            val_up = val.upper()
            if "AGR_NAME" in val_up or "FUNÇÃO" in val_up or "ROLE" in val_up:
                col_agr_c = c_idx
                header_row_idx = r_idx
            elif "TO_DAT" in val_up or "VÁLIDO" in val_up or "VALID" in val_up:
                col_todat_c = c_idx
                header_row_idx = r_idx
                
    results = []
    if header_row_idx == -1 or col_agr_c == -1 or col_todat_c == -1:
        return results
        
    for r_idx, row_cols in rows_data.items():
        if r_idx <= header_row_idx:
            continue
        val_agr = row_cols.get(col_agr_c, "")
        val_todat = row_cols.get(col_todat_c, "")
        if val_agr or val_todat:
            results.append({
                "AGR_NAME": val_agr,
                "TO_DAT": val_todat
            })
            
    return results


def aguardar_sap_livre(session, timeout=90, intervalo=0.25):
    """
    Acompanha a propriedade Busy da sessão e aguarda que ela fique livre (Busy=False).
    Utiliza timeout baseado em time.monotonic().
    Lança RuntimeError se o timeout for atingido.
    """
    t0 = time.monotonic()
    while True:
        try:
            busy = getattr(session, "Busy", False)
        except Exception:
            busy = False
            
        if not busy:
            return True
            
        if time.monotonic() - t0 > timeout:
            raise RuntimeError(f"A sessão SAP permaneceu ocupada (Busy=True) por mais de {timeout} segundos.")
            
        time.sleep(intervalo)


def abrir_se16_usla04(session) -> dict:
    """
    Executa a navegação na SE16 para a tabela USLA04 seguindo a ordem estrita:
    1. Abre a transação /NSE16.
    2. Valida se a transação SE16 foi aberta.
    3. Aguarda o campo de tabela estar disponível.
    4. Preenche USLA04.
    5. Confirma a tabela (Enter).
    6. Aguarda o ecrã de seleção carregar.
    7. Localiza e retorna o mapeamento dos campos BNAME, SUBSYSTEM e TO_DAT.
    
    Lança RuntimeError com mensagem específica se alguma etapa falhar.
    """
    print("[SE16] A abrir transação SE16...")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16"
        session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        raise RuntimeError(f"Não foi possível abrir a transação SE16. Detalhes: {e}")

    # Aguardar sessão livre e tratar popups
    aguardar_sap_livre(session, timeout=90)
    _dismiss_popup(session)

    # Validar transação atual e obter diagnósticos do ecrã inicial
    tcode_init = ""
    program_init = ""
    screen_init = ""
    titulo_init = ""
    try:
        tcode_init = str(getattr(session.Info, "Transaction", "")).strip().upper()
    except Exception:
        pass
    try:
        program_init = str(getattr(session.Info, "ProgramName", "")).strip()
    except Exception:
        pass
    try:
        screen_init = str(getattr(session.Info, "ScreenNumber", ""))
    except Exception:
        pass
    try:
        titulo_init = str(getattr(session.findById("wnd[0]"), "text", "")).strip()
    except Exception:
        pass

    diagnostico_init = f"Transação: '{tcode_init}' | Título: '{titulo_init}' | Programa: '{program_init}' | Dynpro: '{screen_init}'"
    
    if tcode_init != "SE16" and "SE16" not in titulo_init.upper():
        print(f"[SE16] ERRO de navegação - {diagnostico_init}")
        raise RuntimeError(f"Não foi possível abrir a transação SE16 no sistema CUA. Diagnóstico: {diagnostico_init}")

    print(f"[SE16] Transação atual confirmada: SE16 | {diagnostico_init}")

    # Aguardar até que o campo da tabela da SE16 esteja disponível
    tab_field = None
    t0 = time.monotonic()
    while time.monotonic() - t0 <= 10:
        for cid in ["wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "wnd[0]/usr/ctxtTABNAME", "wnd[0]/usr/ctxtGD-TAB"]:
            try:
                session.findById(cid)
                tab_field = cid
                break
            except Exception:
                pass
        if tab_field:
            break
        time.sleep(0.2)

    if not tab_field:
        raise RuntimeError(f"Campo para informar a tabela não encontrado na SE16. Diagnóstico: {diagnostico_init}")

    print("[SE16] Campo da tabela localizado.")

    # Preencher e submeter a tabela USLA04
    try:
        session.findById(tab_field).text = "USLA04"
        print("[SE16] Tabela USLA04 informada. A aguardar processamento...")
        session.findById("wnd[0]").sendVKey(0) # Enviar Enter
    except Exception as e:
        raise RuntimeError(f"A tabela USLA04 não foi aceite. Detalhes: {e}")

    # Aguardar que a sessão termine de processar o Enter
    aguardar_sap_livre(session, timeout=90)
    _dismiss_popup(session)

    # Verificar imediatamente se surgiu popup wnd[1] ou erro na barra de status
    if existe_elemento(session, "wnd[1]"):
        popup_msg = ""
        try:
            popup_msg = str(session.findById("wnd[1]/usr/lbl[0,1]").Text).strip()
        except Exception:
            try:
                popup_msg = str(session.findById("wnd[1]").text).strip()
            except Exception:
                popup_msg = "Popup detectado"
        raise RuntimeError(f"Popup de erro detectado na SE16: '{popup_msg}'")

    sbar_text = ""
    sbar_type = ""
    try:
        sbar = session.findById("wnd[0]/sbar")
        sbar_text = str(sbar.Text).strip()
        sbar_type = str(sbar.MessageType).strip().upper()
    except Exception:
        pass
    if sbar_type == "E" or (sbar_text and sbar_type not in ("S", "I")):
        raise RuntimeError(f"Erro na barra de status do SAP: '{sbar_text}'")

    # Aguardar a transição para o ecrã de seleção da USLA04
    t_start = time.monotonic()
    transition_success = False

    while time.monotonic() - t_start <= 90:
        # Verificar popups de erro recorrentes
        if existe_elemento(session, "wnd[1]"):
            popup_msg = ""
            try:
                popup_msg = str(session.findById("wnd[1]").text).strip()
            except Exception:
                popup_msg = "Popup detectado"
            raise RuntimeError(f"Popup de erro detectado na SE16: '{popup_msg}'")

        # Verificar erro na sbar recorrente
        try:
            sbar = session.findById("wnd[0]/sbar")
            sbar_text = str(sbar.Text).strip()
            sbar_type = str(sbar.MessageType).strip().upper()
        except Exception:
            pass
        if sbar_type == "E" or (sbar_text and sbar_type not in ("S", "I")):
            raise RuntimeError(f"Erro na barra de status do SAP: '{sbar_text}'")

        # Ler estado atual
        curr_title = ""
        curr_screen = ""
        curr_program = ""
        try:
            curr_title = str(getattr(session.findById("wnd[0]"), "text", "")).strip()
        except Exception:
            pass
        try:
            curr_screen = str(getattr(session.Info, "ScreenNumber", ""))
        except Exception:
            pass
        try:
            curr_program = str(getattr(session.Info, "ProgramName", ""))
        except Exception:
            pass

        # Verificar se o campo inicial da tabela deixou de existir/estar visível
        initial_field_exists = False
        try:
            session.findById(tab_field)
            initial_field_exists = True
        except Exception:
            pass

        # Verificar se o botão Executar (btn[8]) está disponível na tela
        exec_btn_exists = False
        for btn_id in ["wnd[0]/tbar[1]/btn[8]", "wnd[0]/tbar[0]/btn[8]"]:
            try:
                session.findById(btn_id)
                exec_btn_exists = True
                break
            except Exception:
                pass

        # Condição de Transição
        transition_by_title = (curr_title != "" and curr_title.upper() != titulo_init.upper()) or "USLA04" in curr_title.upper()
        transition_by_screen = (curr_screen != "" and curr_screen != screen_init) or (curr_program != "" and curr_program != program_init)
        transition_by_elements = (not initial_field_exists) or exec_btn_exists

        if transition_by_title or transition_by_screen or transition_by_elements:
            transition_success = True
            break

        time.sleep(0.25)
        aguardar_sap_livre(session, timeout=90)

    if not transition_success:
        # Diagnóstico em caso de timeout
        curr_tcode = ""
        curr_title = ""
        curr_program = ""
        curr_screen = ""
        curr_busy = False
        try:
            curr_tcode = str(getattr(session.Info, "Transaction", "")).strip()
        except Exception:
            pass
        try:
            curr_title = str(getattr(session.findById("wnd[0]"), "text", "")).strip()
        except Exception:
            pass
        try:
            curr_program = str(getattr(session.Info, "ProgramName", "")).strip()
        except Exception:
            pass
        try:
            curr_screen = str(getattr(session.Info, "ScreenNumber", "")).strip()
        except Exception:
            pass
        try:
            curr_busy = getattr(session, "Busy", False)
        except Exception:
            pass
            
        initial_field_exists = False
        try:
            session.findById(tab_field)
            initial_field_exists = True
        except Exception:
            pass
            
        exec_btn_exists = False
        for btn_id in ["wnd[0]/tbar[1]/btn[8]", "wnd[0]/tbar[0]/btn[8]"]:
            try:
                session.findById(btn_id)
                exec_btn_exists = True
                break
            except Exception:
                pass

        ctrls_dump = []
        try:
            usr = session.findById("wnd[0]/usr")
            for child in usr.Children:
                ctrls_dump.append(f"ID={child.Id}, Tipo={child.Type}, Text='{getattr(child, 'Text', '')}'")
        except Exception:
            pass

        print(f"[SE16][TIMEOUT]")
        print(f"Transação: {curr_tcode}")
        print(f"Título: {curr_title}")
        print(f"Programa: {curr_program}")
        print(f"Dynpro: {curr_screen}")
        print(f"Busy: {curr_busy}")
        print(f"Statusbar: {sbar_text}")
        print(f"Campo inicial da tabela existe: {initial_field_exists}")
        print(f"Botão Executar existe: {exec_btn_exists}")
        print(f"Controlos de input encontrados: {', '.join(ctrls_dump[:10])}")

        raise RuntimeError("O ecrã de seleção da USLA04 não foi carregado. Timeout de carregamento.")

    print("[SE16] Sessão SAP livre.")
    print("[SE16] Transição para o ecrã de seleção confirmada.")
    print("[SE16] Tabela USLA04 confirmada.")

    # Diagnóstico dos controlos reais encontrados no ecrã de seleção da SE16
    print("[SE16] A identificar campos dinâmicos...")
    try:
        usr = session.findById("wnd[0]/usr")
        print("[SE16] Diagnosticando controlos do ecrã de seleção:")
        for idx, child in enumerate(usr.Children):
            c_id = getattr(child, "Id", "")
            c_type = getattr(child, "Type", "")
            c_text = getattr(child, "Text", "")
            c_name = getattr(child, "Name", "")
            c_tooltip = getattr(child, "Tooltip", "")
            print(f"│  - Controlo [{idx}]: ID={c_id} | Tipo={c_type} | Name={c_name} | Text='{c_text}' | Tooltip='{c_tooltip}'")
    except Exception as diag_exc:
        print(f"[SE16] Erro ao registrar diagnóstico de controlos: {diag_exc}")

    # Descobrir campos dinâmicos
    mapeamento = descobrir_campos_se16(session)
    id_bname = mapeamento["BNAME"]
    id_subsys = mapeamento["SUBSYSTEM"]
    id_todat = mapeamento["TO_DAT"]

    if not id_bname or not id_subsys:
        raise RuntimeError("Campos de seleção da USLA04 não encontrados.")

    print(f"[SE16] BNAME localizado: {id_bname}")
    print(f"[SE16] SUBSYSTEM localizado: {id_subsys}")
    if id_todat:
        print(f"[SE16] TO_DAT localizado: {id_todat}")
    else:
        print("[SE16] TO_DAT não localizado (ignorado se não visível).")

    return mapeamento


def consultar_usla04_para_grupo(session, utilizador, sistema) -> list[dict]:
    """
    Abre a SE16 e consulta a tabela USLA04 para o utilizador e sistema indicados.
    Retorna lista de dicionários com AGR_NAME e TO_DAT para registos com TO_DAT >= hoje.

    Fluxo obrigatório:
      1. Abrir SE16 + USLA04 via abrir_se16_usla04()
      2. Definir max_hits
      3. Preencher BNAME e SUBSYSTEM com read-back obrigatório
      4. Configurar TO_DAT >= hoje usando o botão de seleção múltipla (press())
      5. Confirmar todos os valores antes de executar
      6. Executar (F8) e ler resultados
    """
    hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    hoje_str = formatar_data_para_sap(hoje)

    # ── 1. Abrir SE16 + USLA04 ──────────────────────────────────────────────
    campos = abrir_se16_usla04(session)
    id_bname    = campos["BNAME"]
    id_subsystem = campos["SUBSYSTEM"]
    id_todat    = campos.get("TO_DAT")

    if not id_bname or not id_subsystem:
        raise RuntimeError(
            "Campos de seleção da USLA04 não encontrados (BNAME/SUBSYSTEM ausentes). "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )

    # ── 2. Definir Max Ocorrências ────────────────────────────────────────────
    for mcid in ("wnd[0]/usr/txtMAX_SEL", "wnd[0]/usr/txtGD-MAXROWS", "wnd[0]/usr/txtMAX_HITS"):
        try:
            session.findById(mcid).text = "9999"
            break
        except Exception:
            pass

    # ── 3. Preencher BNAME com read-back ─────────────────────────────────────
    campo_bname = session.findById(id_bname)
    campo_bname.text = utilizador
    try:
        campo_bname.caretPosition = len(utilizador)
    except Exception:
        pass

    bname_lido = ""
    try:
        bname_lido = str(session.findById(id_bname).text).strip()
    except Exception:
        pass

    if bname_lido.upper() != utilizador.upper():
        # Segunda tentativa: atribuir via .Text
        try:
            session.findById(id_bname).Text = utilizador
            bname_lido = str(session.findById(id_bname).Text).strip()
        except Exception:
            pass

    if bname_lido.upper() != utilizador.upper():
        raise RuntimeError(
            f"BNAME não conservou o valor após atribuição "
            f"(esperado='{utilizador}', lido='{bname_lido}'). "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )
    print(f"[SE16] BNAME preenchido e confirmado: {bname_lido}")

    # ── 4. Preencher SUBSYSTEM com read-back ─────────────────────────────────
    campo_subsys = session.findById(id_subsystem)
    campo_subsys.text = sistema
    try:
        campo_subsys.caretPosition = len(sistema)
    except Exception:
        pass

    subsys_lido = ""
    try:
        subsys_lido = str(session.findById(id_subsystem).text).strip()
    except Exception:
        pass

    if subsys_lido.upper() != sistema.upper():
        try:
            session.findById(id_subsystem).Text = sistema
            subsys_lido = str(session.findById(id_subsystem).Text).strip()
        except Exception:
            pass

    if subsys_lido.upper() != sistema.upper():
        raise RuntimeError(
            f"SUBSYSTEM não conservou o valor após atribuição "
            f"(esperado='{sistema}', lido='{subsys_lido}'). "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )
    print(f"[SE16] SUBSYSTEM preenchido e confirmado: {subsys_lido}")

    # ── 5. Configurar TO_DAT >= hoje via botão de seleção múltipla ───────────
    ge_confirmado = False

    if id_todat:
        # Derivar ID do botão de seleção múltipla a partir do ID do campo:
        # ctxtIn-LOW  →  btn%_In_%_APP_%-VALU_PUSH
        # ou usar padrão direto descoberto no diagnóstico
        import re as _re_todat

        btn_valu_id = None

        # Estratégia A: derivar do ID do campo (In-LOW → %_In_%_APP_%-VALU_PUSH)
        mat = _re_todat.search(r'ctxt(I\d+)-LOW', id_todat)
        if mat:
            n = mat.group(1)
            # Construir o prefixo base do ID (pode ter /app/con[0]/ses[0]/ à frente)
            base = id_todat.rsplit(f"ctxt{n}-LOW", 1)[0]
            btn_valu_id = f"{base}btn%_{n}_%_APP_%-VALU_PUSH"

        # Estratégia B: procurar o botão de seleção múltipla nos filhos de wnd[0]/usr
        if not btn_valu_id or not existe_elemento(session, btn_valu_id):
            try:
                usr = session.findById("wnd[0]/usr")
                for child in usr.Children:
                    child_id  = str(getattr(child, "Id",   ""))
                    child_typ = str(getattr(child, "Type", ""))
                    child_tip = str(getattr(child, "Tooltip", "")).upper()
                    child_nm  = str(getattr(child, "Name", "")).upper()
                    if child_typ == "GuiButton" and (
                        "VALU_PUSH" in child_id.upper()
                        or "MULTIPLE" in child_tip
                        or "SELECTION" in child_tip
                        or "SELEÇÃO" in child_tip
                        or "VALU_PUSH" in child_nm
                    ):
                        # Verificar se este botão está associado ao campo TO_DAT
                        # (tem o mesmo índice n que o campo TO_DAT)
                        if mat:
                            if f"_%_{mat.group(1)}_%_" in child_id.upper() or mat.group(1).upper() in child_nm:
                                btn_valu_id = child_id
                                break
                        else:
                            btn_valu_id = child_id
                            break
            except Exception:
                pass

        if btn_valu_id and existe_elemento(session, btn_valu_id):
            # ── Abrir popup de seleção múltipla via press() ──────────────────
            try:
                session.findById(btn_valu_id).press()
                time.sleep(0.6)
                aguardar_sap_livre(session, timeout=15)
            except Exception as e:
                raise RuntimeError(
                    f"Não foi possível abrir o popup de seleção múltipla de TO_DAT "
                    f"via press() (botão: {btn_valu_id}). Detalhes: {e}. "
                    "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
                )

            # ── Preencher SIGN=I, OPTION=GE, LOW=hoje no popup ───────────────
            popup_ok = False
            try:
                # Campos standard da caixa de seleção múltipla do SAP:
                # wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4:0100/...
                # Tentamos os IDs mais comuns:
                sign_ids = [
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4:0100/ctxtGS_SELOPT-SIGN",
                    "wnd[1]/usr/ctxtGS_SELOPT-SIGN",
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4H:0101/ctxtGS_SELOPT-SIGN",
                ]
                option_ids = [
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4:0100/ctxtGS_SELOPT-OPTION",
                    "wnd[1]/usr/ctxtGS_SELOPT-OPTION",
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4H:0101/ctxtGS_SELOPT-OPTION",
                ]
                low_ids = [
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4:0100/ctxtGS_SELOPT-LOW",
                    "wnd[1]/usr/ctxtGS_SELOPT-LOW",
                    "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLSDH4H:0101/ctxtGS_SELOPT-LOW",
                ]
                copy_ids = [
                    "wnd[1]/tbar[0]/btn[8]",   # Copy/Executar no popup
                    "wnd[1]/tbar[0]/btn[0]",   # OK
                ]

                sign_elem   = None
                option_elem = None
                low_elem    = None

                for sid in sign_ids:
                    if existe_elemento(session, sid):
                        sign_elem = session.findById(sid)
                        break
                for oid in option_ids:
                    if existe_elemento(session, oid):
                        option_elem = session.findById(oid)
                        break
                for lid in low_ids:
                    if existe_elemento(session, lid):
                        low_elem = session.findById(lid)
                        break

                if sign_elem and option_elem and low_elem:
                    sign_elem.text   = "I"
                    option_elem.text = "GE"
                    low_elem.text    = hoje_str
                    time.sleep(0.2)

                    # Confirmar popup (Copy)
                    copied = False
                    for cid in copy_ids:
                        if existe_elemento(session, cid):
                            session.findById(cid).press()
                            copied = True
                            break
                    if not copied:
                        # Tentar Enter
                        session.findById("wnd[1]").sendVKey(0)

                    time.sleep(0.4)
                    aguardar_sap_livre(session, timeout=15)
                    popup_ok = True
                else:
                    # Popup aberto mas campos SIGN/OPTION/LOW não encontrados
                    # Fechar popup e tentar fallback direto
                    _dismiss_popup(session)
                    time.sleep(0.3)

            except Exception as popup_exc:
                # Garantir que o popup fecha antes de propagar
                _dismiss_popup(session)
                time.sleep(0.3)
                raise RuntimeError(
                    f"Erro ao configurar condição GE no popup de seleção múltipla de TO_DAT. "
                    f"Detalhes: {popup_exc}. "
                    "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
                )

            if popup_ok:
                ge_confirmado = True
                print(f"[SE16] TO_DAT configurado: GE {hoje_str}")
            else:
                # Popup não tinha os campos esperados — tentar fallback direto no campo
                try:
                    session.findById(id_todat).text = hoje_str
                    time.sleep(0.2)
                    todat_lido = str(session.findById(id_todat).text).strip()
                    if todat_lido:
                        # Sem garantia de GE — tratar como falha crítica
                        raise RuntimeError(
                            "Não foi possível configurar TO_DAT >= hoje na SE16 "
                            "(popup de seleção múltipla não tinha campos SIGN/OPTION/LOW). "
                            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
                        )
                except RuntimeError:
                    raise
                except Exception:
                    pass
                raise RuntimeError(
                    "Não foi possível configurar TO_DAT >= hoje na SE16. "
                    "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
                )

        else:
            # Botão de seleção múltipla não encontrado — fallback: escrever direto no campo
            # Sem garantia de GE — bloquear por segurança
            raise RuntimeError(
                f"Botão de seleção múltipla de TO_DAT não encontrado (campo: {id_todat}). "
                "Não é possível garantir TO_DAT >= hoje na SE16. "
                "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
            )

    else:
        # TO_DAT não localizado no ecrã — bloquear por segurança
        raise RuntimeError(
            "Campo TO_DAT não localizado no ecrã de seleção da SE16. "
            "Não é possível garantir TO_DAT >= hoje. "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )

    # ── 6. Validação final obrigatória antes de executar ─────────────────────
    if not ge_confirmado:
        raise RuntimeError(
            "TO_DAT com condição GE não foi confirmado. "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )

    # Re-verificar BNAME e SUBSYSTEM (podem ter mudado se o popup causou refresh)
    try:
        bname_final = str(session.findById(id_bname).text).strip()
    except Exception:
        bname_final = bname_lido

    try:
        subsys_final = str(session.findById(id_subsystem).text).strip()
    except Exception:
        subsys_final = subsys_lido

    print(f"[SE16] Valores antes da execução:")
    print(f"  BNAME={bname_final}")
    print(f"  SUBSYSTEM={subsys_final}")
    print(f"  TO_DAT_OPTION=GE")
    print(f"  TO_DAT_LOW={hoje_str}")

    if bname_final.upper() != utilizador.upper():
        raise RuntimeError(
            f"BNAME foi apagado após configuração do popup de TO_DAT "
            f"(esperado='{utilizador}', lido='{bname_final}'). "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )
    if subsys_final.upper() != sistema.upper():
        raise RuntimeError(
            f"SUBSYSTEM foi apagado após configuração do popup de TO_DAT "
            f"(esperado='{sistema}', lido='{subsys_final}'). "
            "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
        )

    print("[SE16] Filtros preenchidos e confirmados. A executar consulta...")

    # ── 7. Executar (F8) ─────────────────────────────────────────────────────
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(1.5)
    aguardar_sap_livre(session, timeout=60)
    _dismiss_popup(session)
    print("[SE16] Consulta executada. A ler resultados...")

    # ── 8. Ler resultados de forma polimórfica ────────────────────────────────
    # ── 8a. Verificar sbar por mensagem de "nenhum registo" antes de procurar grid
    sbar_pos_exec = ""
    sbar_tipo_pos_exec = ""
    try:
        sbar_obj = session.findById("wnd[0]/sbar")
        sbar_pos_exec = str(getattr(sbar_obj, "Text", "")).strip()
        sbar_tipo_pos_exec = str(getattr(sbar_obj, "MessageType", "")).strip().upper()
    except Exception:
        pass

    sbar_upper = sbar_pos_exec.upper()
    if any(term in sbar_upper for term in [
        "NENHUM", "NO DATA", "NOT FOUND", "NAO EXISTE", "NÃO EXISTE",
        "NENHUMA ENTRADA", "NO ENTRIES", "0 ENTRIES", "NO RECORDS",
        "KEIN EINTRAG", "AUCUN ENREGISTREMENT"
    ]):
        print(f"[SE16] Resultado: tipo=sem_registos | sbar='{sbar_pos_exec}'")
        return []

    if sbar_tipo_pos_exec == "E" and sbar_pos_exec:
        raise RuntimeError(
            f"Erro de autorização ou execução na SE16 após consulta: '{sbar_pos_exec}'. "
            "A leitura do resultado não foi comprovada — não é tratado como nenhum registo."
        )

    # ── 8b. ALV Grid ──────────────────────────────────────────────────────────
    grid = None
    for c in (
        "wnd[0]/usr/cntlRESULT/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
    ):
        try:
            obj = session.findById(c)
        except Exception:
            continue
        try:
            _ = obj.RowCount
            grid = obj
            break
        except Exception as e:
            raise RuntimeError(
                f"Nao foi possivel obter RowCount do ALV — erro de leitura "
                f"(não é tratado como nenhum registo): {e}"
            )

    if grid is not None:
        resultados = ler_alv_grid(grid)
        print(f"[SE16] Resultado: tipo=ALV_grid | linhas={len(resultados)}")
        return resultados

    # ── 8c. GuiTableControl ───────────────────────────────────────────────────
    table_ctrl = None
    try:
        usr = session.findById("wnd[0]/usr")
        for child in usr.Children:
            if getattr(child, "Type", "") == "GuiTableControl":
                table_ctrl = child
                break
    except Exception:
        pass

    if table_ctrl is not None:
        resultados = ler_table_control(table_ctrl)
        print(f"[SE16] Resultado: tipo=table_control | linhas={len(resultados)}")
        return resultados

    # ── 8d. Standard List ─────────────────────────────────────────────────────
    try:
        resultados_labels = ler_lista_standard(session)
        if resultados_labels:
            print(f"[SE16] Resultado: tipo=lista_standard | linhas={len(resultados_labels)}")
            return resultados_labels
    except Exception:
        pass

    # ── 8e. Nenhuma estrutura de resultado encontrada ─────────────────────────
    # Verificar novamente sbar (pode ter chegado após polling)
    try:
        sbar_obj2 = session.findById("wnd[0]/sbar")
        sbar2 = str(getattr(sbar_obj2, "Text", "")).strip().upper()
        if any(term in sbar2 for term in [
            "NENHUM", "NO DATA", "NOT FOUND", "NAO EXISTE", "NÃO EXISTE",
            "NENHUMA ENTRADA", "NO ENTRIES", "0 ENTRIES", "NO RECORDS"
        ]):
            print(f"[SE16] Resultado: tipo=sem_registos (confirmado por sbar) | sbar='{sbar2}'")
            return []
    except Exception:
        pass

    # Não encontrou grid nem mensagem de vazio — resultado inconclusivo
    # Não classificar como "nenhum registo" para não permitir inserção indevida
    print("[SE16] Resultado: tipo=inconclusivo (sem grid, sem sbar definitiva)")
    raise RuntimeError(
        "Resultado da consulta USLA04 inconclusivo: não foi possível localizar "
        "o grid de resultados nem confirmar ausência de registos. "
        "A pré-validação da USLA04 foi cancelada e nenhuma alteração foi efetuada."
    )


def prevalidar_e_processar_atribuicoes(
    df_filtrado,
    session,
    sistema_desejado,
    pedir_confirmacao=True,
    modo_nao_interativo=False
) -> pd.DataFrame:
    """
    Executa a pré-validação das atribuições na USLA04 via SE16 no SAP CUA,
    filtando duplicações do Excel e existentes em CUA antes de aceder à SU10.
    """
    if df_filtrado is None or df_filtrado.empty:
        return df_filtrado

    # Cache de execução da sessão
    cache_execucao = set()

    # 1. Validar duplicações no próprio Excel (UTILIZADOR + SISTEMA + AGR_NAME)
    vistos = {} # (user_norm, sys_norm, role_norm) -> idx_original
    duplicados = [] # list of (idx_duplicado, idx_original)
    linhas_validar_indices = []
    
    for idx_row in df_filtrado.index:
        user = str(df_filtrado.at[idx_row, "UTILIZADOR"]).strip().upper()
        sistema = str(df_filtrado.at[idx_row, "SISTEMA"]).strip().upper()
        role = str(df_filtrado.at[idx_row, "AGR_NAME"]).strip().upper()
        
        if not user or not sistema or not role:
            linhas_validar_indices.append(idx_row)
            continue
            
        chave = (user, sistema, role)
        if chave in vistos:
            duplicados.append((idx_row, vistos[chave]))
        else:
            vistos[chave] = idx_row
            linhas_validar_indices.append(idx_row)

    # DataFrame apenas com as linhas únicas a validar/processar
    df_unicos = df_filtrado.loc[linhas_validar_indices]

    # 2. Agrupar os registos pendentes por UTILIZADOR e SISTEMA
    grupos_validar = {}
    for idx_row in df_unicos.index:
        user = str(df_unicos.at[idx_row, "UTILIZADOR"]).strip().upper()
        sys_name = str(df_unicos.at[idx_row, "SISTEMA"]).strip().upper()
        role = str(df_unicos.at[idx_row, "AGR_NAME"]).strip().upper()
        if not user or not sys_name or not role:
            continue
        chave_grupo = (user, sys_name)
        if chave_grupo not in grupos_validar:
            grupos_validar[chave_grupo] = []
        grupos_validar[chave_grupo].append((idx_row, role))

    # 3. Executar pré-validação na tabela USLA04
    existentes_por_grupo = {} # (user_norm, sys_norm) -> {role_norm: to_dat}
    erros_validacao = {} # idx_row -> erro_msg
    hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    print("\n[Etapa 2] Pré-validação CUA na tabela USLA04...")
    
    for (user_norm, sys_norm), info_linhas in grupos_validar.items():
        try:
            print(f"├─ A consultar USLA04 para {user_norm} no sistema {sys_norm}...")
            linhas_retornadas = consultar_usla04_para_grupo(session, user_norm, sys_norm)
            
            existentes_por_grupo[(user_norm, sys_norm)] = {}
            for reg in linhas_retornadas:
                r_name = str(reg.get("AGR_NAME", "")).strip().upper()
                to_dat_str = str(reg.get("TO_DAT", "")).strip()
                existentes_por_grupo[(user_norm, sys_norm)][r_name] = to_dat_str
                
        except Exception as e:
            # Caso a consulta falhe por completo
            erro_msg = f"Não foi possível validar previamente a atribuição na tabela USLA04. Nenhuma alteração foi efetuada no SAP. Detalhes: {e}"
            print(f"│  ❌ Erro de validação: {erro_msg}")
            for idx_row, role in info_linhas:
                erros_validacao[idx_row] = erro_msg

    # Voltar para o início no SAP após SE16
    voltar_para_inicio(session)

    # 4. Separar linhas com base no resultado da USLA04
    linhas_inexistentes = [] # list of idx_row
    status_originais = {} # idx_original -> (status, msg)

    contadores = {
        "total_pendentes": len(df_filtrado),
        "ja_atribuidas": 0,
        "expiradas": 0,
        "inexistentes": 0,
        "duplicadas": len(duplicados),
        "erros_validacao": 0
    }

    # Atualizar em memória o STATUS, MSG e TIMESTEMP das linhas únicas resolvidas na pré-validação
    for idx_row in df_unicos.index:
        if idx_row in erros_validacao:
            msg_err = erros_validacao[idx_row]
            status_originais[idx_row] = ("ERRO", msg_err)
            marcar_resultado(df_filtrado, idx_row, "ERRO", msg_err)
            contadores["erros_validacao"] += 1
            continue
            
        user = str(df_filtrado.at[idx_row, "UTILIZADOR"]).strip().upper()
        sistema = str(df_filtrado.at[idx_row, "SISTEMA"]).strip().upper()
        role = str(df_filtrado.at[idx_row, "AGR_NAME"]).strip().upper()
        
        if not user or not sistema or not role:
            # Caso os campos obrigatórios estejam em falta
            msg_err = "Dados obrigatórios (UTILIZADOR/SISTEMA/AGR_NAME) vazios."
            status_originais[idx_row] = ("ERRO", msg_err)
            marcar_resultado(df_filtrado, idx_row, "ERRO", msg_err)
            contadores["erros_validacao"] += 1
            continue

        chave_grupo = (user, sistema)
        grupo_dados = existentes_por_grupo.get(chave_grupo, {})
        
        if role in grupo_dados:
            to_dat_str = grupo_dados[role]
            try:
                dt_val = converter_data_sap(to_dat_str)
                if dt_val >= hoje:
                    # Atribuição existente e ativa
                    msg_ja_existe = f"Função '{role}' já atribuída ao utilizador '{user}' no sistema '{sistema}', com validade até {to_dat_str}. Nenhuma alteração efetuada."
                    status_originais[idx_row] = ("CONCLUIDO", msg_ja_existe)
                    marcar_resultado(df_filtrado, idx_row, "CONCLUIDO", msg_ja_existe)
                    cache_execucao.add((user, sistema, role))
                    contadores["ja_atribuidas"] += 1
                else:
                    # Expirada - apta para nova atribuição
                    contadores["expiradas"] += 1
                    linhas_inexistentes.append(idx_row)
            except Exception as dt_exc:
                # Na dúvida, assume ativa
                msg_ja_existe = f"Função '{role}' já atribuída ao utilizador '{user}' no sistema '{sistema}', com validade {to_dat_str} (erro de conversão: {dt_exc}). Nenhuma alteração efetuada."
                status_originais[idx_row] = ("CONCLUIDO", msg_ja_existe)
                marcar_resultado(df_filtrado, idx_row, "CONCLUIDO", msg_ja_existe)
                cache_execucao.add((user, sistema, role))
                contadores["ja_atribuidas"] += 1
        else:
            # Inexistente
            contadores["inexistentes"] += 1
            linhas_inexistentes.append(idx_row)

    # 5. Processar duplicados das linhas já pré-validadas
    for idx_dup, idx_orig in duplicados:
        if idx_orig in status_originais:
            status_orig, msg_orig = status_originais[idx_orig]
            if status_orig == "CONCLUIDO":
                id_orig = df_filtrado.at[idx_orig, "ID"]
                msg_dup = f"Combinação duplicada no ficheiro. A função foi tratada pela linha ID '{id_orig}'."
                marcar_resultado(df_filtrado, idx_dup, "CONCLUIDO", msg_dup)
            elif status_orig == "ERRO":
                id_orig = df_filtrado.at[idx_orig, "ID"]
                msg_dup = f"Não foi possível validar previamente a atribuição na tabela USLA04. Nenhuma alteração foi efetuada no SAP. Dependia da linha ID '{id_orig}'."
                marcar_resultado(df_filtrado, idx_dup, "ERRO", msg_dup)

    # 6. Apresentar o resumo da pré-validação
    print("\n======================================================================")
    print("Pré-validação USLA04 concluída.")
    print(f"Total de linhas pendentes: {contadores['total_pendentes']}")
    print(f"Já atribuídas e válidas: {contadores['ja_atribuidas']}")
    print(f"Expiradas e aptas para nova atribuição: {contadores['expiradas']}")
    print(f"Inexistentes e aptas para inserção: {contadores['inexistentes']}")
    print(f"Duplicadas no Excel: {contadores['duplicadas']}")
    print(f"Erros de validação: {contadores['erros_validacao']}")
    print(f"\nFunções que serão inseridas no SAP: {len(linhas_inexistentes)}")
    print("======================================================================")

    if not linhas_inexistentes:
        print("\n[INFO] Não existem novas atribuições a realizar.")
        return df_filtrado

    # 7. Pedir confirmação apenas para as funções inexistentes
    if not modo_nao_interativo and pedir_confirmacao:
        resposta = input("Deseja lançar as funções inexistentes no SAP? [S/N]: ").strip().upper()
        if resposta != "S":
            print("❌ Lançamento cancelado pelo utilizador.")
            return df_filtrado

    # 8. Executar a SU10 apenas para as funções realmente inexistentes
    df_candidatos = df_filtrado.loc[linhas_inexistentes].copy()
    
    # Executa com modo_nao_interativo=True e pedir_confirmacao=False para a SU10 rodar automaticamente
    df_candidatos_proc = atribuir_funcao_usuario(
        df_candidatos,
        session,
        sistema_desejado,
        pedir_confirmacao=False,
        modo_nao_interativo=True
    )

    # Atualizar o DataFrame principal com os resultados da SU10
    for idx_row in linhas_inexistentes:
        df_filtrado.at[idx_row, "STATUS"] = df_candidatos_proc.at[idx_row, "STATUS"]
        df_filtrado.at[idx_row, "MSG"] = df_candidatos_proc.at[idx_row, "MSG"]
        df_filtrado.at[idx_row, "TIMESTEMP"] = df_candidatos_proc.at[idx_row, "TIMESTEMP"]

    # 9. Resolver os duplicados pendentes que dependiam da inserção da SU10
    for idx_dup, idx_orig in duplicados:
        if idx_orig in linhas_inexistentes:
            status_orig = df_filtrado.at[idx_orig, "STATUS"]
            msg_orig = df_filtrado.at[idx_orig, "MSG"]
            id_orig = df_filtrado.at[idx_orig, "ID"]
            
            # Só atualizamos se o duplicado ainda não foi resolvido
            if df_filtrado.at[idx_dup, "STATUS"] != "CONCLUIDO" and df_filtrado.at[idx_dup, "STATUS"] != "ERRO":
                if status_orig == "CONCLUIDO" or normalizar_valor(status_orig) == "CONCLUIDO":
                    msg_dup = f"Combinação duplicada no ficheiro. A função foi tratada pela linha ID '{id_orig}'."
                    marcar_resultado(df_filtrado, idx_dup, "CONCLUIDO", msg_dup)
                else:
                    msg_dup = f"Linha não processada devido a falha na validação/inserção da primeira ocorrência (ID '{id_orig}'). Detalhes: {msg_orig}"
                    marcar_resultado(df_filtrado, idx_dup, "ERRO", msg_dup)

    return df_filtrado


def obter_funcoes_existentes(shell) -> set[tuple[str, str]]:
    """
    Lê o grid de funções existentes para o utilizador no SAP CUA.
    Retorna um conjunto de tuplos (SUBSYSTEM_NORMALIZADO, AGR_NAME_NORMALIZADO).
    """
    existentes = set()
    
    # 1. Obter RowCount de forma robusta
    try:
        row_count = getattr(shell, "RowCount", None)
        if row_count is None:
            row_count = getattr(shell, "rowCount", None)
        if row_count is None:
            raise RuntimeError("Não foi possível obter a propriedade RowCount/rowCount do shell.")
    except Exception as e:
        print(f"[AVISO] Erro ao aceder ao RowCount: {e}")
        raise RuntimeError(f"Erro ao ler RowCount do shell de funções: {e}")

    # 2. Determinar qual o método de leitura disponível (GetCellValue vs getCellValue)
    metodo_leitura = None
    if hasattr(shell, "GetCellValue"):
        metodo_leitura = shell.GetCellValue
    elif hasattr(shell, "getCellValue"):
        metodo_leitura = shell.getCellValue
    
    if metodo_leitura is None and row_count > 0:
        try:
            shell.GetCellValue(0, "SUBSYSTEM")
            metodo_leitura = shell.GetCellValue
        except Exception:
            try:
                shell.getCellValue(0, "SUBSYSTEM")
                metodo_leitura = shell.getCellValue
            except Exception as e:
                raise RuntimeError(
                    f"Método GetCellValue/getCellValue indisponível ou inacessível no shell: {e}"
                )

    # 3. Ler cada linha
    for i in range(row_count):
        subsystem = ""
        agr_name = ""
        
        # Tenta ler SUBSYSTEM
        try:
            if metodo_leitura is not None:
                subsystem = metodo_leitura(i, "SUBSYSTEM")
            else:
                try:
                    subsystem = shell.GetCellValue(i, "SUBSYSTEM")
                except Exception:
                    subsystem = shell.getCellValue(i, "SUBSYSTEM")
        except Exception as e:
            print(f"[AVISO] Erro ao ler SUBSYSTEM na linha {i}: {e}")
            subsystem = ""

        # Tenta ler AGR_NAME
        try:
            if metodo_leitura is not None:
                agr_name = metodo_leitura(i, "AGR_NAME")
            else:
                try:
                    agr_name = shell.GetCellValue(i, "AGR_NAME")
                except Exception:
                    agr_name = shell.getCellValue(i, "AGR_NAME")
        except Exception as e:
            print(f"[AVISO] Erro ao ler AGR_NAME na linha {i}: {e}")
            agr_name = ""

        sub_str = str(subsystem or "").strip().upper()
        agr_str = str(agr_name or "").strip().upper()
        
        if sub_str or agr_str:
            existentes.add((sub_str, agr_str))
            
    return existentes


###################################################################################
# BLOCO 7: FILTRO DE LINHAS A PROCESSAR
###################################################################################

def filtrar_pendentes(df):
    if df is None or df.empty:
        return pd.DataFrame()

    df2 = df.copy()
    df2["STATUS_NORM"] = df2["STATUS"].apply(normalizar_valor)

    pend = df2[
        (df2["CHAVE_ID"] != "") &
        (df2["STATUS_NORM"] != "CONCLUIDO")
    ].drop(columns=["STATUS_NORM"])

    if pend.empty:
        print("\n⚠️ Nenhuma linha com STATUS ≠ 'Concluído' foi encontrada.")
    else:
        print("\n📋 Linhas a processar:")
        exibir = pend[["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME"]].copy()
        for c in exibir.columns:
            exibir[c] = exibir[c].apply(texto_limpo)
        print(exibir.to_string(index=False))
        print()

    return pend


###################################################################################
# BLOCO 8: EXECUÇÃO SAP
###################################################################################

def marcar_resultado(df_ref, idx, status, msg):
    df_ref.at[idx, "STATUS"] = texto_limpo(status)
    df_ref.at[idx, "MSG"] = texto_limpo(msg)
    df_ref.at[idx, "TIMESTEMP"] = agora_str()


def atribuir_funcao_usuario(df_filtrado, session, sistema_desejado, pedir_confirmacao=True, modo_nao_interativo=False):
    """
    Atribui AGR_NAME ao UTILIZADOR via SU10 de forma agrupada por utilizador/sistema.
    Garante idempotência, evitando atribuir funções que já existam no SAP CUA,
    e previne duplicações no próprio ficheiro Excel.
    """
    if df_filtrado is None or df_filtrado.empty:
        return df_filtrado

    # 0. Identificar e marcar duplicados no próprio Excel (UTILIZADOR + SISTEMA + AGR_NAME)
    vistos = set()
    duplicados_marcar = []
    indices_unicos = []
    
    for idx_row in df_filtrado.index:
        user_val = str(df_filtrado.at[idx_row, "UTILIZADOR"]).strip()
        sys_val = str(df_filtrado.at[idx_row, "SISTEMA"]).strip()
        role_val = str(df_filtrado.at[idx_row, "AGR_NAME"]).strip()
        
        user_norm = user_val.upper()
        sys_norm = sys_val.upper()
        role_norm = role_val.upper()
        
        if not user_norm or not sys_norm or not role_norm:
            indices_unicos.append(idx_row)
            continue
            
        chave = (user_norm, sys_norm, role_norm)
        if chave in vistos:
            duplicados_marcar.append(idx_row)
        else:
            vistos.add(chave)
            indices_unicos.append(idx_row)
            
    # Marcar os duplicados imediatamente como CONCLUIDO
    for idx_row in duplicados_marcar:
        user = df_filtrado.at[idx_row, "UTILIZADOR"]
        sys_name = df_filtrado.at[idx_row, "SISTEMA"]
        role = df_filtrado.at[idx_row, "AGR_NAME"]
        msg_dup = f"Combinação '{role}' no sistema '{sys_name}' para o utilizador '{user}' duplicada no Excel. Tratada na primeira ocorrência."
        marcar_resultado(df_filtrado, idx_row, "CONCLUIDO", msg_dup)

    # DataFrame com as linhas únicas a processar
    df_unicos = df_filtrado.loc[indices_unicos]

    # 1. Agrupar em memória por UTILIZADOR e SISTEMA
    grupos = df_unicos.groupby(["UTILIZADOR", "SISTEMA"], sort=False)
    total_grupos = len(grupos)
    
    total_linhas_pendentes = len(df_unicos)
    total_roles_distintas = df_unicos["AGR_NAME"].nunique()
    
    print(f"\n[INFO] Utilizadores a processar agrupados (excluindo duplicados do Excel): {total_grupos}")
    print(f"[INFO] Linhas únicas pendentes: {total_linhas_pendentes}")
    print(f"[INFO] Roles distintas: {total_roles_distintas}")

    if not modo_nao_interativo and pedir_confirmacao:
        resposta = input("Deseja lançar essas funções no SAP? [S/N]: ").strip().upper()
        if resposta != "S":
            print("[X] Lançamento cancelado pelo utilizador.")
            return df_filtrado

    tempo_total_inicio = time.time()

    for idx_grupo, ((utilizador, sistema), df_grupo) in enumerate(grupos, 1):
        inicio = time.time()
        eventos_status = []
        
        # Obter a lista de roles únicas a adicionar para este utilizador e sistema
        roles_list = list(dict.fromkeys([str(r).strip() for r in df_grupo["AGR_NAME"] if str(r).strip()]))
        
        print("\n======================================================================")
        print(f">>> [{idx_grupo}/{total_grupos}] INICIANDO UTILIZADOR: {utilizador} | Sistema: {sistema} | Roles: {len(roles_list)}")
        print("======================================================================")

        # Verificar dados vazios
        if not utilizador or not sistema or not roles_list:
            msg = "Dados obrigatórios (UTILIZADOR/SISTEMA/ROLES) vazios."
            for idx_row in df_grupo.index:
                marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
            duracao_str = formatar_tempo(time.time() - inicio)
            print(f"🔴 ERRO: {msg} ⏱️ (Tempo: {duracao_str})")
            continue

        try:
            sistema_conectado = texto_limpo(session.Info.SystemName).upper()
            if sistema_conectado != sistema_desejado:
                msg = f"Sistema SAP incorreto: esperado {sistema_desejado}, conectado a {sistema_conectado}"
                for idx_row in df_grupo.index:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
                duracao_str = formatar_tempo(time.time() - inicio)
                print(f"🔴 ERRO: {msg} ⏱️ (Tempo: {duracao_str})")
                continue

            # 1) Abre SU10
            print("\n[Etapa 1] Pesquisa de Utilizador")
            print("├─ Abrindo SU10...")
            ir_para_transacao(session, "SU10")
            capturar_status_bar(session, eventos_status, origem="ABERTURA_SU10", tentativas=5, espera=0.20)

            grid_input = "wnd[0]/usr/tblSAPLSUID_MAINTENANCETC_USERS"
            campo_utilizador = grid_input + "/ctxtSUID_ST_BNAME-BNAME[0,0]"
            btn_selecionar = "wnd[0]/tbar[1]/btn[18]"
            tab_funcoes = "wnd[0]/usr/tabsTABSTRIP1/tabpACTG"
            shell_funcoes = (
                "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/"
                "ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/"
                "cntlG_ROLES_CONTAINER/shellcont/shell"
            )

            if not esperar_elemento(session, campo_utilizador, tentativas=20, espera=0.5):
                msg = "Falha ao abrir SU10."
                for idx_row in df_grupo.index:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
                duracao_str = formatar_tempo(time.time() - inicio)
                print(f"ERRO: {msg} (Tempo: {duracao_str})")
                continue

            # 2) Preenche utilizador e seleciona
            print(f"├─ Inserindo utilizador: {utilizador}")
            campo = session.findById(campo_utilizador)
            campo.text = ""
            campo.text = utilizador
            campo.caretPosition = len(utilizador)

            print("└─ Selecionando utilizador...")
            session.findById(btn_selecionar).press()
            time.sleep(0.60)
            tipo_sel, _, msg_sel = capturar_status_bar(
                session,
                eventos_status,
                origem="SELECAO_UTILIZADOR",
                tentativas=6,
                espera=0.20,
            )

            if normalizar_valor(tipo_sel) in ("E", "A", "X"):
                msg_final = montar_msg_final(eventos_status) or msg_sel
                for idx_row in df_grupo.index:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg_final)
                duracao_str = formatar_tempo(time.time() - inicio)
                print(f"ERRO: {msg_final} (Tempo: {duracao_str})")
                continue

            # 3) Vai para tab de funções
            print("\n[Etapa 2] Atribuição de Funções no SAP CUA")
            print("├─ Acedendo à aba de funções...")
            session.findById(tab_funcoes).select()
            time.sleep(0.40)
            capturar_status_bar(session, eventos_status, origem="ABERTURA_TAB_FUNCOES", tentativas=4, espera=0.20)

            shell = esperar_elemento(session, shell_funcoes, tentativas=20, espera=0.5)
            if not shell:
                msg = "Não foi possível abrir a aba de funções no SU10."
                for idx_row in df_grupo.index:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
                duracao_str = formatar_tempo(time.time() - inicio)
                print(f"ERRO: {msg} (Tempo: {duracao_str})")
                continue

            # 3b) Obter funções existentes no SAP CUA
            try:
                funcoes_existentes = obter_funcoes_existentes(shell)
                print(f"├─ Funções já atribuídas detetadas no grid do utilizador ({len(funcoes_existentes)}):")
                for sub_e, agr_e in funcoes_existentes:
                    print(f"│  - {sub_e} / {agr_e}")
            except Exception as read_exc:
                msg = f"Erro ao ler as funções existentes do utilizador no SAP: {read_exc}"
                print(f"ERRO: {msg}")
                for idx_row in df_grupo.index:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
                continue

            # Identificar quais das roles pedidas já estão no grid
            roles_a_adicionar = []
            for role_name in roles_list:
                role_norm = str(role_name).strip().upper()
                sistema_norm = str(sistema).strip().upper()
                
                if (sistema_norm, role_norm) in funcoes_existentes:
                    indices_da_role = df_grupo[df_grupo["AGR_NAME"] == role_name].index
                    for idx_row in indices_da_role:
                        msg_ja_existe = f"Função '{role_name}' já atribuída ao utilizador '{utilizador}' no sistema '{sistema}'. Nenhuma alteração efetuada."
                        marcar_resultado(df_filtrado, idx_row, "CONCLUIDO", msg_ja_existe)
                    print(f"│  [INFO] Função '{role_name}' já atribuída ao utilizador no sistema '{sistema}'. Ignorando.")
                else:
                    roles_a_adicionar.append(role_name)

            if not roles_a_adicionar:
                print("└─ Nenhuma nova função para adicionar (todas já atribuídas). Gravação ignorada.")
                continue

            # 4) Preenche subsystem e AGR_NAME para cada uma das roles
            print(f"├─ Preparando inserção de {len(roles_a_adicionar)} role(s)...")
            role_errors = {}
            row_idx = 0

            for r_idx, role_name in enumerate(roles_a_adicionar):
                print(f"├─ Inserindo role {r_idx+1}/{len(roles_a_adicionar)}: {role_name}")
                
                # Procurar a primeira linha vazia a partir de row_idx
                while row_idx < shell.rowCount:
                    subsys = None
                    try:
                        if hasattr(shell, "GetCellValue"):
                            subsys = shell.GetCellValue(row_idx, "SUBSYSTEM")
                        elif hasattr(shell, "getCellValue"):
                            subsys = shell.getCellValue(row_idx, "SUBSYSTEM")
                    except Exception:
                        pass
                    
                    agr = None
                    try:
                        if hasattr(shell, "GetCellValue"):
                            agr = shell.GetCellValue(row_idx, "AGR_NAME")
                        elif hasattr(shell, "getCellValue"):
                            agr = shell.getCellValue(row_idx, "AGR_NAME")
                    except Exception:
                        pass
                        
                    subsys_str = str(subsys or "").strip()
                    agr_str = str(agr or "").strip()
                    if not subsys_str and not agr_str:
                        break
                    row_idx += 1
                
                try:
                    if row_idx >= 5:
                        shell.firstVisibleRow = row_idx - 4
                        
                    shell.modifyCell(row_idx, "SUBSYSTEM", sistema)
                    shell.modifyCell(row_idx, "AGR_NAME", role_name)
                    shell.currentCellColumn = "AGR_NAME"
                    shell.pressEnter()
                    time.sleep(0.5)
                    
                    local_events = []
                    tipo_pre, _, msg_pre = capturar_status_bar(
                        session,
                        local_events,
                        origem=f"VAL_ROLE_{role_name}",
                        tentativas=5,
                        espera=0.15,
                    )
                    
                    if normalizar_valor(tipo_pre) in ("E", "A", "X"):
                        err_msg = montar_msg_final(local_events) or msg_pre
                        role_errors[role_name] = err_msg
                        print(f"│  [AVISO] Falha na validação da role '{role_name}': {err_msg}")
                        
                        # Limpar a linha problemática
                        shell.modifyCell(row_idx, "SUBSYSTEM", "")
                        shell.modifyCell(row_idx, "AGR_NAME", "")
                        shell.pressEnter()
                        time.sleep(0.3)
                    else:
                        role_errors[role_name] = None
                        row_idx += 1
                        
                except Exception as cell_exc:
                    err_msg = str(cell_exc)
                    role_errors[role_name] = err_msg
                    print(f"│  [AVISO] Erro técnico ao inserir role '{role_name}': {err_msg}")

            # 5) Save - se pelo menos uma role correu bem
            salvou_com_sucesso = False
            save_msg = "Nenhuma role com sucesso para gravar."
            sucesso_roles = [r for r, err in role_errors.items() if err is None]

            if sucesso_roles:
                print("└─ Guardando alterações...")
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                time.sleep(0.40)
                
                save_events = []
                capturar_status_bar(
                    session,
                    save_events,
                    origem="SAVE_IMEDIATO",
                    tentativas=8,
                    espera=0.20,
                )
                
                tratar_popups_pos_save(session, save_events, max_popups=5)
                
                capturar_status_bar(
                    session,
                    save_events,
                    origem="SAVE_FINAL",
                    tentativas=10,
                    espera=0.25,
                )
                
                status_final = decidir_status_pelo_historico(save_events)
                save_msg = montar_msg_final(save_events)
                
                if status_final == "CONCLUÍDO" or normalizar_valor(status_final) == "CONCLUIDO":
                    salvou_com_sucesso = True
                else:
                    for r in sucesso_roles:
                        role_errors[r] = f"Falha na gravação final: {save_msg}"
            else:
                print("└─ Gravação ignorada (todas as roles falharam).")

            # 6) Atribuir resultados linha a linha no df original
            total_ok = 0
            for idx_row in df_grupo.index:
                if df_filtrado.at[idx_row, "STATUS"] == "CONCLUIDO":
                    continue
                    
                row_role = str(df_filtrado.at[idx_row, "AGR_NAME"]).strip()
                err = role_errors.get(row_role)
                
                if err:
                    marcar_resultado(df_filtrado, idx_row, "ERRO", err)
                else:
                    if salvou_com_sucesso:
                        msg_sucesso = f"{save_msg or 'Atribuído com sucesso'} | Role atribuída no processamento agrupado"
                        marcar_resultado(df_filtrado, idx_row, "CONCLUIDO", msg_sucesso)
                        total_ok += 1
                    else:
                        marcar_resultado(df_filtrado, idx_row, "ERRO", f"Não gravado: {save_msg}")

            # 7) Log de resultado do utilizador
            print("\n[Etapa 3] Resultado do Utilizador")
            duracao = time.time() - inicio
            duracao_str = formatar_tempo(duracao)
            
            roles_ok = sum(1 for err in role_errors.values() if err is None)
            if salvou_com_sucesso and roles_ok == len(roles_a_adicionar):
                print(f"SUCESSO: Utilizador tratado por completo! Roles: {roles_ok}/{len(roles_a_adicionar)} (Tempo: {duracao_str})")
            else:
                print(f"ERRO: Atribuição parcial ou falha na gravação. Roles: {total_ok}/{len(roles_a_adicionar)} com sucesso. (Tempo: {duracao_str})")

        except Exception as e:
            msg = str(e)
            for idx_row in df_grupo.index:
                if df_filtrado.at[idx_row, "STATUS"] != "CONCLUIDO":
                    marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
            duracao_str = formatar_tempo(time.time() - inicio)
            print(f"ERRO: {msg} (Tempo: {duracao_str})")

        finally:
            voltar_para_inicio(session)

    tempo_total = time.time() - tempo_total_inicio
    print(f"\nTempo total: {formatar_tempo(tempo_total)}")

    status_norm = df_filtrado["STATUS"].apply(normalizar_valor)
    total_ok = (status_norm == "CONCLUIDO").sum()
    total_erro = (status_norm == "ERRO").sum()
    print(f"Total concluído: {total_ok} | Com erro: {total_erro}")

    return df_filtrado


###################################################################################
# BLOCO 9: GRAVAÇÃO NO EXCEL SEM PERDER FORMATAÇÃO
###################################################################################

def gravar_preservando_formatacao(caminho_ficheiro, nome_sheet, df_atualizado):
    """
    Atualiza APENAS:
    - STATUS
    - MSG
    - TIMESTEMP

    Faz match pela coluna ID.
    """
    try:
        ext = os.path.splitext(caminho_ficheiro)[1].lower()
        keep_vba = ext == ".xlsm"

        wb = load_workbook(caminho_ficheiro, keep_vba=keep_vba)
        if nome_sheet not in wb.sheetnames:
            print(f"❌ Sheet '{nome_sheet}' não existe para gravar.")
            return False

        ws = wb[nome_sheet]
        mapa_cols = mapear_cabecalhos_openpyxl(ws)

        if "TIMESTAMP" in mapa_cols and "TIMESTEMP" not in mapa_cols:
            mapa_cols["TIMESTEMP"] = mapa_cols["TIMESTAMP"]

        obrig_excel = ["ID", "STATUS", "MSG", "TIMESTEMP"]
        falta = [c for c in obrig_excel if c not in mapa_cols]
        if falta:
            print(f"❌ Cabeçalhos obrigatórios em falta na sheet para gravação: {', '.join(falta)}")
            return False

        col_id = mapa_cols["ID"]
        col_status = mapa_cols["STATUS"]
        col_msg = mapa_cols["MSG"]
        col_timestemp = mapa_cols["TIMESTEMP"]

        mapa_linhas_por_id = {}
        for r in range(2, ws.max_row + 1):
            valor_id = ws.cell(row=r, column=col_id).value
            id_chave = chave_id(valor_id)
            if id_chave:
                mapa_linhas_por_id[id_chave] = r

        atualizados = 0
        nao_encontrados = 0

        for _, row in df_atualizado.iterrows():
            id_chave = texto_limpo(row.get("CHAVE_ID", ""))
            if not id_chave:
                continue

            linha_excel = mapa_linhas_por_id.get(id_chave)
            if not linha_excel:
                nao_encontrados += 1
                print(f"⚠️ ID não encontrado na sheet para gravação: {id_chave}")
                continue

            ws.cell(row=linha_excel, column=col_status).value = texto_limpo(row.get("STATUS", ""))
            ws.cell(row=linha_excel, column=col_msg).value = texto_limpo(row.get("MSG", ""))
            ws.cell(row=linha_excel, column=col_timestemp).value = texto_limpo(row.get("TIMESTEMP", ""))

            atualizados += 1

        wb.save(caminho_ficheiro)

        print(
            f"💾 Ficheiro atualizado com formatação preservada "
            f"(sheet '{nome_sheet}') | Linhas atualizadas: {atualizados}"
        )

        if nao_encontrados:
            print(f"⚠️ IDs não encontrados na sheet: {nao_encontrados}")

        return True

    except PermissionError:
        base, ext = os.path.splitext(caminho_ficheiro)
        alternativo = f"{base}_resultado{ext}"
        try:
            wb.save(alternativo)
            print(f"⚠️ Ficheiro estava aberto. Foi criada uma cópia:\n   {alternativo}")
            return True
        except Exception as e:
            print(f"❌ Erro ao salvar cópia: {e}")
            return False

    except Exception as e:
        print(f"❌ Erro ao salvar: {e}")
        return False


###################################################################################
# BLOCO 10: API PARA O COCKPIT
###################################################################################

def executar(
    ambiente_cockpit,
    pfcg_object=None,
    caminho_ficheiro=None,
    request_transporte=None,
    modo_nao_interativo=False,
    pedir_confirmacao=True,
    **kwargs
):
    tempo_inicio_total = time.time()
    print(f"✅ Processo selecionado: {NOME_SCRIPT}")

    sheet_alvo = pfcg_object if pfcg_object else NOME_SHEET
    print(f"📄 Script atual: {NOME_SCRIPT} | Sheet alvo: '{sheet_alvo}'")

    if modo_nao_interativo:
        if not caminho_ficheiro:
            raise ValueError("Faltou o parâmetro caminho_ficheiro em modo web/não-interativo.")
    else:
        if not caminho_ficheiro:
            caminho_ficheiro = selecionar_ficheiro_excel()

    if not caminho_ficheiro:
        return False

    print("\n[Etapa 1] Leitura do Excel")
    df = ler_ficheiro(caminho_ficheiro, sheet_alvo)
    if df is None:
        return False

    sistema_desejado = MAPA_SISTEMA.get(ambiente_cockpit)
    if not sistema_desejado:
        print(f"❌ Ambiente inválido: {ambiente_cockpit}. Use: {', '.join(MAPA_SISTEMA.keys())}")
        return False

    session = conectar_sap(sistema_desejado)
    if not session:
        return False

    df_pend = filtrar_pendentes(df)
    if df_pend.empty:
        tempo_decorrido_total = time.time() - tempo_inicio_total
        print(f"\n⏱️ Tempo total da operação: {formatar_tempo(tempo_decorrido_total)}")
        print("🔁 Fim.")
        return True

    df_proc = prevalidar_e_processar_atribuicoes(
        df_pend.copy(),
        session,
        sistema_desejado,
        pedir_confirmacao=pedir_confirmacao,
        modo_nao_interativo=modo_nao_interativo
    )

    print("\n[Etapa 4] Gravação de Resultados")
    ok_save = gravar_preservando_formatacao(caminho_ficheiro, sheet_alvo, df_proc)
    if ok_save:
        print("💾 Resultados gravados com sucesso no Excel!")

    tempo_decorrido_total = time.time() - tempo_inicio_total
    print(f"\n⏱️ Tempo total da operação: {formatar_tempo(tempo_decorrido_total)}")
    print("🔁 Fim.")
    return ok_save


###################################################################################
# BLOCO 11: EXECUÇÃO DIRETA
###################################################################################

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD", "CUA"])
    parser.add_argument("--xlsx")
    parser.add_argument("--auto", action="store_true")
    parser.add_argument("--no-confirm", action="store_true")
    args = parser.parse_args()

    env_cli = args.ambiente or "CUA"
    executar(
        ambiente_cockpit=env_cli,
        caminho_ficheiro=args.xlsx,
        modo_nao_interativo=bool(args.auto),
        pedir_confirmacao=(not args.no_confirm)
    )