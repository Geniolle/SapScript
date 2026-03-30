# -*- coding: utf-8 -*-

###################################################################################
# PROCESSO: Remover Funções SU01 (sheet = nome do ficheiro .py sem prefixo)
###################################################################################
# Ex.: "I. CUA_REMOVE.py"  →  Sheet "CUA_REMOVE"
#
# COLUNAS ESPERADAS NA SHEET:
# ID | UTILIZADOR | SISTEMA | AGR_NAME | STATUS | MSG | TIMESTEMP
#
# MODO TEMPORÁRIO DE VALIDAÇÃO:
# - Só avança quando o utilizador carregar Enter
# - Mostra no terminal cada passo antes de executar
# - Faz varredura real dos campos do popup de filtro
###################################################################################

###################################################################################
# BLOCO 1: IMPORTAÇÕES E CONFIGURAÇÃO
###################################################################################

import os
import time
import pandas as pd
import unicodedata
import win32com.client
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from datetime import datetime

###################################################################################
# BLOCO 2: DETETAR NOME DO SCRIPT E MAPA DE SISTEMAS
###################################################################################

try:
    NOME_SCRIPT = os.path.splitext(os.path.basename(__file__))[0]
except NameError:
    NOME_SCRIPT = "I. CUA_REMOVE"  # fallback

NOME_SHEET = NOME_SCRIPT.split(".", 1)[-1].strip() if "." in NOME_SCRIPT else NOME_SCRIPT

MAPA_SISTEMA = {
    "DEV": "S4D",
    "QAD": "S4Q",
    "PRD": "S4P",
    "CUA": "SPA"
}

# TEMPORÁRIO: modo passo a passo
MODO_DEBUG_PASSO_A_PASSO = True

###################################################################################
# BLOCO 3: UTILITÁRIOS
###################################################################################

def normalizar_coluna(valor):
    return unicodedata.normalize("NFKD", str(valor)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

def agora_timestamp():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def texto_limpo(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def pausar(msg):
    if MODO_DEBUG_PASSO_A_PASSO:
        input(f"\n⏸️ {msg}\nPrima Enter para continuar...")

def selecionar_ficheiro_excel():
    """
    Abre o popup SEM diretoria predefinida.
    O Windows abrirá naturalmente na última pasta utilizada.
    """
    try:
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        root.attributes("-topmost", True)

        caminho = filedialog.askopenfilename(
            title="Selecione o ficheiro Excel",
            filetypes=[("Ficheiros Excel", "*.xlsx *.xls"), ("Todos os ficheiros", "*.*")]
        )

        root.destroy()

        if not caminho:
            print("⚠️ Seleção de ficheiro cancelada pelo utilizador.")
            return None

        print(f"✅ Ficheiro selecionado: {caminho}")
        return caminho

    except Exception as e:
        print(f"❌ Erro ao abrir a janela de seleção: {e}")
        return None

def obter_status_bar(session):
    try:
        sbar = session.findById("wnd[0]/sbar")
        tipo = str(getattr(sbar, "MessageType", "") or "").strip().upper()
        texto = str(getattr(sbar, "Text", "") or "").strip()
        return tipo, texto
    except Exception:
        return "", ""

def tipo_sbar_para_status(tipo_sbar):
    tipo = (tipo_sbar or "").strip().upper()

    if tipo == "S":
        return "CONCLUÍDO"
    if tipo == "W":
        return "AVISO"
    if tipo in ("E", "A"):
        return "ERRO"
    if tipo == "I":
        return "INFO"
    return "ERRO"

def registar_resultado(df, idx, status, msg):
    df.loc[idx, "STATUS"] = status
    df.loc[idx, "MSG"] = msg
    df.loc[idx, "TIMESTEMP"] = agora_timestamp()

def existe_objeto(session, obj_id):
    try:
        session.findById(obj_id)
        return True
    except Exception:
        return False

def obter_objeto(session, obj_id):
    try:
        return session.findById(obj_id)
    except Exception as e:
        raise Exception(f"Objeto SAP não encontrado: {obj_id} | {e}")

def setar_texto_debug(session, obj_id, valor, descricao="campo"):
    print(f"📝 A preencher {descricao}: {obj_id} = '{valor}'")
    pausar(f"Validar antes de preencher {descricao}")

    try:
        obj = obter_objeto(session, obj_id)
        try:
            obj.setFocus()
        except Exception:
            pass

        try:
            obj.text = ""
        except Exception:
            pass

        obj.text = str(valor)

        try:
            obj.caretPosition = len(str(valor))
        except Exception:
            pass

    except Exception as e:
        raise Exception(f"Falha ao preencher {descricao} ({obj_id}) com valor '{valor}': {e}")

def pressionar_botao_debug(session, obj_id, descricao="botão"):
    print(f"🖱️ A pressionar {descricao}: {obj_id}")
    pausar(f"Validar antes de pressionar {descricao}")

    try:
        obj = obter_objeto(session, obj_id)
        obj.press()
    except Exception as e:
        raise Exception(f"Falha ao pressionar {descricao} ({obj_id}): {e}")

def enviar_vkey_debug(session, wnd_id, vkey, descricao="ação"):
    print(f"⌨️ A executar {descricao}: {wnd_id}.sendVKey({vkey})")
    pausar(f"Validar antes de executar {descricao}")

    try:
        wnd = obter_objeto(session, wnd_id)
        wnd.sendVKey(vkey)
    except Exception as e:
        raise Exception(f"Falha ao executar {descricao} em {wnd_id} com VKey={vkey}: {e}")

def selecionar_tab_debug(session, obj_id, descricao="tab"):
    print(f"📑 A selecionar {descricao}: {obj_id}")
    pausar(f"Validar antes de selecionar {descricao}")

    try:
        obj = obter_objeto(session, obj_id)
        obj.select()
    except Exception as e:
        raise Exception(f"Falha ao selecionar {descricao} ({obj_id}): {e}")

def obter_children(obj):
    try:
        children = obj.Children
        total = children.Count
        return [children.Item(i) for i in range(total)]
    except Exception:
        try:
            return list(obj.Children)
        except Exception:
            return []

def coletar_componentes_recursivo(obj, lista):
    """
    Percorre recursivamente os componentes SAP a partir de um nó.
    """
    try:
        lista.append(obj)
    except Exception:
        return

    for filho in obter_children(obj):
        coletar_componentes_recursivo(filho, lista)

def listar_campos_popup(session):
    """
    Lista todos os componentes do wnd[1], para debug.
    """
    componentes = []
    try:
        wnd1 = session.findById("wnd[1]")
    except Exception:
        return componentes

    coletar_componentes_recursivo(wnd1, componentes)
    return componentes

def descrever_componente(obj):
    try:
        obj_id = str(getattr(obj, "Id", "") or "")
    except Exception:
        obj_id = ""

    try:
        obj_type = str(getattr(obj, "Type", "") or "")
    except Exception:
        obj_type = ""

    try:
        name = str(getattr(obj, "Name", "") or "")
    except Exception:
        name = ""

    try:
        text = str(getattr(obj, "Text", "") or "")
    except Exception:
        text = ""

    try:
        changeable = getattr(obj, "Changeable")
    except Exception:
        changeable = "?"

    return {
        "id": obj_id,
        "type": obj_type,
        "name": name,
        "text": text,
        "changeable": changeable
    }

def obter_campos_low_popup(session):
    """
    Procura no popup todos os campos candidatos que terminem com -LOW.
    Ordena por prioridade:
      1) changeable=True
      2) ids contendo 'ctxt'
      3) ids contendo 'txt'
    """
    candidatos = []

    for comp in listar_campos_popup(session):
        info = descrever_componente(comp)
        obj_id = info["id"].upper()

        if "-LOW" not in obj_id:
            continue

        candidatos.append(info)

    def score(info):
        changeable = info["changeable"] is True
        obj_id = info["id"].lower()
        prioridade_tipo = 0
        if "/ctxt" in obj_id:
            prioridade_tipo = 2
        elif "/txt" in obj_id:
            prioridade_tipo = 1
        return (1 if changeable else 0, prioridade_tipo, len(obj_id))

    candidatos.sort(key=score, reverse=True)
    return candidatos

def preencher_popup_filtro(session, valor):
    """
    Em vez de assumir um ID fixo, procura todos os campos LOW do popup
    e tenta escrever no primeiro que realmente aceite .text.
    """
    pausar("Validar popup de filtro antes de procurar os campos reais")

    candidatos = obter_campos_low_popup(session)

    if not candidatos:
        raise Exception("Nenhum campo '*-LOW' foi encontrado no wnd[1].")

    print("🔎 Campos candidatos encontrados no popup:")
    for idx, info in enumerate(candidatos, 1):
        print(
            f"   {idx}. id='{info['id']}' | type='{info['type']}' | "
            f"name='{info['name']}' | changeable='{info['changeable']}' | text='{info['text']}'"
        )

    erros = []

    for info in candidatos:
        obj_id = info["id"]

        print(f"🧪 Tentativa de preencher candidato do popup: {obj_id} = '{valor}'")
        pausar(f"Validar antes de testar o campo {obj_id}")

        try:
            obj = session.findById(obj_id)

            try:
                obj.setFocus()
            except Exception:
                pass

            try:
                atual = str(getattr(obj, "Text", "") or "")
            except Exception:
                atual = ""

            try:
                obj.text = ""
            except Exception:
                pass

            obj.text = str(valor)

            try:
                escrito = str(getattr(obj, "Text", "") or "")
            except Exception:
                escrito = str(valor)

            print(f"✅ Campo aceite: {obj_id} | antes='{atual}' | depois='{escrito}'")
            pressionar_botao_debug(session, "wnd[1]/tbar[0]/btn[0]", "confirmar filtro")
            return

        except Exception as e:
            erros.append(f"{obj_id} => {e}")
            print(f"⚠️ Campo rejeitou escrita: {obj_id} | erro: {e}")

    raise Exception(
        "Nenhum dos campos LOW do popup aceitou escrita. "
        + " | ".join(erros)
    )

def obter_grid_roles(session):
    return session.findById(
        "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/"
        "ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/"
        "cntlG_ROLES_CONTAINER/shellcont/shell"
    )

def obter_row_count_grid(shell):
    for attr in ("RowCount", "rowCount"):
        try:
            return int(getattr(shell, attr))
        except Exception:
            pass
    return 0

def obter_valor_celula_grid(shell, row, coluna):
    for metodo in ("GetCellValue", "getCellValue"):
        try:
            fn = getattr(shell, metodo)
            return str(fn(row, coluna)).strip()
        except Exception:
            pass
    return ""

###################################################################################
# BLOCO 4: LEITURA DO EXCEL
###################################################################################

def ler_ficheiro_excel(caminho_ficheiro, nome_sheet):
    if not caminho_ficheiro or not os.path.exists(caminho_ficheiro):
        print("❌ Ficheiro não encontrado ou caminho inválido.")
        return None

    try:
        wb = load_workbook(caminho_ficheiro, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()

        if nome_sheet not in sheets:
            print(f"❌ Sheet '{nome_sheet}' não encontrada. Disponíveis: {', '.join(sheets)}")
            return None

        df = pd.read_excel(caminho_ficheiro, sheet_name=nome_sheet, dtype=object)
        df.columns = [normalizar_coluna(c) for c in df.columns]

        df.rename(columns={
            "USER": "UTILIZADOR",
            "USERNAME": "UTILIZADOR",
            "SYSTEM": "SISTEMA",
            "NOME FUNCAO": "AGR_NAME",
            "FUNCAO": "AGR_NAME",
            "ROLE": "AGR_NAME",
            "TIMESTAMP": "TIMESTEMP"
        }, inplace=True)

        obrigatorias = ["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME"]
        faltantes = [c for c in obrigatorias if c not in df.columns]
        if faltantes:
            print(f"❌ Colunas obrigatórias em falta na sheet '{nome_sheet}': {', '.join(faltantes)}")
            return None

        for col in ["STATUS", "MSG", "TIMESTEMP"]:
            if col not in df.columns:
                df[col] = ""

        df["_LINHA_EXCEL"] = df.index + 2

        for col in ["ID", "UTILIZADOR", "SISTEMA", "AGR_NAME", "STATUS", "MSG", "TIMESTEMP"]:
            df[col] = df[col].fillna("").astype(str)

        df["STATUS"] = df["STATUS"].str.strip().str.upper()
        df["MSG"] = df["MSG"].str.strip()
        df["TIMESTEMP"] = df["TIMESTEMP"].str.strip()

        print(f"📄 Sheet carregada: '{nome_sheet}' | Registos: {len(df)}")
        return df

    except Exception as e:
        print(f"❌ Erro ao ler o ficheiro/sheet: {e}")
        return None

###################################################################################
# BLOCO 5: CONEXÃO COM SAP GUI
###################################################################################

def conectar_sap(sistema_desejado):
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        for conn in application.Children:
            for sess in conn.Children:
                try:
                    if sess.Info.SystemName.upper() == sistema_desejado.upper():
                        print(
                            f"✅ Conectado ao SAP: {sess.Info.SystemName.upper()} | "
                            f"Utilizador: {sess.Info.User} | Cliente: {sess.Info.Client}"
                        )
                        return sess
                except Exception:
                    continue

        print(f"❌ Sessão SAP não encontrada para o sistema '{sistema_desejado}'.")
        return None

    except Exception as e:
        print(f"❌ Erro na conexão SAP GUI: {e}")
        return None

###################################################################################
# BLOCO 6: EXECUÇÃO (REMOVER FUNÇÃO NO SU01)
###################################################################################

def remover_funcao_usuario(df, session):
    total = len(df)

    for i, (idx, row) in enumerate(df.iterrows(), 1):
        id_linha = texto_limpo(row.get("ID", ""))
        utilizador = texto_limpo(row.get("UTILIZADOR", ""))
        sistema = texto_limpo(row.get("SISTEMA", ""))
        agr_name = texto_limpo(row.get("AGR_NAME", ""))

        print(
            f"\n🔧 {i}/{total} | ID={id_linha} | "
            f"UTILIZADOR='{utilizador}' | SISTEMA='{sistema}' | AGR_NAME='{agr_name}'"
        )

        inicio = time.time()

        if not utilizador:
            msg = "UTILIZADOR vazio."
            print(f"❌ {msg}")
            registar_resultado(df, idx, "ERRO", msg)
            continue

        if not sistema:
            msg = "SISTEMA vazio."
            print(f"❌ {msg}")
            registar_resultado(df, idx, "ERRO", msg)
            continue

        if not agr_name:
            msg = "AGR_NAME vazio."
            print(f"❌ {msg}")
            registar_resultado(df, idx, "ERRO", msg)
            continue

        try:
            print("➡️ Passo 1: Entrar na SU01")
            setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/nSU01", "campo de comando")
            enviar_vkey_debug(session, "wnd[0]", 0, "confirmar entrada na SU01")
            time.sleep(0.3)

            print("➡️ Passo 2: Informar utilizador")
            setar_texto_debug(session, "wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME", utilizador, "campo utilizador")
            enviar_vkey_debug(session, "wnd[0]", 0, "confirmar utilizador")
            time.sleep(0.4)

            tipo_sbar, texto_sbar = obter_status_bar(session)
            print(f"📣 STATUS BAR após abrir utilizador: tipo='{tipo_sbar}' | texto='{texto_sbar}'")

            if tipo_sbar in ("E", "A"):
                msg = texto_sbar or f"Erro ao abrir o utilizador '{utilizador}'."
                print(f"❌ {msg}")
                registar_resultado(df, idx, tipo_sbar_para_status(tipo_sbar), msg)
                continue

            print("➡️ Passo 3: Entrar em modo alteração")
            pressionar_botao_debug(session, "wnd[0]/tbar[1]/btn[18]", "botão alterar")
            time.sleep(0.3)

            print("➡️ Passo 4: Selecionar tab de funções")
            selecionar_tab_debug(session, "wnd[0]/usr/tabsTABSTRIP1/tabpACTG", "tab ACTG")
            time.sleep(0.3)

            print("➡️ Passo 5: Obter grid de funções")
            pausar("Validar grid de funções antes de obter o objeto")
            shell = obter_grid_roles(session)

            print("➡️ Passo 6: Filtrar SUBSYSTEM")
            pausar("Validar antes de abrir filtro SUBSYSTEM")
            shell.currentCellColumn = "SUBSYSTEM"
            shell.contextMenu()
            shell.selectContextMenuItem("&FILTER")
            preencher_popup_filtro(session, sistema)
            time.sleep(0.3)

            print("➡️ Passo 7: Filtrar AGR_NAME")
            pausar("Validar antes de abrir filtro AGR_NAME")
            shell.currentCellColumn = "AGR_NAME"
            shell.contextMenu()
            shell.selectContextMenuItem("&FILTER")
            preencher_popup_filtro(session, agr_name)
            time.sleep(0.3)

            row_count = obter_row_count_grid(shell)
            print(f"📊 RowCount do grid após filtros: {row_count}")

            if row_count <= 0:
                msg = f"Nenhum registo encontrado para SISTEMA='{sistema}' e AGR_NAME='{agr_name}'."
                print(f"❌ {msg}")
                registar_resultado(df, idx, "ERRO", msg)
                setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/N", "campo de comando")
                enviar_vkey_debug(session, "wnd[0]", 0, "sair da transação")
                continue

            agr_encontrado = obter_valor_celula_grid(shell, 0, "AGR_NAME")
            sistema_encontrado = obter_valor_celula_grid(shell, 0, "SUBSYSTEM")

            print(f"🔍 Linha 0 do grid | AGR_NAME='{agr_encontrado}' | SUBSYSTEM='{sistema_encontrado}'")

            if agr_encontrado and normalizar_coluna(agr_encontrado) != normalizar_coluna(agr_name):
                msg = (
                    f"Registo encontrado no grid não corresponde ao AGR_NAME esperado. "
                    f"Esperado='{agr_name}' | Encontrado='{agr_encontrado}'"
                )
                print(f"❌ {msg}")
                registar_resultado(df, idx, "ERRO", msg)
                setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/N", "campo de comando")
                enviar_vkey_debug(session, "wnd[0]", 0, "sair da transação")
                continue

            if sistema_encontrado and normalizar_coluna(sistema_encontrado) != normalizar_coluna(sistema):
                msg = (
                    f"Registo encontrado no grid não corresponde ao SISTEMA esperado. "
                    f"Esperado='{sistema}' | Encontrado='{sistema_encontrado}'"
                )
                print(f"❌ {msg}")
                registar_resultado(df, idx, "ERRO", msg)
                setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/N", "campo de comando")
                enviar_vkey_debug(session, "wnd[0]", 0, "sair da transação")
                continue

            print("➡️ Passo 8: Selecionar linha 0 e remover")
            pausar("Validar antes de remover a linha do grid")
            shell.setCurrentCell(0, "AGR_NAME")
            shell.selectedRows = "0"
            shell.pressToolbarButton("DEL_LINE")
            time.sleep(0.3)

            print("➡️ Passo 9: Gravar")
            enviar_vkey_debug(session, "wnd[0]", 11, "gravar alteração")
            time.sleep(0.5)

            tipo_sbar, texto_sbar = obter_status_bar(session)
            duracao = time.time() - inicio

            msg_sap = texto_sbar or f"Processo concluído para AGR_NAME='{agr_name}'."
            msg_final = f"{msg_sap} | Tempo: {duracao:.1f}s"
            status_final = tipo_sbar_para_status(tipo_sbar if tipo_sbar else "S")

            print(f"📣 STATUS BAR final: tipo='{tipo_sbar}' | texto='{texto_sbar}'")
            print(f"{'✅' if status_final == 'CONCLUÍDO' else '⚠️' if status_final == 'AVISO' else '❌'} {msg_final}")
            registar_resultado(df, idx, status_final, msg_final)

            print("➡️ Passo 10: Sair da transação")
            setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/N", "campo de comando")
            enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saída da transação")

        except Exception as e:
            tipo_sbar, texto_sbar = obter_status_bar(session)
            detalhe_sap = f" | SAP: {texto_sbar}" if texto_sbar else ""
            erro = f"{str(e).strip()}{detalhe_sap}"
            print(f"❌ Erro ao remover AGR_NAME='{agr_name}' do utilizador '{utilizador}': {erro}")
            registar_resultado(df, idx, "ERRO", erro)

            try:
                print("↩️ Tentativa de sair da transação após erro")
                setar_texto_debug(session, "wnd[0]/tbar[0]/okcd", "/N", "campo de comando")
                enviar_vkey_debug(session, "wnd[0]", 0, "confirmar saída após erro")
            except Exception:
                pass

    return df

###################################################################################
# BLOCO 7: GUARDAR RESULTADOS PRESERVANDO FORMATAÇÃO
###################################################################################

def mapear_cabecalhos(ws):
    mapa = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=1, column=col).value
        if valor is None:
            continue
        mapa[normalizar_coluna(valor)] = col
    return mapa

def garantir_coluna_sheet(ws, mapa_cols, nome_coluna):
    chave = normalizar_coluna(nome_coluna)
    if chave in mapa_cols:
        return mapa_cols[chave]

    nova_col = ws.max_column + 1
    ws.cell(row=1, column=nova_col).value = nome_coluna
    mapa_cols[chave] = nova_col
    return nova_col

def salvar_resultado(df, caminho_ficheiro, nome_sheet):
    try:
        wb = load_workbook(caminho_ficheiro)

        if nome_sheet not in wb.sheetnames:
            print(f"❌ Sheet '{nome_sheet}' não existe para gravar.")
            return

        ws = wb[nome_sheet]
        mapa_cols = mapear_cabecalhos(ws)

        col_status = garantir_coluna_sheet(ws, mapa_cols, "STATUS")
        col_msg = garantir_coluna_sheet(ws, mapa_cols, "MSG")
        col_timestamp = garantir_coluna_sheet(ws, mapa_cols, "TIMESTEMP")

        total_atualizadas = 0

        for _, row in df.iterrows():
            linha_excel = row.get("_LINHA_EXCEL")
            if pd.isna(linha_excel):
                continue

            linha_excel = int(linha_excel)
            ws.cell(row=linha_excel, column=col_status).value = texto_limpo(row.get("STATUS", ""))
            ws.cell(row=linha_excel, column=col_msg).value = texto_limpo(row.get("MSG", ""))
            ws.cell(row=linha_excel, column=col_timestamp).value = texto_limpo(row.get("TIMESTEMP", ""))
            total_atualizadas += 1

        wb.save(caminho_ficheiro)
        print(
            f"💾 Resultados atualizados na sheet '{nome_sheet}' "
            f"(apenas STATUS / MSG / TIMESTEMP). Linhas atualizadas: {total_atualizadas}"
        )

    except PermissionError:
        base, ext = os.path.splitext(caminho_ficheiro)
        alternativo = f"{base}_resultado{ext}"
        wb.save(alternativo)
        print(f"⚠️ Ficheiro estava aberto. Foi criada uma cópia:\n   {alternativo}")

    except Exception as e:
        print(f"❌ Erro ao salvar preservando formatação: {e}")

###################################################################################
# BLOCO 8: EXECUTAR PROCESSO
###################################################################################

def executar(ambiente):
    print(f"📄 Script atual: {NOME_SCRIPT} | Sheet alvo: '{NOME_SHEET}'")

    caminho = selecionar_ficheiro_excel()
    if not caminho:
        return

    df = ler_ficheiro_excel(caminho, NOME_SHEET)
    if df is None:
        return

    sistema_desejado = MAPA_SISTEMA.get(str(ambiente).strip().upper())
    if not sistema_desejado:
        print(f"❌ Ambiente inválido: {ambiente}. Use: {', '.join(MAPA_SISTEMA.keys())}")
        return

    session = conectar_sap(sistema_desejado)
    if not session:
        return

    df_final = remover_funcao_usuario(df, session)
    salvar_resultado(df_final, caminho, NOME_SHEET)

###################################################################################
# EXEMPLO DE CHAMADA:
# executar("CUA")
###################################################################################