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

        wb = load_workbook(
            caminho_ficheiro,
            read_only=True,
            data_only=False,
            keep_vba=keep_vba,
        )
        sheets = wb.sheetnames
        wb.close()

        if nome_sheet not in sheets:
            print(f"❌ Sheet '{nome_sheet}' não encontrada. Disponíveis: {', '.join(sheets)}")
            return None

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


def atribuir_funcao_usuario(df_filtrado, session, sistema_desejado):
    """
    Atribui AGR_NAME ao UTILIZADOR via SU10.
    STATUS final definido pelo histórico do wnd[0]/sbar.
    """
    if df_filtrado is None or df_filtrado.empty:
        return df_filtrado

    total = len(df_filtrado)
    tempo_total_inicio = time.time()

    resposta = input("Deseja lançar essas funções no SAP? [S/N]: ").strip().upper()
    if resposta != "S":
        print("❌ Lançamento cancelado pelo utilizador.")
        return df_filtrado

    for i, (idx, row) in enumerate(df_filtrado.iterrows(), 1):
        inicio = time.time()
        eventos_status = []

        utilizador = texto_limpo(row["UTILIZADOR"])
        sistema = texto_limpo(row["SISTEMA"])
        agr_name = texto_limpo(row["AGR_NAME"])

        print(f"\n🔧 {i}/{total} - Utilizador: {utilizador} | Sistema: {sistema} | AGR_NAME: {agr_name}")

        if valor_vazio(row["ID"]):
            msg = "ID vazio."
            marcar_resultado(df_filtrado, idx, "ERRO", msg)
            print(f"❌ {msg}")
            continue

        if not utilizador:
            msg = "UTILIZADOR vazio."
            marcar_resultado(df_filtrado, idx, "ERRO", msg)
            print(f"❌ {msg}")
            continue

        if not sistema:
            msg = "SISTEMA vazio."
            marcar_resultado(df_filtrado, idx, "ERRO", msg)
            print(f"❌ {msg}")
            continue

        if not agr_name:
            msg = "AGR_NAME vazio."
            marcar_resultado(df_filtrado, idx, "ERRO", msg)
            print(f"❌ {msg}")
            continue

        try:
            sistema_conectado = texto_limpo(session.Info.SystemName).upper()
            if sistema_conectado != sistema_desejado:
                msg = (
                    f"Sistema SAP incorreto: esperado {sistema_desejado}, "
                    f"conectado a {sistema_conectado}"
                )
                marcar_resultado(df_filtrado, idx, "ERRO", msg)
                print(f"❌ {msg}")
                continue

            # 1) Abre SU10
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
                marcar_resultado(df_filtrado, idx, "ERRO", msg)
                print(f"❌ {msg}")
                continue

            # 2) Preenche utilizador e seleciona
            campo = session.findById(campo_utilizador)
            campo.text = ""
            campo.text = utilizador
            campo.caretPosition = len(utilizador)

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
                marcar_resultado(df_filtrado, idx, "ERRO", montar_msg_final(eventos_status) or msg_sel)
                print(f"❌ {montar_msg_final(eventos_status)}")
                continue

            # 3) Vai para tab de funções
            session.findById(tab_funcoes).select()
            time.sleep(0.40)
            capturar_status_bar(session, eventos_status, origem="ABERTURA_TAB_FUNCOES", tentativas=4, espera=0.20)

            shell = esperar_elemento(session, shell_funcoes, tentativas=20, espera=0.5)
            if not shell:
                msg = "Não foi possível abrir a aba de funções no SU10."
                marcar_resultado(df_filtrado, idx, "ERRO", msg)
                print(f"❌ {msg}")
                continue

            # 4) Preenche subsystem e AGR_NAME
            shell.modifyCell(0, "SUBSYSTEM", sistema)
            shell.modifyCell(0, "AGR_NAME", agr_name)
            shell.currentCellColumn = "AGR_NAME"
            shell.pressEnter()
            time.sleep(0.70)

            tipo_pre, _, _ = capturar_status_bar(
                session,
                eventos_status,
                origem="VALIDACAO_AGR_NAME",
                tentativas=8,
                espera=0.20,
            )

            if normalizar_valor(tipo_pre) in ("E", "A", "X"):
                msg_final = montar_msg_final(eventos_status)
                marcar_resultado(df_filtrado, idx, "ERRO", msg_final)
                print(f"❌ {msg_final}")
                continue

            # 5) Save - captura imediata antes de qualquer navegação
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(0.40)

            capturar_status_bar(
                session,
                eventos_status,
                origem="SAVE_IMEDIATO",
                tentativas=8,
                espera=0.20,
            )

            # 6) Trata popups e volta a capturar sbar
            tratar_popups_pos_save(session, eventos_status, max_popups=5)

            capturar_status_bar(
                session,
                eventos_status,
                origem="SAVE_FINAL",
                tentativas=10,
                espera=0.25,
            )

            # 7) Decide resultado final
            status_final = decidir_status_pelo_historico(eventos_status)
            msg_final = montar_msg_final(eventos_status)

            marcar_resultado(df_filtrado, idx, status_final, msg_final)

            duracao = time.time() - inicio
            print(f"{'✅' if status_final == 'CONCLUÍDO' else '❌'} {msg_final} | Tempo: {duracao:.1f}s")

        except Exception as e:
            msg = str(e)
            marcar_resultado(df_filtrado, idx, "ERRO", msg)
            print(f"❌ Erro ao atribuir '{agr_name}' a '{utilizador}': {msg}")

        finally:
            voltar_para_inicio(session)

    tempo_total = time.time() - tempo_total_inicio
    print(f"\n⏱️ Tempo total: {tempo_total:.1f}s")

    status_norm = df_filtrado["STATUS"].apply(normalizar_valor)
    total_ok = (status_norm == "CONCLUIDO").sum()
    total_erro = (status_norm == "ERRO").sum()
    print(f"📊 Total concluído: {total_ok} | Com erro: {total_erro}")

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

def executar(ambiente):
    print(f"✅ Processo selecionado: {NOME_SCRIPT}")
    print(f"📄 Script atual: {NOME_SCRIPT} | Sheet alvo: '{NOME_SHEET}'")

    caminho = selecionar_ficheiro_excel()
    if not caminho:
        return False

    df = ler_ficheiro(caminho, NOME_SHEET)
    if df is None:
        return False

    sistema_desejado = MAPA_SISTEMA.get(ambiente)
    if not sistema_desejado:
        print(f"❌ Ambiente inválido: {ambiente}. Use: {', '.join(MAPA_SISTEMA.keys())}")
        return False

    session = conectar_sap(sistema_desejado)
    if not session:
        return False

    df_pend = filtrar_pendentes(df)
    if df_pend.empty:
        return True

    df_proc = atribuir_funcao_usuario(df_pend.copy(), session, sistema_desejado)

    ok_save = gravar_preservando_formatacao(caminho, NOME_SHEET, df_proc)
    return ok_save


###################################################################################
# BLOCO 11: EXECUÇÃO DIRETA
###################################################################################

if __name__ == "__main__":
    executar("CUA")