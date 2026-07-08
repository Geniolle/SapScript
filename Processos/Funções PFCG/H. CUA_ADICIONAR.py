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
    """
    if df_filtrado is None or df_filtrado.empty:
        return df_filtrado

    # 1. Agrupar em memória por UTILIZADOR e SISTEMA
    grupos = df_filtrado.groupby(["UTILIZADOR", "SISTEMA"], sort=False)
    total_grupos = len(grupos)
    
    total_linhas_pendentes = len(df_filtrado)
    total_roles_distintas = df_filtrado["AGR_NAME"].nunique()
    
    print(f"\n📋 Utilizadores a processar agrupados: {total_grupos}")
    print(f"📋 Linhas pendentes: {total_linhas_pendentes}")
    print(f"📋 Roles distintas: {total_roles_distintas}")

    if not modo_nao_interativo and pedir_confirmacao:
        resposta = input("Deseja lançar essas funções no SAP? [S/N]: ").strip().upper()
        if resposta != "S":
            print("❌ Lançamento cancelado pelo utilizador.")
            return df_filtrado

    tempo_total_inicio = time.time()

    for idx_grupo, ((utilizador, sistema), df_grupo) in enumerate(grupos, 1):
        inicio = time.time()
        eventos_status = []
        
        # Obter a lista de roles únicas a adicionar para este utilizador e sistema
        roles_list = list(dict.fromkeys([str(r).strip() for r in df_grupo["AGR_NAME"] if str(r).strip()]))
        
        print("\n======================================================================")
        print(f"▶ [{idx_grupo}/{total_grupos}] INICIANDO UTILIZADOR: {utilizador} | Sistema: {sistema} | Roles: {len(roles_list)}")
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
                print(f"🔴 ERRO: {msg} ⏱️ (Tempo: {duracao_str})")
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
                print(f"🔴 ERRO: {msg_final} ⏱️ (Tempo: {duracao_str})")
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
                print(f"🔴 ERRO: {msg} ⏱️ (Tempo: {duracao_str})")
                continue

            # 4) Preenche subsystem e AGR_NAME para cada uma das roles
            print(f"├─ Preparando inserção de {len(roles_list)} role(s)...")
            role_errors = {}
            row_idx = 0

            for r_idx, role_name in enumerate(roles_list):
                print(f"├─ Inserindo role {r_idx+1}/{len(roles_list)}: {role_name}")
                
                # Procurar a primeira linha vazia a partir de row_idx
                while row_idx < shell.rowCount:
                    subsys = str(shell.getCellValue(row_idx, "SUBSYSTEM")).strip()
                    agr = str(shell.getCellValue(row_idx, "AGR_NAME")).strip()
                    if not subsys and not agr:
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
                        print(f"│  ⚠️ Falha na validação da role '{role_name}': {err_msg}")
                        
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
                    print(f"│  ⚠️ Erro técnico ao inserir role '{role_name}': {err_msg}")

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
            if salvou_com_sucesso and roles_ok == len(roles_list):
                print(f"🟢 SUCESSO: Utilizador tratado por completo! Roles: {roles_ok}/{len(roles_list)} ⏱️ (Tempo: {duracao_str})")
            else:
                print(f"🔴 ERRO: Atribuição parcial ou falha na gravação. Roles: {total_ok}/{len(roles_list)} com sucesso. ⏱️ (Tempo: {duracao_str})")

        except Exception as e:
            msg = str(e)
            for idx_row in df_grupo.index:
                marcar_resultado(df_filtrado, idx_row, "ERRO", msg)
            duracao_str = formatar_tempo(time.time() - inicio)
            print(f"🔴 ERRO: {msg} ⏱️ (Tempo: {duracao_str})")

        finally:
            voltar_para_inicio(session)

    tempo_total = time.time() - tempo_total_inicio
    print(f"\n⏱️ Tempo total: {formatar_tempo(tempo_total)}")

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

    df_proc = atribuir_funcao_usuario(
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