import pandas as pd
from pathlib import Path
import win32com.client
import time
import unicodedata
import re
# Bibliotecas para o seletor de ficheiros
import tkinter as tk
from tkinter import filedialog

###################################################################################
# BLOCO 1: FUNÇÕES UTILITÁRIAS
###################################################################################

def normalizar_texto(s: str) -> str:
    """Remove acentos, espaços duplicados e coloca em MAIÚSCULAS."""
    s = "" if s is None else str(s)
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s.strip())
    return s.upper()

def encontrar_coluna_nome(df: pd.DataFrame) -> tuple[str, str] | tuple[None, None]:
    """
    Procura a coluna 'NOME CADEIA DE PESQUISA' no DF.
    Devolve (nome_coluna_original, nome_coluna_normalizado) ou (None, None).
    """
    alvo_norm = "NOME CADEIA DE PESQUISA"
    m = {col: normalizar_texto(col) for col in df.columns}
    for original, norm in m.items():
        if alvo_norm == norm:
            return original, norm
    # fallback tolerante: contém as palavras NOME+CADEIA+PESQUISA
    for original, norm in m.items():
        if all(p in norm for p in ["NOME", "CADEIA", "PESQUISA"]):
            return original, norm
    return None, None

def carregar_df_cadeias(caminho: Path) -> tuple[pd.DataFrame, str, str]:
    """
    Lê o Excel e encontra a folha que contém 'NOME CADEIA DE PESQUISA'.
    Tenta primeiro cabeçalho padrão; se falhar, usa a primeira linha como cabeçalho.
    Devolve (df, sheet_name_usado, nome_coluna_original).
    Lança ValueError se não encontrar.
    """
    xls = pd.ExcelFile(caminho)
    erros = []

    for folha in xls.sheet_names:
        # 1) Tentativa normal (cabeçalho na 1ª linha)
        try:
            df = pd.read_excel(caminho, sheet_name=folha, engine="openpyxl", dtype=str).fillna("")
            # normaliza visualmente os cabeçalhos (sem alterar df.columns ainda)
            df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
            nome_col_or, _ = encontrar_coluna_nome(df)
            if nome_col_or:
                return df, folha, nome_col_or
        except Exception as e:
            erros.append((folha, f"header padrão: {e}"))

        # 2) Tentativa alternativa (primeira linha como cabeçalho)
        try:
            bruto = pd.read_excel(caminho, sheet_name=folha, engine="openpyxl", header=None, dtype=str).fillna("")
            if bruto.shape[0] >= 2:
                cabecalhos = [re.sub(r"\s+", " ", str(v).strip()) for v in bruto.iloc[0].tolist()]
                df2 = bruto.iloc[1:].copy()
                df2.columns = cabecalhos
                df2 = df2.reset_index(drop=True).astype(str).fillna("")
                nome_col_or, _ = encontrar_coluna_nome(df2)
                if nome_col_or:
                    return df2, folha, nome_col_or
        except Exception as e:
            erros.append((folha, f"linha como header: {e}"))

    detalhes = "; ".join([f"[{f}: {msg}]" for f, msg in erros]) or "Sem detalhes."
    raise ValueError(f"Não foi possível encontrar a coluna 'NOME CADEIA DE PESQUISA' em nenhuma folha. {detalhes}")

def obter_sessao_sap(ambiente: str):
    """Obtém uma sessão SAP ativa no ambiente desejado (S4D/S4Q/S4P)."""
    mapa = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    sistema = mapa.get(ambiente)
    if not sistema:
        print(f"❌ Ambiente '{ambiente}' não reconhecido. Use DEV, QAD ou PRD.")
        return None

    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        for conn in application.Children:
            for sess in conn.Children:
                if str(sess.Info.SystemName).upper() == sistema:
                    return sess
        print(f"❌ Nenhuma sessão SAP encontrada para '{ambiente}' (Sistema: {sistema}).")
        print("   Abra o SAP Logon e inicie sessão no sistema correto.")
        return None
    except Exception as e:
        print("❌ SAP GUI não disponível ou ocorreu um erro inesperado na ligação.")
        print(f"   Detalhes: {e}")
        return None

###################################################################################
# BLOCO 2: EXECUÇÃO PRINCIPAL CHAMADA PELO COCKPIT
###################################################################################

def executar(ambiente):
    print(f"\n🚀 Iniciando criação de cadeias de pesquisa com STATUS VAZIO ({ambiente})")

    # --- SELEÇÃO DE FICHEIRO FLEXÍVEL ---
    print("Por favor, selecione o ficheiro Excel de Cadeias de Pesquisa na janela que abriu...")
    root = tk.Tk(); root.withdraw()
    caminho_str = filedialog.askopenfilename(
        title="Selecione o ficheiro de Cadeias de Pesquisa",
        filetypes=[("Ficheiros Excel", "*.xlsx"), ("Todos os ficheiros", "*.*")]
    )
    if not caminho_str:
        print("❌ Nenhum ficheiro selecionado. A execução foi cancelada.")
        return
    CAMINHO_EXCEL = Path(caminho_str)
    print(f"✅ Ficheiro a processar: {CAMINHO_EXCEL}")

    if not CAMINHO_EXCEL.exists():
        print(f"❌ Ficheiro não encontrado: {CAMINHO_EXCEL}")
        return

    # --- LEITURA ROBUSTA DO EXCEL ---
    print("\nLendo cabeçalhos do ficheiro e a detetar a folha correta...")
    try:
        df, folha_usada, nome_coluna = carregar_df_cadeias(CAMINHO_EXCEL)
    except Exception as e:
        print(f"❌ {e}")
        return

    print(f"✅ Folha selecionada: {folha_usada}")
    print("Cabeçalhos normalizados com sucesso.")
    # Sanitiza a coluna de nome
    df[nome_coluna] = df[nome_coluna].astype(str).str.strip()

    # Garante STATUS e MSG
    if "STATUS" not in df.columns: df["STATUS"] = ""
    if "MSG" not in df.columns: df["MSG"] = ""

    # Seleciona apenas linhas com STATUS vazio
    cadeias_a_processar = (
        df.loc[df["STATUS"].astype(str).str.strip() == "", nome_coluna]
        .dropna()
        .astype(str)
        .str.strip()
        .replace({"nan": ""})
    )
    cadeias_a_processar = sorted([c for c in cadeias_a_processar.unique().tolist() if c])

    if not cadeias_a_processar:
        print("✅ Nenhuma cadeia com STATUS VAZIO para processar.")
        return

    print(f"\nCadeias a serem criadas ({len(cadeias_a_processar)}): {', '.join(cadeias_a_processar)}")
    ordem = input("📦 Introduz a ordem de transporte (ex: S4DK951842): ").strip()
    if not ordem:
        print("❌ Ordem de transporte não inserida. A execução foi cancelada.")
        return

    # --- SESSÃO SAP ---
    print("\n🔎 A procurar uma sessão SAP ativa...")
    session = obter_sessao_sap(ambiente)
    if not session:
        return
    print(f"✅ Conectado ao SAP: {session.Info.SystemName} (ambiente {ambiente})")
    print(f"👤 Utilizador SAP: {session.Info.User} | Cliente: {session.Info.Client}")

    # --- EXECUÇÃO ---
    for nome_cadeia in cadeias_a_processar:
        # regra técnica: nome da cadeia tem no máximo 20 caracteres no SAP
        nome_cadeia_limite = str(nome_cadeia)[:20]
        sucesso = criar_cadeia_sap(session, nome_cadeia_limite, ordem)

        status_val = "CRIADO" if sucesso else "ERRO"
        msg_val = "Criado com sucesso" if sucesso else "Erro ao criar no SAP"

        df.loc[df[nome_coluna] == nome_cadeia, "STATUS"] = status_val
        df.loc[df[nome_coluna] == nome_cadeia, "MSG"] = msg_val

    # --- GRAVAÇÃO (apenas na folha usada) ---
    try:
        with pd.ExcelWriter(CAMINHO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=folha_usada)
        print(f"\n💾 Folha '{folha_usada}' atualizada em: {CAMINHO_EXCEL.name}")
    except Exception as e:
        print(f"❌ Erro ao guardar o ficheiro: {e}")

###################################################################################
# BLOCO 3: MAPEAMENTO SAP GUI (com limpeza de 20 campos fixos)
###################################################################################

def criar_cadeia_sap(session, nome_cadeia, ordem_transporte):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NOTPM"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/tbar[1]/btn[25]").press()  # Criar
        session.findById("wnd[0]/tbar[1]/btn[5]").press()   # Modo edição

        session.findById("wnd[0]/usr/txtV_TPAMA-PANAM").text = nome_cadeia
        session.findById("wnd[0]/usr/txtV_TPAMA-NOTE").text = nome_cadeia
        session.findById("wnd[0]/usr/txtV_TPAMA-REGEX").text = nome_cadeia
        session.findById("wnd[0]/usr/txtV_TPAMA-REGEX").setFocus()
        session.findById("wnd[0]/usr/txtV_TPAMA-REGEX").caretPosition = len(nome_cadeia)

        session.findById("wnd[0]").sendVKey(0)

        # Limpeza das 20 linhas de mapeamento
        for i in range(20):
            campo = f"wnd[0]/usr/subSUB_PAMA:SAPLPAMI:0210/tblSAPLPAMITC_MAP/txtT_MAP-MXCHAR[3,{i}]"
            try:
                session.findById(campo).text = ""
            except:
                pass

        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()                     # Guardar
        session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = ordem_transporte
        session.findById("wnd[1]/tbar[0]/btn[0]").press()                      # Confirmar transporte

        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)

        print(f"✅ Criada cadeia: {nome_cadeia}")
        return True

    except Exception as e:
        print(f"❌ Erro ao criar cadeia '{nome_cadeia}': {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass
        return False
