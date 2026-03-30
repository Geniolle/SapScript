###################################################################################
# FI01 — Importação de Bancos (BNKA) a partir de Excel
#
# Funcionalidades:
#  - Seleção de arquivo e sheet via pop-up.
#  - Mapeamento estrito de cabeçalhos.
#  - Normalização e validação de dados (banco, país).
#  - Detetor de banco "já existe" com opção de atualização automática via FI02.
#  - Gravação segura no SAP (validação de campo).
#  - Registro de logs e atualização do Excel com status e mensagens.
###################################################################################

import os
import sys
import time
import traceback
import math
from datetime import datetime
from typing import Optional, List

import pandas as pd
import win32com.client

# GUI (pop-ups)
import tkinter as tk
from tkinter import filedialog, messagebox

###################################################################################
# CONFIGURAÇÕES
###################################################################################

# Se for usado num ambiente controlado, defina o sistema SAP esperado (ex: "S4D", "S4Q", "S4P")
MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}

# Atualizar automaticamente via FI02 quando o banco "já existe" no SAP
UPDATE_IF_EXISTS = True

# Comprimento desejado da chave do banco (BANKL) por país (preenchido com zeros)
PAD_LENGTH_BY_COUNTRY = {
    "PT": 8,
}

# Limites de caracteres para evitar erros do SAP
TRUNCATE_FIELDS = True
LIM_BNKA_BANKA = 60
LIM_BNKA_STRAS = 60
LIM_BNKA_ORT01 = 35
LIM_BNKA_BRNCH = 15
LIM_BNKA_SWIFT = 11

# Colunas obrigatórias no Excel (após mapeamento)
REQUIRED_FIELDS = ["BANKS", "BANKL", "BANKA"]

# Mapeamento de cabeçalhos: Apenas os nomes fornecidos serão aceites
HEADER_MAP = {
    "BANKS": "BANKS",
    "BANKL": "BANKL",
    "BANKA": "BANKA",
    "STRAS": "STRAS",
    "ORT01": "ORT01",
    "BRNCH": "BRNCH",
    "SWIFT": "SWIFT",
}

###################################################################################
# UTILITÁRIOS GERAIS
###################################################################################

def _now() -> str: return datetime.now().strftime("%H:%M:%S")
def log_info(msg: str): print(f"{_now()}\tINFO\t{msg}")
def log_warn(msg: str): print(f"{_now()}\tAVISO\t{msg}")
def log_err(msg: str): print(f"{_now()}\tERRO\t{msg}")

class Lap:
    """Ferramenta para medir tempo em milissegundos."""
    def __init__(self): self.t0 = time.time()
    def ms(self) -> int: return int((time.time() - self.t0) * 1000)

def norm_header(s: str) -> str:
    """Normaliza o nome do cabeçalho para uso no mapeamento."""
    return str(s).strip().upper()

def mapear_cabecalhos(cols_norm: List[str]) -> List[str]:
    """Mapeia os cabeçalhos normalizados para os nomes de campo do SAP."""
    mapped = []
    for c in cols_norm:
        if c in HEADER_MAP:
            mapped.append(HEADER_MAP[c])
        else:
            mapped.append(c) # Mantém o nome se não estiver no mapeamento
    return mapped

def get_str(row: dict, key: str) -> str:
    """Lê um valor da linha de forma segura e retorna uma string."""
    val = row.get(key, "")
    if pd.isna(val) if isinstance(val, (float, pd.Series)) else val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val)).strip()
    return str(val).strip()

def trunc(s: str, lim: int | None) -> str:
    """Trunca uma string para o limite definido."""
    if not TRUNCATE_FIELDS or lim is None:
        return s
    return s[:lim]

###################################################################################
# GESTÃO DE ARQUIVOS EXCEL
###################################################################################

def escolher_ficheiro_excel(initial_dir: Optional[str] = None) -> str:
    """Abre uma janela pop-up para o utilizador selecionar um arquivo Excel."""
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    kwargs = {"title": "Selecionar arquivo Excel", "filetypes": [("Excel files", "*.xlsx;*.xlsm;*.xls")]}
    if initial_dir and os.path.isdir(initial_dir): kwargs["initialdir"] = initial_dir
    caminho = filedialog.askopenfilename(**kwargs)
    root.destroy()
    return caminho

def escolher_sheet_popup(opcoes: List[str]) -> Optional[str]:
    """Abre uma janela pop-up para o utilizador selecionar uma sheet."""
    if not opcoes: return None
    if len(opcoes) == 1: return opcoes[0]
    sel = {"value": None}
    def on_ok():
        idxs = lb.curselection()
        if not idxs:
            messagebox.showwarning("Seleção", "Por favor, selecione uma sheet."); return
        sel["value"] = opcoes[idxs[0]]; win.destroy()
    win = tk.Tk(); win.title("Selecionar Sheet"); win.attributes("-topmost", True); win.geometry("360x400")
    tk.Label(win, text="Escolha a folha (sheet) a importar:", font=("Segoe UI", 10)).pack(padx=10, pady=(10, 6), anchor="w")
    lb = tk.Listbox(win, selectmode="single", height=min(15, len(opcoes)))
    for name in opcoes: lb.insert(tk.END, name)
    lb.pack(fill="both", expand=True, padx=10, pady=6); lb.selection_set(0)
    tk.Button(win, text="OK", command=on_ok).pack(pady=(0, 10))
    win.mainloop()
    return sel["value"]

def carregar_excel(path: str, sheet: Optional[str] = None) -> pd.DataFrame:
    """Lê e valida o arquivo Excel, aplicando o mapeamento de cabeçalhos."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {path}")
    
    if sheet is None:
        xls = pd.ExcelFile(path)
        chosen = escolher_sheet_popup(xls.sheet_names) or xls.sheet_names[0]
        log_info(f"🗂️ Folha selecionada: {chosen}")
        sheet = chosen

    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    cols_norm = [norm_header(c) for c in df.columns]
    df.columns = mapear_cabecalhos(cols_norm)
    
    # Adicionar colunas de status se não existirem
    if "STATUS" not in df.columns: df["STATUS"] = ""
    if "MSG" not in df.columns: df["MSG"] = ""

    # Aplicar regras de padding se as colunas existirem
    if "BANKS" in df.columns and "BANKL" in df.columns:
        df["BANKS"] = df["BANKS"].astype(str).str.upper()
        df["BANKL"] = df["BANKL"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        for pais, length in PAD_LENGTH_BY_COUNTRY.items():
            mask = df["BANKS"].eq(pais.upper())
            df.loc[mask, "BANKL"] = df.loc[mask, "BANKL"].str.zfill(length)

    missing = [c for c in REQUIRED_FIELDS if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias em falta: {missing}. Cabeçalhos lidos: {list(df.columns)}")
    
    return df

def salvar_excel_seguro(path: str, df: pd.DataFrame) -> bool:
    """Salva o DataFrame de volta no arquivo Excel com tratamento de erros."""
    try:
        df.to_excel(path, index=False)
        log_info(f"💾 Excel guardado: {path}")
        return True
    except Exception as e:
        log_err(f"Falha ao gravar Excel: {e}")
        return False

###################################################################################
# INTERAÇÃO COM O SAP
###################################################################################

def obter_sessao_ativa():
    """Conecta-se à sessão SAP ativa mais recente."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        for conn_idx in range(application.Children.Count):
            connection = application.Children(conn_idx)
            if connection.Children.Count > 0:
                return connection.Children(0)
    except Exception as e:
        raise RuntimeError(f"Não foi possível encontrar uma sessão SAP ativa. Verifique se o SAP Logon está aberto. Erro: {e}")

def ir_transacao(session, tcode: str):
    """Navega para uma transação SAP."""
    session.findById("wnd[0]/tbar[0]/okcd").text = tcode
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(0.2)

def sbar_text(session) -> str:
    """Lê o texto da barra de status do SAP."""
    try: return session.findById("wnd[0]/sbar").Text
    except: return ""

def ensure_field(session, path: str, value: str, read_attr: str = "text", retries: int = 5, delay: float = 0.2) -> bool:
    """Escreve e confirma que o campo ficou com o valor (lendo de volta)."""
    try:
        ctrl = session.findById(path)
        for _ in range(retries):
            ctrl.text = value
            time.sleep(0.05)
            got = getattr(ctrl, read_attr, None) or ""
            if got.strip() == value.strip():
                return True
            time.sleep(delay)
    except Exception as e:
        log_err(f"Falha ao escrever no campo '{path}'. Erro: {e}")
    return False

def fechar_popups(session):
    """Tenta fechar pop-ups comuns do SAP."""
    for idx in (0, 1, 2):
        try: session.findById(f"wnd[{idx}]").sendVKey(12) # VKey 12 = Cancelar
        except: pass
        try: session.findById(f"wnd[{idx}]").sendVKey(0) # VKey 0 = Enter
        except: pass
    time.sleep(0.2)


def processar_linha_banco(session, row: dict) -> tuple[str, str]:
    """Processa uma única linha do Excel para criar/atualizar um banco no SAP."""
    country = get_str(row, "BANKS")
    bankl = get_str(row, "BANKL")
    banka = get_str(row, "BANKA")
    stras = get_str(row, "STRAS")
    ort01 = get_str(row, "ORT01")
    brnch = get_str(row, "BRNCH")
    swift = get_str(row, "SWIFT")

    # Aplica truncagem de campos
    if TRUNCATE_FIELDS:
        banka = trunc(banka, LIM_BNKA_BANKA)
        stras = trunc(stras, LIM_BNKA_STRAS)
        ort01 = trunc(ort01, LIM_BNKA_ORT01)
        brnch = trunc(brnch, LIM_BNKA_BRNCH)
        swift = trunc(swift, LIM_BNKA_SWIFT)

    if not country or not bankl or not banka:
        return ("ERRO", "Dados obrigatórios em falta (BANKS/BANKL/BANKA).")

    # Navega para a transação de criação
    ir_transacao(session, "/nFI01")
    ensure_field(session, "wnd[0]/usr/ctxtBNKA-BANKS", country)
    ensure_field(session, "wnd[0]/usr/ctxtBNKA-BANKL", bankl)
    session.findById("wnd[0]").sendVKey(0)
    
    # Aumentar o tempo de espera para dar tempo à tela de carregar
    time.sleep(3.0)

    # Verifica se o banco já existe
    msg_init = (sbar_text(session) or "").upper()
    if any(k in msg_init for k in ["JÁ EXISTE", "ALREADY EXISTS"]):
        if UPDATE_IF_EXISTS:
            log_info("Banco já existe, a tentar atualizar via FI02...")
            return atualizar_banco_FI02(session, country, bankl, banka, stras, ort01, brnch, swift)
        else:
            return ("EXISTE", "Banco já existe e a atualização está desativada.")

    # Preenche os campos para criação na nova ordem
    ensure_field(session, "wnd[0]/usr/txtBNKA-STRAS", stras)
    ensure_field(session, "wnd[0]/usr/txtBNKA-ORT01", ort01)
    ensure_field(session, "wnd[0]/usr/txtBNKA-BRNCH", brnch)
    ensure_field(session, "wnd[0]/usr/txtBNKA-SWIFT", swift)
    
    # Preenche BANKA por último
    if not ensure_field(session, "wnd[0]/usr/txtBNKA-BANKA", banka):
        return ("ERRO", "Falha ao preencher 'Nome do Banco' (BNKA-BANKA).")

    # Grava e verifica o resultado
    session.findById("wnd[0]").sendVKey(11)  # Gravar
    time.sleep(0.2)
    final_msg = sbar_text(session) or "Sem mensagem na status bar."
    if any(k in final_msg.upper() for k in ["GRAVADO", "SALVO", "CRIADO", "SAVED", "CREATED"]):
        return ("OK", final_msg)
    elif "ERRO" in final_msg.upper():
        return ("ERRO", final_msg)
    return ("AVISO", final_msg)

def atualizar_banco_FI02(session, country: str, bankl: str,
                         banka: str, stras: str, ort01: str, brnch: str, swift: str) -> tuple[str, str]:
    """Atualiza os dados de um banco existente via FI02."""
    ir_transacao(session, "/nFI02")
    ensure_field(session, "wnd[0]/usr/ctxtBNKA-BANKS", country)
    ensure_field(session, "wnd[0]/usr/ctxtBNKA-BANKL", bankl)
    session.findById("wnd[0]").sendVKey(0)

    # Aumentar o tempo de espera para dar tempo à tela de carregar
    time.sleep(3.0)

    # Preenche os campos na nova ordem
    ensure_field(session, "wnd[0]/usr/txtBNKA-STRAS", stras)
    ensure_field(session, "wnd[0]/usr/txtBNKA-ORT01", ort01)
    ensure_field(session, "wnd[0]/usr/txtBNKA-BRNCH", brnch)
    ensure_field(session, "wnd[0]/usr/txtBNKA-SWIFT", swift)

    # Preenche BANKA por último
    if not ensure_field(session, "wnd[0]/usr/txtBNKA-BANKA", banka):
        return ("ERRO", "Falha ao preencher 'Nome do Banco' (BNKA-BANKA).")

    session.findById("wnd[0]").sendVKey(11) # Gravar
    time.sleep(0.2)
    final_msg = sbar_text(session) or "Sem mensagem na status bar."
    if any(k in final_msg.upper() for k in ["GRAVADO", "SALVO", "ALTERADO", "SAVED", "UPDATED"]):
        return ("ATUALIZADO", final_msg)
    return ("AVISO", final_msg)


###################################################################################
# PIPELINE PRINCIPAL
###################################################################################

def run_import_bancos(excel_path: str, ambiente_cockpit: Optional[str] = None):
    log_warn("▶️ Início: Importação de Bancos (FI01)")
    cron = Lap()

    try:
        # Carregar Excel
        log_info(f"📄 A ler Excel: {os.path.basename(excel_path)}")
        df = carregar_excel(excel_path)

        # Conectar ao SAP
        log_info("🔎 A procurar sessão SAP ativa...")
        session = obter_sessao_ativa()

        sysname = session.Info.SystemName
        user = session.Info.User
        client = session.Info.Client
        log_info(f"🔧 Conectado a: {sysname} | Cliente: {client} | Utilizador: {user}")

        # Processar linhas
        total = len(df); ok = err = aviso = existe = atualizado = 0
        for i, row in df.iterrows():
            trow = Lap()
            
            country = get_str(row, "BANKS")
            bankl = get_str(row, "BANKL")
            banka = get_str(row, "BANKA")
            log_info(f"🔧 Linha {i+1}/{total} | Banks={country} | BankL={bankl} | BankA={banka}")
            
            try:
                status, msg = processar_linha_banco(session, row)
                df.at[i, "STATUS"] = status
                df.at[i, "MSG"] = msg

                if status == "OK": ok += 1
                elif status == "ERRO": err += 1
                elif status == "EXISTE": existe += 1
                elif status == "ATUALIZADO": atualizado += 1
                else: aviso += 1
                
                log_info(f"⏱️ Tempo linha: {trow.ms()} ms | STATUS={status} | MSG={msg}")

            except Exception as e:
                df.at[i, "STATUS"] = "ERRO"
                df.at[i, "MSG"] = f"{type(e).__name__}: {e}"
                err += 1
                log_err(f"Falha na linha {i+1}: {e}")
                fechar_popups(session)

    except Exception as e:
        log_err(f"Execução abortada: {e}")
        traceback.print_exc()
        return

    # Salvar resultados
    salvar_excel_seguro(excel_path, df)
    log_warn(f"⏹️ Fim: Importação | OK={ok} | Atualizado={atualizado} | Existe={existe} | Aviso={aviso} | Erro={err} | Total={total} | Tempo={cron.ms()} ms")


###################################################################################
# EXECUÇÃO DO SCRIPT
###################################################################################

if __name__ == "__main__":
    AMBIENTE = None # Defina "DEV", "QAD", "PRD" ou "CUA" se precisar de validação
    
    initial_dir = r"C:\SAP Script\Processos\Chave de Banco"
    excel_path = escolher_ficheiro_excel(initial_dir)

    if excel_path:
        run_import_bancos(excel_path, ambiente_cockpit=AMBIENTE)
    else:
        log_err("❌ Nenhum arquivo selecionado. A encerrar.")