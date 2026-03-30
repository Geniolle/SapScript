import pandas as pd
from pathlib import Path
import win32com.client
import time

###################################################################################
# Função principal obrigatória para o Cockpit
###################################################################################

def executar(ambiente):
    print(f"\n🚀 Iniciando verificação de cadeias de pesquisa no ambiente {ambiente}...")

    CAMINHO_EXCEL = Path(r"C:\SAP Script\Processos\Cadeias de Pesquisa\Script_Atribuir_Cadeias_Pesquisa.xlsx")
    if not CAMINHO_EXCEL.exists():
        print(f"❌ Ficheiro não encontrado: {CAMINHO_EXCEL}")
        return

    try:
        df = pd.read_excel(CAMINHO_EXCEL, engine="openpyxl")
    except Exception as e:
        print(f"❌ Erro ao ler o ficheiro Excel: {e}")
        return

    colunas = [col for col in df.columns if "NOME CADEIA DE PESQUISA" in col.upper()]
    if not colunas:
        print("❌ Coluna 'NOME CADEIA DE PESQUISA' não encontrada.")
        return

    nome_coluna = colunas[0]
    df[nome_coluna] = df[nome_coluna].astype(str).str.strip()

    # Garantir colunas STATUS e MSG como texto
    df["STATUS"] = df.get("STATUS", "").astype(str)
    df["MSG"] = df.get("MSG", "").astype(str)

    # Criar lista única de cadeias para processar
    cadeias = df[nome_coluna].dropna().unique().tolist()

    print(f"🔍 Total de cadeias a verificar: {len(cadeias)}")

    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print("❌ Erro ao conectar ao SAP GUI. Verifique se o SAP está aberto e com sessão ativa.")
        print(f"Detalhes: {e}")
        return

    for cadeia in cadeias:
        existe = verificar_cadeia_sap(session, cadeia)
        status_val = "ERRO" if existe else "OK"
        msg_val = "Não encontrada na TPAMA" if existe else "Cadeia existente"

        df.loc[df[nome_coluna] == cadeia, "STATUS"] = status_val
        df.loc[df[nome_coluna] == cadeia, "MSG"] = msg_val

        print(f"{'❌' if existe else '✅'} {cadeia}")

    try:
        df.to_excel(CAMINHO_EXCEL, index=False)
        print(f"\n💾 Resultados atualizados no ficheiro: {CAMINHO_EXCEL.name}")
    except Exception as e:
        print(f"❌ Erro ao guardar o ficheiro: {e}")

###################################################################################
# Verifica existência da cadeia na TPAMA via SAP GUI
###################################################################################

def verificar_cadeia_sap(session, nome_cadeia):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16H"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "TPAMA"
        session.findById("wnd[0]").sendVKey(0)

        campo = "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]"
        session.findById(campo).text = nome_cadeia
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # ⏳ Esperar por mensagem significativa no status bar
        for _ in range(10):
            status = session.findById("wnd[0]/sbar").Text.strip().lower()
            if any(msg in status for msg in ["no values", "nenhuma", "não existe", "erro", "not found"]):
                return False
            if status and not status.startswith("seleção"):
                return True
            time.sleep(1)

        return False

    except Exception as e:
        print(f"⚠️ Erro ao verificar '{nome_cadeia}': {e}")
        return False
