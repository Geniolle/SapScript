import pandas as pd
from pathlib import Path
import win32com.client
import time
import json
import tkinter as tk
from tkinter import filedialog

###################################################################################
# Função principal obrigatória para o Cockpit
###################################################################################

def executar(ambiente):
    print(f"\n🚀 Iniciando verificação de cadeias de pesquisa no ambiente {ambiente}...")

    # Diretorias de cache
    raiz_dir = Path(__file__).resolve().parent.parent.parent
    cache_dir = raiz_dir / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    config_file = cache_dir / "validar_cadeia_config.json"

    caminho_excel = None
    ultimo_caminho = None

    # Tenta ler o caminho guardado
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                config = json.load(f)
                caminho_excel = config.get("caminho_excel")
                ultimo_caminho = config.get("ultimo_caminho")
        except Exception as e:
            print(f"⚠️ Erro ao ler a configuração: {e}")

    # Se não existe caminho configurado ou o ficheiro não existe, abre popup
    if not caminho_excel or not Path(caminho_excel).exists():
        print("📁 Selecione o ficheiro Excel (janela em primeiro plano)...")
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.attributes("-topmost", True)
        root.focus_force()
        root.update()

        # Determina o diretório inicial
        initial_dir = ultimo_caminho if (ultimo_caminho and Path(ultimo_caminho).exists()) else str(Path(__file__).resolve().parent)

        caminho_selecionado = filedialog.askopenfilename(
            parent=root,
            title="Selecione o ficheiro de Cadeias de Pesquisa",
            initialdir=initial_dir,
            filetypes=[("Ficheiros Excel", "*.xlsx"), ("Todos os ficheiros", "*.*")],
        )
        root.destroy()

        if not caminho_selecionado:
            print("❌ Operação cancelada. Nenhum ficheiro selecionado.")
            return

        caminho_excel = caminho_selecionado
        ultimo_caminho = str(Path(caminho_selecionado).parent)

        # Guarda a configuração
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump({
                    "caminho_excel": caminho_excel,
                    "ultimo_caminho": ultimo_caminho
                }, f, indent=4, ensure_ascii=False)
            print(f"💾 Caminho configurado e guardado.")
        except Exception as e:
            print(f"⚠️ Erro ao guardar a configuração: {e}")

    CAMINHO_EXCEL = Path(caminho_excel)
    print(f"✅ Ficheiro Excel a utilizar: {CAMINHO_EXCEL}")

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
        existe, status_sap = verificar_cadeia_sap(session, cadeia)
        status_val = "OK" if existe else "ERRO"
        msg_val = status_sap if status_sap else ("Cadeia existente" if existe else "Não encontrada na TPAMA")

        df.loc[df[nome_coluna] == cadeia, "STATUS"] = status_val
        df.loc[df[nome_coluna] == cadeia, "MSG"] = msg_val

        print(f"{'✅' if existe else '❌'} {cadeia} - {status_sap}")

    # Dar o comando /n no final da pesquisa em SAP para limpar a sessão
    try:
        okcd = _esperar_objeto(session, "wnd[0]/tbar[0]/okcd", timeout=2.0)
        if okcd:
            okcd.text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)
    except Exception as e:
        print(f"⚠️ Erro ao regressar ao ecrã inicial do SAP (/n): {e}")

    try:
        df.to_excel(CAMINHO_EXCEL, index=False)
        print(f"\n💾 Resultados atualizados no ficheiro: {CAMINHO_EXCEL.name}")
    except Exception as e:
        print(f"❌ Erro ao guardar o ficheiro: {e}")

###################################################################################
# Helpers SAP
###################################################################################

def _safe_find(session, sap_id):
    try:
        return session.findById(sap_id)
    except Exception:
        return None

def _esperar_objeto(session, sap_id, timeout=4.0, pausa=0.05):
    limite = time.time() + timeout
    while time.time() < limite:
        obj = _safe_find(session, sap_id)
        if obj:
            return obj
        time.sleep(pausa)
    return None

###################################################################################
# Verifica existência da cadeia na TPAMA via SAP GUI
###################################################################################

def verificar_cadeia_sap(session, nome_cadeia):
    try:
        # Normalizar e truncar a 20 caracteres (PANAM) para corresponder à chave no SAP
        import unicodedata
        import re
        nome_sanitizado = ""
        if nome_cadeia:
            n_cad = "".join(c for c in unicodedata.normalize("NFKD", nome_cadeia) if not unicodedata.combining(c))
            n_cad = re.sub(r"[^A-Za-z0-9_\-\.\/ ]", "", n_cad)
            nome_sanitizado = n_cad.strip().upper()[:20]

        # Reset de transação para garantir que estamos limpos
        try:
            okcd = _esperar_objeto(session, "wnd[0]/tbar[0]/okcd", timeout=2.0)
            if okcd:
                okcd.text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.2)
        except Exception:
            pass

        # Ir para a SE16H
        okcd = _esperar_objeto(session, "wnd[0]/tbar[0]/okcd")
        if not okcd:
            print("⚠️ Campo de comando okcd não encontrado.")
            return False, "Campo de comando okcd não encontrado"
        okcd.text = "/NSE16H"
        session.findById("wnd[0]").sendVKey(0)

        # Esperar e preencher o nome da tabela
        gd_tab = _esperar_objeto(session, "wnd[0]/usr/ctxtGD-TAB")
        if not gd_tab:
            gd_tab = _esperar_objeto(session, "wnd[0]/usr/ctxtDATABROWSE-TABLENAME") or _esperar_objeto(session, "wnd[0]/usr/ctxtTABNAME")
        
        if not gd_tab:
            print("⚠️ Campo da tabela não encontrado na SE16H.")
            return False, "Campo da tabela não encontrado na SE16H"

        gd_tab.text = "TPAMA"
        session.findById("wnd[0]").sendVKey(0)

        # Esperar que os campos da tabela carreguem na SE16H
        campo_id = "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]"
        campo = _esperar_objeto(session, campo_id, timeout=3.0)
        if not campo:
            print("⚠️ Campo PANAM (Cadeia de Pesquisa) não carregado a tempo.")
            return False, "Campo PANAM não carregado a tempo"

        campo.text = nome_sanitizado
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # ⏳ Esperar por mensagem significativa no status bar ou mudança de ecrã (ALV)
        status = ""
        for _ in range(15):
            # 1. Verifica se navegou para o ecrã de resultados (ALV Grid)
            wnd0 = _safe_find(session, "wnd[0]")
            wnd_title = wnd0.Text.strip().lower() if wnd0 else ""
            if any(t in wnd_title for t in ["entradas encontradas", "entries found", "exibição"]):
                sbar_obj = _safe_find(session, "wnd[0]/sbar")
                status = sbar_obj.Text.strip() if sbar_obj else ""
                return True, status or "valor encontrado"

            # 2. Verifica mensagens de erro ou insucesso na barra de status
            sbar_obj = _safe_find(session, "wnd[0]/sbar")
            status = sbar_obj.Text.strip() if sbar_obj else ""
            if status:
                print(f"[SAP_SBAR] {status}")
            status_lower = status.lower()
            if any(msg in status_lower for msg in ["no values", "nenhuma", "nenhum", "não existe", "erro", "not found"]):
                return False, status or "Nenhum valor encontrado"
            if status and not status_lower.startswith("seleção"):
                return True, status or "valor encontrado"
            time.sleep(0.5)

        return False, status or "Nenhum valor encontrado"

    except Exception as e:
        print(f"⚠️ Erro ao verificar '{nome_cadeia}': {e}")
        return False, f"Erro: {e}"
