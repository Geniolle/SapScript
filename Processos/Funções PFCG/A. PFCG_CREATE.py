# -*- coding: utf-8 -*-

###################################################################################
# A. PFCG_CREATE.py
# PFCG - Criar/Modificar Roles + Atribuir TCODEs + Perfil + Transporte
#
# Regras:
#  - Logger visual estruturado por Etapas
#  - Integração com 'pesquisar_request.py'
#  - Inserção direta e rápida de TCODEs e de Ordem de Transporte
#  - Menu de Request Unificado
###################################################################################

def executar(
    ambiente_cockpit,
    caminho_ficheiro=None,
    request_transporte=None,
    modo_nao_interativo=False,
    pedir_confirmacao=True
):
    import sys
    import os
    import time
    import re
    import unicodedata
    import tkinter as tk

    import pandas as pd
    import win32com.client
    from tkinter import filedialog
    from datetime import datetime
    from math import ceil
    from openpyxl import load_workbook

    tempo_inicio_total = time.time()

    # --- CORREÇÃO DA ESTRUTURA DE PASTAS ---
    dir_atual = os.path.dirname(os.path.abspath(__file__))
    dir_processos = os.path.dirname(dir_atual) 
    if dir_processos not in sys.path:
        sys.path.insert(0, dir_processos)
    # ---------------------------------------

    NOME_SHEET = "PFCG_CREATE"
    SEARCH_HEADER_IN_FIRST_ROWS = 20
    
    COLUNAS_OBRIGATORIAS = {"AGR_NAME", "TEXT", "TCODE", "STATUS", "MSG", "TIMESTEMP"}

    MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    SISTEMA_ESPERADO = MAPA_SISTEMA.get(str(ambiente_cockpit).upper().strip() or "", None)
    if not SISTEMA_ESPERADO:
        raise ValueError(f"Ambiente inválido: '{ambiente_cockpit}'. Use DEV, QAD ou PRD.")

    SLEEP_UI = 0.20
    SLEEP_ACTION = 0.40
    TCODE_BLOCK_SIZE = 20 

    def now_ts():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _sleep(t=SLEEP_UI):
        time.sleep(t)

    def norm_col(s):
        if s is None: s = ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def norm_txt(s):
        if s is None: s = ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def formatar_tempo(segundos):
        h, resto = divmod(segundos, 3600)
        m, s = divmod(resto, 60)
        if h > 0:
            return f"{int(h):02d}h {int(m):02d}m {int(s):02d}s"
        return f"{int(m):02d}m {int(s):02d}s"

    def selecionar_ficheiro():
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title=f"Selecione o ficheiro Excel (sheet '{NOME_SHEET}')",
            filetypes=(("Ficheiros Excel", "*.xlsx"), ("Todos os ficheiros", "*.*"))
        )
        root.destroy()
        return path

    def split_tcodes(raw):
        if not raw: return []
        s = str(raw).replace("\r", "\n").replace("\t", " ").strip().upper()
        parts = re.split(r"[;, \n]+", s)
        out = []
        for p in parts:
            p = p.strip()
            if not p: continue
            if p.startswith("/N") or p.startswith("/O"): p = p[2:].strip()
            if p: out.append(p)
        return list(dict.fromkeys(out))

    ###################################################################################
    # BLOCO 1: LER EXCEL
    ###################################################################################
    if not caminho_ficheiro:
        if modo_nao_interativo:
            raise ValueError("Faltou o parâmetro --xlsx em modo não-interativo.")
        print("📂 Selecione o ficheiro Excel…")
        caminho_ficheiro = selecionar_ficheiro()
        if not caminho_ficheiro:
            print("❌ Operação cancelada.")
            return

    if not os.path.exists(caminho_ficheiro):
        print(f"❌ Ficheiro não encontrado: {caminho_ficheiro}")
        return

    try:
        wb = load_workbook(caminho_ficheiro)
    except Exception as e:
        print(f"❌ Não consegui abrir o Excel: {e}")
        return

    if NOME_SHEET in wb.sheetnames:
        ws = wb[NOME_SHEET]
    else:
        if len(wb.sheetnames) == 1:
            ws = wb[wb.sheetnames[0]]
        else:
            print(f"❌ Sheet '{NOME_SHEET}' não encontrada.")
            wb.close()
            return

    header_row = None
    header_map = {}
    melhor_linha = 0
    max_matches = 0

    for r in range(1, SEARCH_HEADER_IN_FIRST_ROWS + 1):
        row_vals = [norm_col(c.value) for c in ws[r]]
        colunas_encontradas = set(row_vals).intersection(COLUNAS_OBRIGATORIAS)
        qtd_encontradas = len(colunas_encontradas)
        
        if qtd_encontradas > max_matches:
            max_matches = qtd_encontradas
            melhor_linha = r
            
        if qtd_encontradas == len(COLUNAS_OBRIGATORIAS):
            header_row = r
            for idx, name in enumerate(row_vals, start=1):
                if name: header_map[name] = idx
            break

    if not header_row:
        wb.close()
        print("\n❌ Não encontrei a linha de cabeçalho completa.")
        return

    records = []
    def get_cell(row_idx, col_name):
        if col_name not in header_map: return ""
        v = ws.cell(row=row_idx, column=header_map[col_name]).value
        return "" if v is None else str(v).strip()

    for r in range(header_row + 1, ws.max_row + 1):
        agr = get_cell(r, "AGR_NAME")
        if not agr: continue

        records.append({
            "_row": r,
            "AGR_NAME": agr,
            "TEXT": get_cell(r, "TEXT"),
            "TCODE": get_cell(r, "TCODE"),
            "STATUS": get_cell(r, "STATUS"),
            "MSG": get_cell(r, "MSG"),
            "TIMESTEMP": get_cell(r, "TIMESTEMP"),
        })

    df_original = pd.DataFrame(records)
    if df_original.empty:
        wb.close()
        print("⚠️ Não encontrei linhas para processar.")
        return

    df_original["STATUS"] = df_original["STATUS"].apply(norm_txt)
    df_original["TCODE_LIST"] = df_original["TCODE"].apply(split_tcodes)

    df_proc = df_original[df_original["STATUS"] != "CONCLUIDO"].copy()
    if df_proc.empty:
        wb.close()
        print("⚠️ Nada para processar (tudo CONCLUIDO).")
        return

    roles_agrupadas = []
    for agr, grp in df_proc.groupby("AGR_NAME"):
        textos = [t for t in grp["TEXT"] if str(t).strip()]
        texto_final = textos[0] if textos else ""
        tcodes_comb = []
        for t_list in grp["TCODE_LIST"]: tcodes_comb.extend(t_list)
        roles_agrupadas.append({
            "AGR_NAME": agr,
            "TEXT": texto_final,
            "TCODE_LIST": list(dict.fromkeys(tcodes_comb))
        })
        
    df_proc_agrupado = pd.DataFrame(roles_agrupadas)

    # =================================================================================
    # CAPTURA SESSÃO SAP
    # =================================================================================
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = next((sess for conn in application.Children for sess in conn.Children if sess.Info.SystemName.upper() == SISTEMA_ESPERADO), None)
    except Exception as e:
        session = None

    if not session:
        wb.close()
        print(f"❌ Não encontrei sessão do ambiente '{ambiente_cockpit}'.")
        return

    def _safe_find(sap_id):
        try: return session.findById(sap_id)
        except: return None

    def _criar_nova_request_no_sap_local(sess):
        okcd = _safe_find("wnd[0]/tbar[0]/okcd")
        if okcd:
            okcd.text = "/nSE10"
            sess.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)

        print("\nTipo da ordem:")
        print('1 - Ordem customizing')
        print('2 - Ordem workbench')

        while True:
            tipo = input("Digite a opção (1/2): ").strip()
            if tipo in ("1", "2"): break
            print("❌ Opção inválida. Use apenas 1 ou 2.")

        desc = input("Descrição da request (máx 60): ").strip()
        desc = desc[:60] if desc else "REQUEST CRIADA VIA SCRIPT"

        sess.findById("wnd[0]/tbar[1]/btn[6]").press()
        time.sleep(0.4)

        if tipo == "2":
            try: sess.findById("wnd[1]/usr/radKO042-REQ_CONS_K").select()
            except: pass

        sess.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.4)

        try: sess.findById("wnd[1]/usr/txtKO013-AS4TEXT").text = desc
        except: pass
        
        sess.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.6)

        trkorr = None
        for sap_id in ["wnd[0]/usr/lbl[20,9]", "wnd[0]/usr/lbl[1,1]"]:
            try:
                txt = sess.findById(sap_id).Text
                match = re.search(r"\b[A-Z0-9]{3,4}K\d{6,}\b", txt)
                if match: trkorr = match.group(0)
            except: pass
            if trkorr: break

        if okcd:
            okcd.text = "/n"
            sess.findById("wnd[0]").sendVKey(0)

        tipo_txt = "Customizing" if tipo == "1" else "Workbench"
        print("\n✔️ Request criada.")
        print(f"Tipo: {tipo_txt} | Descrição: {desc}")

        if not trkorr:
            trkorr = input("Não consegui extrair a request automaticamente. Cole aqui: ").strip().upper()
        
        print(f"Request: {trkorr}")
        return trkorr

    if not request_transporte and not modo_nao_interativo:
        print("\n============================================================")
        print("🚚 Opções de configuração de Transporte.\n")
        print("   1 - Escreva o número da Request")
        print("   2 - Criar nova ordem de transporte")
        print("   3 - Pesquisar suas request criadas.")
        print("   4 - Prima [Enter] vazio para NÃO transportar")
        print("============================================================")
        
        while True:
            req_input = input("\n👉 Opção: ").strip()
            if req_input in ("1", "2", "3", "4", ""):
                if req_input == "": req_input = "4"
                break
            print("❌ Opção inválida. Use 1, 2, 3, 4 ou apenas pressione Enter.")
        
        if req_input == "1":
            request_transporte = input("🔢 Numero da Request (ex: S4QK900396): ").strip().upper()
        
        elif req_input == "2":
            request_transporte = _criar_nova_request_no_sap_local(session)

        elif req_input == "3":
            try:
                import pesquisar_request
                print("\n🔍 A abrir nova sessão em segundo plano para pesquisar (SE16H)...")
                
                resultados_pesquisa = pesquisar_request.listar_requests(
                    system_name=SISTEMA_ESPERADO, 
                    include_requests=True, 
                    use_new_mode=True,     
                    minimize=True,         
                    close_after=True       
                )
                
                if resultados_pesquisa:
                    escolha = input("\n👉 Digite o número (N) da Request que deseja utilizar (ou Enter para cancelar): ").strip()
                    if escolha.isdigit() and 1 <= int(escolha) <= len(resultados_pesquisa):
                        request_transporte = resultados_pesquisa[int(escolha) - 1][0]
                        print(f"✔️ Selecionou a Request: {request_transporte}")
                    else:
                        print("❌ Seleção cancelada. Não haverá transporte.")
                else:
                    print("⚠️ Não foram encontradas Requests abertas para o seu utilizador.")
            except ImportError as e:
                print(f"❌ Erro de Importação: Não consegui encontrar o módulo pesquisar_request.py. Detalhe: {e}")
        
        elif req_input == "4":
            print("⏭️  Nenhuma request selecionada (Transporte ignorado).")
            request_transporte = None
        print("============================================================")

    print(f"\n📋 Roles a processar (agrupadas): {len(df_proc_agrupado)}")
    for _, rr in df_proc_agrupado.iterrows():
        print(f" - {rr['AGR_NAME']}: {rr['TEXT']} (TCODEs: {len(rr['TCODE_LIST'])})")

    if pedir_confirmacao and not modo_nao_interativo:
        if input("\nDeseja lançar esses dados no SAP? [S/N]: ").strip().upper() != "S":
            wb.close()
            return

    ###################################################################################
    # BLOCO 2: SAP GUI helpers
    ###################################################################################
    def get_statusbar():
        try:
            sbar = session.findById("wnd[0]/sbar")
            return (getattr(sbar, "MessageType", "").strip().upper(), (sbar.Text or "").strip())
        except: return ("", "")

    def try_actions(actions):
        for a in actions:
            try:
                ctrl = session.findById(a["path"])
                if a["op"] == "text":
                    ctrl.setFocus()
                    ctrl.text = a["val"]
                    return True
                elif a["op"] == "press":
                    if hasattr(ctrl, "Enabled") and not ctrl.Enabled: continue
                    ctrl.press()
                    return True
                elif a["op"] == "select":
                    ctrl.select()
                    return True
            except: continue
        return False

    def tratar_popup_modal(max_loops=10):
        for _ in range(max_loops):
            try:
                time.sleep(0.15)
                if session.ActiveWindow.Type != "GuiModalWindow": return False
                try:
                    if session.findById("wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE", False) or \
                       session.findById("wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE1", False):
                        return False 
                except:
                    pass
                candidatos = ["wnd[1]/usr/btnBUTTON_1", "wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[19]", "wnd[1]/tbar[0]/btn[11]"]
                for p in candidatos:
                    if try_actions([{"path": p, "op": "press"}]):
                        time.sleep(0.2)
                        break
                else: return True
            except: return False
        return True

    def wait_wnd1_close(timeout=3.0):
        t0 = time.time()
        while time.time() - t0 < timeout:
            if not _safe_find("wnd[1]"): return True
            time.sleep(0.1)
        return not _safe_find("wnd[1]")

    ###################################################################################
    # BLOCO 3: Page Object PFCG (Com Novo Log Estruturado)
    ###################################################################################
    class PFCGPage:
        def __init__(self, sess):
            self.sess = sess

        def open(self):
            print("  ├─ Abrindo a transação /NPFCG...")
            try:
                self.sess.findById("wnd[0]").maximize()
            except Exception:
                pass

            self.sess.findById("wnd[0]/tbar[0]/okcd").text = "/NPFCG"
            self.sess.findById("wnd[0]").sendVKey(0)
            tratar_popup_modal()
            _sleep(SLEEP_ACTION)

        def set_role_name(self, nome):
            print(f"  ├─ Inserindo o nome da Role: {nome}")
            return try_actions([
                {"path": "wnd[0]/usr/ctxtAGR_NAME_NEU", "op": "text", "val": nome},
                {"path": "wnd[0]/usr/ctxtAGR_NAME", "op": "text", "val": nome},
            ])

        def open_for_edit(self):
            print("  ├─ Tentando abrir em modo de 'Criação'...")
            if not try_actions([{"path": "wnd[0]/usr/btn%#AUTOTEXT003", "op": "press"}, {"path": "wnd[0]/tbar[1]/btn[5]", "op": "press"}]):
                raise Exception("Não consegui clicar em Criar.")
            _sleep(SLEEP_ACTION)
            tratar_popup_modal()

            mt, sb = get_statusbar()
            if "EXISTE" in norm_txt(sb) or "EXISTS" in norm_txt(sb):
                print("  ├─ A Role já existe. Alterando para modo de 'Alteração'...")
                if not try_actions([{"path": "wnd[0]/usr/btn%#AUTOTEXT001", "op": "press"}, {"path": "wnd[0]/tbar[1]/btn[2]", "op": "press"}]):
                    raise Exception("Role já existe, mas não consegui abrir Alterar.")
                _sleep(SLEEP_ACTION)
                tratar_popup_modal()
                return "CHANGE"
            return "CREATE"

        def set_description(self, desc):
            print("  ├─ Preenchendo a descrição da Role...")
            cands = ["wnd[0]/usr/txtS_AGR_TEXTS-TEXT", "wnd[0]/usr/txtS_AGR_TEXTS-TEXT1", "wnd[0]/usr/txtAGR_TEXTS-TEXT"]
            for c in cands:
                obj = _safe_find(c)
                if obj:
                    obj.text = desc
                    self.sess.findById("wnd[0]").sendVKey(0)
                    tratar_popup_modal()
                    return True
            return False

        def save(self, log_msg="  └─ Guardando alterações..."):
            print(log_msg)
            try: self.sess.findById("wnd[0]").sendVKey(11)
            except: try_actions([{"path": "wnd[0]/tbar[0]/btn[11]", "op": "press"}])
            _sleep(SLEEP_ACTION)
            tratar_popup_modal()

            mt, sb = get_statusbar()
            if sb:
                if mt in ("E", "A"):
                    print(f"     ❌ SAP Erro: {sb}")
                    raise Exception(f"Falha ao guardar: {sb}")
                else:
                    print(f"     ✔️ SAP: {sb}")
            else:
                print("     ✔️ SAP: Operação concluída sem mensagem do sistema.")

        def goto_menu_tab(self):
            print("  ├─ Acedendo à aba 'Menu' (TAB9)...")
            if not try_actions([{"path": "wnd[0]/usr/tabsTABSTRIP1/tabpTAB9", "op": "select"}]):
                raise Exception("Não consegui abrir a aba Menu (TAB9).")
            _sleep(SLEEP_ACTION)
            tratar_popup_modal()

        def _open_tcode_wizard(self):
            shells = [
                "wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubSUB1:SAPLPRGN_TREE:0321/cntlTOOL_CONTROL/shellcont/shell",
                "wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubSUB1:SAPLPRGN_TREE:0320/cntlTOOL_CONTROL/shellcont/shell"
            ]
            for s in shells:
                obj = _safe_find(s)
                if obj:
                    print("  ├─ Abrindo o Wizard de inserção de Transações (TB03)...")
                    obj.pressButton("TB03")
                    _sleep(SLEEP_ACTION)
                    return True
            raise Exception("Não encontrei o botão TB03.")

        def _fill_tcodes_fast(self, table_base, tcodes):
            if not tcodes: return 0
            inserted = 0
            for i, t in enumerate(tcodes):
                if i >= TCODE_BLOCK_SIZE: break
                try:
                    cell_id = f"{table_base}/ctxtS_TCODES-TCODE[0,{i}]"
                    cell = self.sess.findById(cell_id)
                    cell.text = t
                    inserted += 1
                except Exception:
                    continue
            _sleep(0.1) 
            return inserted

        def add_tcodes(self, tcodes):
            if not tcodes: return 0
            print(f"  ├─ Preparando a inserção rápida de {len(tcodes)} TCODE(s)...")
            inserted_total = 0
            total_blocos = max(1, ceil(len(tcodes) / TCODE_BLOCK_SIZE))

            for bloco in range(total_blocos):
                sub = tcodes[bloco * TCODE_BLOCK_SIZE : bloco * TCODE_BLOCK_SIZE + TCODE_BLOCK_SIZE]
                self._open_tcode_wizard()
                
                table_base = "wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE"
                if not _safe_find(table_base):
                    table_base = "wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE1"

                print(f"  ├─ Injetando bloco {bloco+1} na tabela...")
                qtd = self._fill_tcodes_fast(table_base, sub)
                inserted_total += qtd

                try: self.sess.findById("wnd[1]").sendVKey(0)
                except: pass
                _sleep(0.2)

                print("  ├─ Confirmando transações (Transferir)...")
                if not try_actions([{"path": "wnd[1]/tbar[0]/btn[19]", "op": "press"}, {"path": "wnd[1]/tbar[0]/btn[0]", "op": "press"}]):
                    pass 
                    
                _sleep(SLEEP_ACTION)
                wait_wnd1_close(timeout=2.0)
                tratar_popup_modal()

            return inserted_total

        def generate_authorization_profile(self):
            print("  ├─ Acedendo à aba 'Autorizações' (TAB5)...")
            if not try_actions([{"path": "wnd[0]/usr/tabsTABSTRIP1/tabpTAB5", "op": "select"}]):
                return False
            _sleep(SLEEP_ACTION)
            tratar_popup_modal()

            print("  ├─ Clicando em 'Sugerir nome de perfil'...")
            try_actions([{"path": "wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubSUB1:SAPLPRGN_TREE:0350/btnPROFIL1", "op": "press"}])
            _sleep(0.3)
            
            if _safe_find("wnd[1]"):
                print("  ├─ Confirmando a sugestão de nome de perfil no popup...")
                try_actions([{"path": "wnd[1]/tbar[0]/btn[11]", "op": "press"}])
                _sleep(0.3)

            self.save("  ├─ Guardando a Role antes de gerar as autorizações...")

            print("  ├─ Acionando a Geração de Perfil... a aguardar...")
            try_actions([{"path": "wnd[0]/tbar[1]/btn[17]", "op": "press"}])
            _sleep(1.0)

            if _safe_find("wnd[1]"):
                print("  ├─ Confirmando a geração de perfil na janela intermédia...")
                try_actions([{"path": "wnd[1]/usr/btnBUTTON_1", "op": "press"}])
                _sleep(0.5)

            if _safe_find("wnd[1]"):
                print("  └─ Fechando popup de logs de autorização...")
                try: self.sess.findById("wnd[1]").sendVKey(0)
                except: pass
                _sleep(0.5)

            try: self.sess.findById("wnd[0]").sendVKey(0)
            except: pass
            _sleep(0.3)
            tratar_popup_modal() 
            
            mt, sb = get_statusbar()
            if sb and mt not in ("E", "A"):
                print(f"     ✔️ SAP: {sb}")
            else:
                print("     ✔️ SAP: Perfil gerado e confirmado.")
            return True

        def execute_transport_and_exit(self, req_num):
            if req_num:
                print("  ├─ Recuando para a base da PFCG para pedir Transporte (F3 x2)...")
                for _ in range(2):
                    try_actions([{"path": "wnd[0]/tbar[0]/btn[3]", "op": "press"}])
                    _sleep(0.3)
                    tratar_popup_modal()

                print("  ├─ Acedendo ao Menu Função -> Transporte...")
                try_actions([{"path": "wnd[0]/mbar/menu[0]/menu[9]", "op": "select"}])
                _sleep(0.3)
                tratar_popup_modal()

                print("  ├─ Clicando em Executar transporte...")
                try_actions([{"path": "wnd[0]/tbar[1]/btn[8]", "op": "press"}])
                _sleep(0.5)

                print(f"  ├─ Injetando a Request ({req_num}) diretamente no popup...")
                for _ in range(10):
                    try:
                        req_field = self.sess.findById("wnd[1]/usr/ctxtKO008-TRKORR", False)
                        if req_field:
                            req_field.text = str(req_num)
                            break
                    except:
                        pass
                    time.sleep(0.2)
                
                try_actions([{"path": "wnd[1]/tbar[0]/btn[0]", "op": "press"}])
                _sleep(0.5)
                tratar_popup_modal()
                
                mt, sb = get_statusbar()
                if sb and mt not in ("E", "A"):
                    print(f"     ✔️ SAP: {sb}")
                else:
                    print("     ✔️ SAP: Transporte associado com sucesso!")

            print("  └─ Regressando em segurança ao ecrã principal SAP Easy Access (F3)...")
            for _ in range(3):
                try_actions([{"path": "wnd[0]/tbar[0]/btn[3]", "op": "press"}])
                _sleep(0.2)
                tratar_popup_modal()

    ###################################################################################
    # BLOCO 4: EXECUÇÃO ESTRUTURADA
    ###################################################################################
    pfcg = PFCGPage(session)
    resultados = {}

    for idx, rr in df_proc_agrupado.iterrows():
        nome, desc = (rr["AGR_NAME"] or "").strip(), (rr["TEXT"] or "").strip()
        tcodes = rr["TCODE_LIST"]
        
        print("\n======================================================================")
        print(f"▶ [{idx+1}/{len(df_proc_agrupado)}] INICIANDO ROLE: {nome} | TCODEs: {len(tcodes)}")
        print("======================================================================")
        
        tempo_inicio_role = time.time()
        
        try:
            print("\n[Etapa 1] Preparação e Dados Básicos")
            pfcg.open() 
            if not pfcg.set_role_name(nome): raise Exception("Falha ao escrever AGR_NAME.")
            modo = pfcg.open_for_edit()
            pfcg.set_description(desc)
            pfcg.save("  └─ Guardando alterações iniciais...")
            
            print("\n[Etapa 2] Atribuição de Transações (Aba Menu)")
            pfcg.goto_menu_tab()
            qtd_ins = pfcg.add_tcodes(tcodes)
            pfcg.save("  └─ Guardando Transações inseridas...")
            
            print("\n[Etapa 3] Geração do Perfil de Autorizações")
            pfcg.generate_authorization_profile()

            print("\n[Etapa 4] Ordem de Transporte e Encerramento")
            pfcg.execute_transport_and_exit(request_transporte)

            tempo_decorrido_role = time.time() - tempo_inicio_role
            str_tempo = formatar_tempo(tempo_decorrido_role)

            msg_transporte = f" | Add Req {request_transporte}" if request_transporte else ""
            resultados[nome] = {
                "STATUS": "CONCLUIDO",
                "MSG": f"Sucesso ({modo}) | {qtd_ins}/{len(tcodes)} TCODEs | Perfil Gerado{msg_transporte}.",
                "TIMESTEMP": now_ts()
            }
            print(f"\n🟢 SUCESSO: Role tratada por completo! ⏱️ (Tempo: {str_tempo})")
            print("----------------------------------------------------------------------")
            
        except Exception as e:
            tempo_decorrido_role = time.time() - tempo_inicio_role
            str_tempo = formatar_tempo(tempo_decorrido_role)
            
            err = str(e)
            mt, sb = get_statusbar()
            if mt in ("E", "A"): err = sb
            resultados[nome] = {"STATUS": "ERRO", "MSG": err, "TIMESTEMP": now_ts()}
            
            print(f"\n🔴 ERRO: {err} ⏱️ (Tempo: {str_tempo})")
            print("----------------------------------------------------------------------")
            
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                session.findById("wnd[0]").sendVKey(0)
            except: pass

    ###################################################################################
    # BLOCO 5: GRAVAR EXCEL E TEMPO TOTAL
    ###################################################################################
    try:
        col_st, col_ms, col_tm = header_map.get("STATUS"), header_map.get("MSG"), header_map.get("TIMESTEMP")
        for rec in records:
            chave_busca = str(rec["AGR_NAME"]).strip()
            res = resultados.get(chave_busca)
            if res:
                if col_st: ws.cell(row=rec["_row"], column=col_st).value = res["STATUS"]
                if col_ms: ws.cell(row=rec["_row"], column=col_ms).value = res["MSG"]
                if col_tm: ws.cell(row=rec["_row"], column=col_tm).value = res["TIMESTEMP"]
                
        wb.save(caminho_ficheiro)
        wb.close()
        print("\n💾 Resultados gravados com sucesso no Excel!")
    except Exception as e:
        print(f"\n❌ Erro a gravar Excel: {e}")

    tempo_decorrido_total = time.time() - tempo_inicio_total
    print(f"\n⏱️ Tempo total da operação: {formatar_tempo(tempo_decorrido_total)}")

    print("🔁 Fim.")
    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"])
    parser.add_argument("--xlsx")
    parser.add_argument("--request", help="Número da Request de Transporte (Opcional)")
    parser.add_argument("--auto", action="store_true")
    parser.add_argument("--no-confirm", action="store_true")
    args = parser.parse_args()

    env_cli = args.ambiente or (input("Ambiente (DEV/QAD/PRD): ").strip().upper() or "DEV")
    
    executar(
        ambiente_cockpit=env_cli, 
        caminho_ficheiro=args.xlsx, 
        request_transporte=args.request,
        modo_nao_interativo=bool(args.auto), 
        pedir_confirmacao=(not args.no_confirm)
    )