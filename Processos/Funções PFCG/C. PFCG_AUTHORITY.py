# -*- coding: utf-8 -*-

###################################################################################
# C. PFCG_AUTHORITY.py
# PFCG - Inserção Massiva de Valores de Autorização via PFCGMASSVAL & Funções Compostas
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
    from openpyxl import load_workbook

    tempo_inicio_total = time.time()

    dir_atual = os.path.dirname(os.path.abspath(__file__))
    dir_processos = os.path.dirname(dir_atual) 
    if dir_processos not in sys.path:
        sys.path.insert(0, dir_processos)

    NOME_SHEET = "PFCG_AUTHORITY"
    SEARCH_HEADER_IN_FIRST_ROWS = 20
    COLUNAS_MINIMAS = {"AGR_NAME", "STATUS", "MSG"} 

    MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}
    SISTEMA_ESPERADO = MAPA_SISTEMA.get(str(ambiente_cockpit).upper().strip() or "", None)
    
    if not SISTEMA_ESPERADO:
        raise ValueError(f"Ambiente inválido: '{ambiente_cockpit}'.")

    def formatar_tempo(segundos):
        h, resto = divmod(segundos, 3600)
        m, s = divmod(resto, 60)
        if h > 0: return f"{int(h):02d}h {int(m):02d}m {int(s):02d}s"
        return f"{int(m):02d}m {int(s):02d}s"

    def norm_col(s):
        if s is None: s = ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def norm_txt(s):
        if s is None: s = ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def now_ts():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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

    ###################################################################################
    # LER EXCEL (MAPEAMENTO DINÂMICO DE TODAS AS COLUNAS)
    ###################################################################################
    if not caminho_ficheiro:
        if modo_nao_interativo:
            raise ValueError("Faltou o parâmetro --xlsx em modo não-interativo.")
        print("📂 Selecione o ficheiro Excel…")
        caminho_ficheiro = selecionar_ficheiro()
        if not caminho_ficheiro: return

    try:
        wb = load_workbook(caminho_ficheiro)
        ws = wb[NOME_SHEET]
    except Exception as e:
        print(f"❌ Erro ao abrir Excel: {e}")
        return

    header_row = None
    col_agr_composta = None
    col_text_composta = None
    col_agr_simples = None
    col_text_simples = None
    col_objeto = None
    col_status = None
    col_msg = None
    col_timestamp = None
    dynamic_fields = {} # name -> col_idx

    for r in range(1, SEARCH_HEADER_IN_FIRST_ROWS + 1):
        row_vals = [norm_col(c.value) for c in ws[r]]
        colunas_encontradas = set(row_vals).intersection(COLUNAS_MINIMAS)
        if len(colunas_encontradas) >= len(COLUNAS_MINIMAS):
            header_row = r
            
            seen_agr_name = False
            for idx, val in enumerate(row_vals, start=1):
                if not val:
                    continue
                if val == "AGR_NAME_COMPOSTA":
                    col_agr_composta = idx
                elif val == "AGR_NAME":
                    col_agr_simples = idx
                    seen_agr_name = True
                elif val == "TEXT":
                    if not seen_agr_name:
                        col_text_composta = idx
                    else:
                        col_text_simples = idx
                elif val in ("OBJETO DE AUTORIZACAO", "OBJETO DE AUTORIZAÇÃO", "OBJETO"):
                    col_objeto = idx
                elif val == "STATUS":
                    col_status = idx
                elif val == "MSG":
                    col_msg = idx
                elif val in ("TIMESTEMP", "TIMESTAMP"):
                    col_timestamp = idx
                elif val not in ("ID",):
                    dynamic_fields[val] = idx
            break

    if not header_row:
        print("\n❌ Cabeçalho não encontrado.")
        wb.close()
        return

    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        agr_val = ws.cell(row=r, column=col_agr_simples).value if col_agr_simples else None
        agr = "" if agr_val is None else str(agr_val).strip()
        if not agr: continue
        
        rec = {"_row": r}
        rec["AGR_NAME"] = agr
        
        if col_agr_composta:
            val = ws.cell(row=r, column=col_agr_composta).value
            rec["AGR_NAME_COMPOSTA"] = "" if val is None else str(val).strip()
        else:
            rec["AGR_NAME_COMPOSTA"] = ""
            
        if col_text_composta:
            val = ws.cell(row=r, column=col_text_composta).value
            rec["TEXT_COMPOSTA"] = "" if val is None else str(val).strip()
        else:
            rec["TEXT_COMPOSTA"] = ""
            
        if col_text_simples:
            val = ws.cell(row=r, column=col_text_simples).value
            rec["TEXT"] = "" if val is None else str(val).strip()
        else:
            rec["TEXT"] = ""
            
        if col_objeto:
            val = ws.cell(row=r, column=col_objeto).value
            rec["OBJETO DE AUTORIZACAO"] = "" if val is None else str(val).strip()
        else:
            rec["OBJETO DE AUTORIZACAO"] = ""
            
        if col_status:
            val = ws.cell(row=r, column=col_status).value
            rec["STATUS"] = "" if val is None else str(val).strip()
        else:
            rec["STATUS"] = ""
            
        if col_msg:
            val = ws.cell(row=r, column=col_msg).value
            rec["MSG"] = "" if val is None else str(val).strip()
        else:
            rec["MSG"] = ""
            
        if col_timestamp:
            val = ws.cell(row=r, column=col_timestamp).value
            rec["TIMESTEMP"] = "" if val is None else str(val).strip()
        else:
            rec["TIMESTEMP"] = ""
            
        # Add dynamic fields
        for f_name, f_idx in dynamic_fields.items():
            val = ws.cell(row=r, column=f_idx).value
            rec[f_name] = "" if val is None else str(val).strip()
            
        records.append(rec)

    df_original = pd.DataFrame(records)
    df_proc = df_original[~df_original["STATUS"].str.upper().str.contains("CONCLUIDO", na=False)].copy()
    if df_proc.empty:
        print("⚠️ Tudo concluído.")
        wb.close()
        return

    ###################################################################################
    # CONEXÃO SAP E MENU PREMIUM DE TRANSPORTE
    ###################################################################################
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = next((sess for conn in application.Children for sess in conn.Children if sess.Info.SystemName.upper() == SISTEMA_ESPERADO), None)
    except: session = None

    if not session:
        wb.close()
        print(f"❌ Não encontrei sessão do ambiente '{ambiente_cockpit}'.")
        return

    def object_exists(id_string):
        try:
            session.findById(id_string)
            return True
        except: return False

    def _safe_find(sap_id):
        try: return session.findById(sap_id)
        except: return None

    def _sap_busy():
        try: return bool(getattr(session, "Busy", False))
        except: return False

    def _esperar_sap_livre(timeout=8.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not _sap_busy(): return True
            time.sleep(pausa)
        return False

    def _esperar_objeto(sap_id, timeout=5.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            obj = _safe_find(sap_id)
            if obj: return obj
            time.sleep(pausa)
        return None

    def _esperar_sumir(sap_id, timeout=5.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not _safe_find(sap_id): return True
            time.sleep(pausa)
        return False

    def _send_vkey(vkey, wait_after=True):
        session.findById("wnd[0]").sendVKey(vkey)
        if wait_after: _esperar_sap_livre()

    def _press_if_exists(sap_id, timeout=2.0):
        obj = _esperar_objeto(sap_id, timeout=timeout)
        if not obj: return False
        try:
            obj.press()
            _esperar_sap_livre()
            return True
        except: return False

    sap_id_cache = {}
    def _resolver_id(cache_key, candidatos):
        sap_id = sap_id_cache.get(cache_key)
        if sap_id:
            obj = _safe_find(sap_id)
            if obj: return sap_id, obj
            sap_id_cache.pop(cache_key, None)
        for sap_id in candidatos:
            obj = _safe_find(sap_id)
            if obj:
                sap_id_cache[cache_key] = sap_id
                return sap_id, obj
        return None, None

    def _resolver_id_esperando(cache_key, candidatos, timeout=3.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            sap_id, obj = _resolver_id(cache_key, candidatos)
            if obj: return sap_id, obj
            time.sleep(pausa)
        return None, None

    def get_statusbar():
        try:
            sbar = session.findById("wnd[0]/sbar")
            tipo = getattr(sbar, "MessageType", "").strip().upper()
            texto = (sbar.Text or "").strip()
            if texto: print(f"[SAP_SBAR] {texto}")
            return (tipo, texto)
        except: return ("", "")

    def try_actions(actions):
        for a in actions:
            try:
                ctrl = session.findById(a["path"])
                if a["op"] == "text":
                    ctrl.setFocus()
                    ctrl.text = a["val"]
                    time.sleep(0.1)
                    return True
                elif a["op"] == "press":
                    if hasattr(ctrl, "Enabled") and not ctrl.Enabled: continue
                    ctrl.press()
                    time.sleep(0.15)
                    return True
                elif a["op"] == "select":
                    ctrl.select()
                    time.sleep(0.1)
                    return True
            except: continue
        return False

    def tratar_popup_modal(max_loops=6):
        for _ in range(max_loops):
            try:
                time.sleep(0.1)
                if session.ActiveWindow.Type != "GuiModalWindow": return False
                candidatos = [
                    "wnd[1]/usr/btnBUTTON_1",
                    "wnd[1]/usr/btnSPOP-OPTION1",
                    "wnd[1]/tbar[0]/btn[0]",
                    "wnd[1]/tbar[0]/btn[19]",
                    "wnd[1]/tbar[0]/btn[11]"
                ]
                for p in candidatos:
                    if try_actions([{"path": p, "op": "press"}]): return True
                return True
            except: return False
        return True

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

    print("\n📋 Resumo das Funções a processar:")
    for idx, r in df_proc.iterrows():
        r_nome = r["AGR_NAME"]
        r_obj = r["OBJETO DE AUTORIZACAO"] or "F_KNA1_GRP"
        comp_info = f" -> Composta: {r['AGR_NAME_COMPOSTA']}" if r['AGR_NAME_COMPOSTA'] else ""
        print(f"   - {r_nome} (Obj: {r_obj}){comp_info}")
        
    print(f"\n🔢 Total de linhas para processamento: {len(df_proc)}")

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

    if pedir_confirmacao and not modo_nao_interativo:
        if input("\nDeseja iniciar o processamento no SAP? [S/N]: ").strip().upper() != "S":
            wb.close()
            print("❌ Operação cancelada pelo utilizador.")
            return

    ###################################################################################
    # LOGGER DE AUDITORIA & EXECUTOR
    ###################################################################################
    class PFCG_AuthPage_Auditor:
        def __init__(self, sess):
            self.sess = sess

        def get_sbar(self):
            try:
                sbar = self.sess.findById("wnd[0]/sbar")
                return getattr(sbar, "MessageType", "").strip().upper(), (sbar.Text or "").strip()
            except: return "", ""

        def audit_step(self, descricao, path, acao="press", valor=None, vkey=None, silencioso=False):
            if not silencioso:
                print(f"\n  🔎 [AUDIT] {descricao}")
                log_detail = f"      ↳ ID: {path} | Ação: {acao}"
                if valor is not None: log_detail += f" | Valor: '{valor}'"
                if vkey is not None: log_detail += f" | VKey: {vkey}"
                print(log_detail)

            try:
                if path: elem = self.sess.findById(path)

                if acao == "text": elem.text = valor
                elif acao == "press": elem.press()
                elif acao == "select": 
                    if hasattr(elem, "selected"): elem.selected = True
                    else: elem.select()
                elif acao == "sendVKey":
                    if path: elem.sendVKey(vkey)
                    else: self.sess.findById("wnd[0]").sendVKey(vkey)

                if not silencioso:
                    print("      ✅ SUCESSO")
                    mtype, mtext = self.get_sbar()
                    if mtext:
                        icone = "🔴" if mtype in ["E", "A"] else ("🟡" if mtype == "W" else "🟢")
                        print(f"      {icone} SAP STATUS: [{mtype}] {mtext}")
                
            except Exception as e:
                if not silencioso:
                    print(f"      ❌ FALHA AQUI (Erro 619): O ID [{path}] falhou.")
                    raise Exception(f"FALHA NO PASSO: '{descricao}' -> ID: {path}")

        def ensure_role_exists(self, nome, desc):
            self.audit_step("Chamar transação /npfcg", "wnd[0]/tbar[0]/okcd", "text", "/npfcg")
            self.audit_step("Enter (Ir para PFCG)", "wnd[0]", "sendVKey", vkey=0)
            
            self.audit_step("Inserir Nome da Role", "wnd[0]/usr/ctxtAGR_NAME_NEU", "text", nome)
            self.audit_step("Clicar Criar Papel Simples", "wnd[0]/usr/btn%#AUTOTEXT003", "press")
            
            if object_exists("wnd[0]/usr/txtS_AGR_TEXTS-TEXT"):
                descricao_final = desc if desc else f"Criado via Script - {nome}"
                self.audit_step("Preencher Descrição da Role", "wnd[0]/usr/txtS_AGR_TEXTS-TEXT", "text", descricao_final)
                self.audit_step("Enter após descrição", "wnd[0]", "sendVKey", vkey=0)
                
                if self.sess.Children.Count > 1 and object_exists("wnd[1]/usr/btnBUTTON_1"):
                    self.audit_step("Confirmar popup criação (Sim)", "wnd[1]/usr/btnBUTTON_1", "press")
                
                if object_exists("wnd[0]/tbar[0]/btn[11]"):
                    self.audit_step("Guardar Role", "wnd[0]/tbar[0]/btn[11]", "press")
            else:
                print("  ├─ [SAP] A função já existe.")

        def update_mass_values_dynamic(self, nome, objeto, row_data):
            self.audit_step("Chamar transação /nPFCGMASSVAL", "wnd[0]/tbar[0]/okcd", "text", "/nPFCGMASSVAL")
            self.audit_step("Enter (Ir para MASSVAL)", "wnd[0]", "sendVKey", vkey=0)
            
            self.audit_step("Selecionar Execução Direta", "wnd[0]/usr/radMOD_EXE", "select")
            self.audit_step("Selecionar Inserção Manual", "wnd[0]/usr/radSEL_NAU", "select")
            self.audit_step("ENTER Crucial do VBS para atualizar tela", "wnd[0]", "sendVKey", vkey=0)
            
            self.audit_step("Preencher ROLE-LOW por segurança", "wnd[0]/usr/ctxtROLE-LOW", "text", nome)
            
            self.audit_step("Preencher OBJOBJ (Objeto de Autorização)", "wnd[0]/usr/ctxtOBJOBJ", "text", objeto)
            self.audit_step("Enter após OBJOBJ (A carregar campos...)", "wnd[0]", "sendVKey", vkey=0)
            
            time.sleep(0.5)

            # ---------------------------------------------------------
            # LEITURA DINÂMICA SILENCIOSA
            # ---------------------------------------------------------
            campos_encontrados_sap = 0
            
            for j in range(1, 15):
                btn_id = f"wnd[0]/usr/btnPOBJ{j}N"
                if not object_exists(btn_id):
                    break 
                
                campos_encontrados_sap += 1
                
                campo_sap_tecnico = ""
                ids_para_testar = [
                    f"wnd[0]/usr/txtOBJFLD{j}",      
                    f"wnd[0]/usr/ctxtOBJFLD{j}",     
                    f"wnd[0]/usr/ctxtS_AGR_DEFINE-FNAM{j}",
                    f"wnd[0]/usr/txtS_AGR_DEFINE-FNAM{j}",
                    f"wnd[0]/usr/ctxtFNAM{j}",
                    f"wnd[0]/usr/txtFNAM{j}",
                    f"wnd[0]/usr/ctxtFIELD{j}",
                    f"wnd[0]/usr/txtFIELD{j}"
                ]
                
                for cid in ids_para_testar:
                    if object_exists(cid):
                        campo_sap_tecnico = self.sess.findById(cid).text.strip().upper()
                        break

                if not campo_sap_tecnico:
                    campo_sap_tecnico = str(row_data.get(f'CAMPO {j}', '')).strip().upper()

                if not campo_sap_tecnico:
                    continue 

                valor_no_excel = str(row_data.get(campo_sap_tecnico, '')).strip()

                if valor_no_excel and valor_no_excel != 'NAN':
                    print(f"\n  ⭐ [CAMPO IDENTIFICADO] Nome: '{campo_sap_tecnico}' | Valor a inserir: '{valor_no_excel}'")
                    
                    self.audit_step(f"Clicar Botão VALS para '{campo_sap_tecnico}'", btn_id, "press")
                    time.sleep(0.5) 
                    
                    if valor_no_excel == "*":
                        self.audit_step(f"Full Auth (*) no Campo '{campo_sap_tecnico}'", "wnd[1]/usr/btnGES2", "press")
                    else:
                        valores_lista = [v.strip() for v in valor_no_excel.split(',')]
                        
                        if object_exists("wnd[1]/usr/tblSAPLSUPRNACT_TC"):
                            print(f"      ℹ️ Janela de Atividades Detetada! Valores: {valores_lista}")
                            for linha in range(15):
                                try:
                                    act_code = self.sess.findById(f"wnd[1]/usr/tblSAPLSUPRNACT_TC/txtH_FVAL-LOW[1,{linha}]").text.strip()
                                    if act_code in valores_lista:
                                        self.audit_step(f"Marcar Checkbox '{act_code}'", f"wnd[1]/usr/tblSAPLSUPRNACT_TC/chkH_FVAL-MARK[0,{linha}]", "select", silencioso=True)
                                        print(f"      ✅ Checkbox '{act_code}' marcada com sucesso.")
                                except: break 
                        else:
                            for idx_val, val in enumerate(valores_lista):
                                try:
                                    self.audit_step(f"Preencher Valor '{val}' para '{campo_sap_tecnico}'", f"wnd[1]/usr/tblSAPLSUPRNVAL_TC/ctxtH_FVAL_LOW[0,{idx_val}]", "text", val, silencioso=True)
                                except Exception:
                                    self.audit_step(f"Preencher Valor '{val}' para '{campo_sap_tecnico}'", f"wnd[1]/usr/tblSAPLSUPRNVAL_TC/ctxtH_FVAL_LOW[1,{idx_val}]", "text", val, silencioso=True)
                                print(f"      ✅ Valor '{val}' inserido com sucesso.")

                    self.audit_step(f"Confirmar Popup de '{campo_sap_tecnico}'", "wnd[1]/tbar[0]/btn[0]", "press")

            if campos_encontrados_sap == 0:
                print("  ⚠️ AVISO: O Objeto inserido não gerou campos visíveis.")

            self.audit_step("Clicar Executar (Relógio)", "wnd[0]/tbar[1]/btn[8]", "press")
            self.audit_step("Clicar Guardar (Disquete)", "wnd[0]/tbar[1]/btn[20]", "press")
            
            if self.sess.Children.Count > 1 and object_exists("wnd[1]/tbar[0]/btn[0]"):
                self.audit_step("Confirmar popup Sucesso Gravação", "wnd[1]/tbar[0]/btn[0]", "press")

        def execute_transport(self, req_num, nome):
            self.audit_step("Chamar transação /nPFCG para Transporte", "wnd[0]/tbar[0]/okcd", "text", "/nPFCG")
            self.audit_step("Enter", "wnd[0]", "sendVKey", vkey=0)
            
            self.audit_step("Preencher Função", "wnd[0]/usr/ctxtAGR_NAME_NEU", "text", nome)
            self.audit_step("Selecionar Menu Transporte", "wnd[0]/mbar/menu[0]/menu[9]", "select")
            self.audit_step("Executar Transporte", "wnd[0]/tbar[1]/btn[8]", "press")
            time.sleep(0.3)
            
            if self.sess.Children.Count > 1 and object_exists("wnd[1]/usr/ctxtKO008-TRKORR"):
                self.audit_step("Inserir Request", "wnd[1]/usr/ctxtKO008-TRKORR", "text", req_num)
                self.audit_step("Confirmar Request", "wnd[1]/tbar[0]/btn[0]", "press")

        # =============================================================================
        # COMPOSITE ROLES METHODS
        # =============================================================================
        def set_role_name(self, nome):
            sap_id, obj = _resolver_id(
                "role_name_field",
                ["wnd[0]/usr/ctxtAGR_NAME_NEU", "wnd[0]/usr/ctxtAGR_NAME"]
            )
            if not obj: return False
            try:
                obj.setFocus()
                obj.text = nome
                _esperar_sap_livre()
                return True
            except: return False

        def open_for_edit(self):
            if not try_actions([
                {"path": "wnd[0]/usr/btn%#AUTOTEXT004", "op": "press"},
                {"path": "wnd[0]/tbar[1]/btn[5]", "op": "press"}
            ]):
                raise Exception("Não consegui clicar em Criar Função Composta.")

            tratar_popup_modal()

            mt, sb = get_statusbar()
            if "EXISTE" in norm_txt(sb) or "EXISTS" in norm_txt(sb):
                print("  ├─ A Role já existe. Alterando para modo de 'Alteração'...")
                if not try_actions([
                    {"path": "wnd[0]/usr/btn%#AUTOTEXT001", "op": "press"},
                    {"path": "wnd[0]/tbar[1]/btn[2]", "op": "press"}
                ]):
                    raise Exception("Role já existe, mas não consegui abrir Alterar.")
                tratar_popup_modal()
                return "CHANGE"
            return "CREATE"

        def set_description(self, desc):
            sap_id, obj = _resolver_id(
                "role_desc_field",
                ["wnd[0]/usr/txtS_AGR_TEXTS-TEXT", "wnd[0]/usr/txtS_AGR_TEXTS-TEXT1", "wnd[0]/usr/txtAGR_TEXTS-TEXT"]
            )
            if not obj: return False
            try:
                obj.text = desc
                _send_vkey(0)
                tratar_popup_modal()
                return True
            except: return False

        def save_composite(self, log_msg="  └─ Guardando alterações..."):
            print(log_msg)
            try:
                self.sess.findById("wnd[0]").sendVKey(11)
            except:
                try_actions([{"path": "wnd[0]/tbar[0]/btn[11]", "op": "press"}])

            _esperar_sap_livre()
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

        def ensure_composite_role_exists(self, nome, desc):
            self.audit_step("Chamar transação /npfcg para Composta", "wnd[0]/tbar[0]/okcd", "text", "/npfcg")
            self.audit_step("Enter (Ir para PFCG)", "wnd[0]", "sendVKey", vkey=0)
            
            if not self.set_role_name(nome):
                raise Exception("Falha ao escrever nome da Função Composta.")
                
            modo = self.open_for_edit()
            self.set_description(desc)
            self.save_composite("  └─ Guardando alterações iniciais da Composta...")
            return modo

        def goto_roles_tab(self):
            sap_id, obj = _resolver_id("roles_tab", ["wnd[0]/usr/tabsTABSTRIP1/tabpTAB8"])
            if not obj:
                raise Exception("Não consegui abrir a aba Funções (TAB8).")
            try:
                obj.select()
                _esperar_sap_livre()
            except:
                raise Exception("Não consegui abrir a aba Funções (TAB8).")
            tratar_popup_modal()

        def add_roles(self, roles_list):
            if not roles_list: return 0

            print(f"  ├─ A preparar inserção de {len(roles_list)} função(ões) componente(s)...")
            
            table_id, table_obj = _resolver_id(
                "roles_table",
                [
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubSUB1:SAPLPRGN_TREE:0600/tblSAPLPRGN_TREECTRL_AGRLIST2",
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubSUB1:SAPLPRGN_TREE:0610/tblSAPLPRGN_TREECTRL_AGRLIST2",
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubSUB1:SAPLPRGN_TREE:0620/tblSAPLPRGN_TREECTRL_AGRLIST2",
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubSUB1:SAPLPRGN_TREE:0330/tblSAPLPRGN_TREECTRL_AGRLIST2",
                ]
            )
            if not table_obj:
                raise Exception("Não encontrei a tabela de funções componentes.")

            visible_rows = 10
            try:
                visible_rows = int(table_obj.VisibleRowCount)
            except: pass

            inserted = 0
            for idx, role in enumerate(roles_list):
                row_in_page = idx % visible_rows
                
                if idx > 0 and row_in_page == 0:
                    try:
                        table_obj.VerticalScrollbar.Position = idx
                        _esperar_sap_livre()
                    except:
                        try:
                            self.sess.findById("wnd[0]").sendVKey(0)
                            _esperar_sap_livre()
                        except: pass
                
                cell_id = f"{table_id}/ctxtI_ACTGROUPS-AGR_NAME[0,{row_in_page}]"
                cell = _esperar_objeto(cell_id, timeout=2.0)
                if not cell:
                    cell = _safe_find(cell_id)
                
                if cell:
                    cell.text = role
                    inserted += 1
                    print(f"     ├─ Inserindo {role} na linha {idx}")
                else:
                    print(f"     ⚠️ Não consegui encontrar o campo para a linha {idx} (ID: {cell_id})")

            _esperar_sap_livre()
            try: self.sess.findById("wnd[0]").sendVKey(0)
            except: pass
            _esperar_sap_livre()
            tratar_popup_modal()
            return inserted

        def execute_transport_composite(self, req_num, nome):
            if req_num:
                print("  ├─ Recuando para a base da PFCG para pedir Transporte (F3)...")
                try_actions([{"path": "wnd[0]/tbar[0]/btn[3]", "op": "press"}])
                tratar_popup_modal()

                print("  ├─ Acedendo ao Menu Função -> Transporte...")
                try_actions([{"path": "wnd[0]/mbar/menu[0]/menu[9]", "op": "select"}])
                tratar_popup_modal()

                print("  ├─ Clicando em Executar transporte...")
                try_actions([{"path": "wnd[0]/tbar[1]/btn[8]", "op": "press"}])

                field_id, req_field = _resolver_id_esperando(
                    "transport_req_field",
                    ["wnd[1]/usr/ctxtKO008-TRKORR"],
                    timeout=3.0
                )

                print(f"  ├─ Injetando a Request ({req_num}) diretamente no popup...")
                if req_field:
                    req_field.text = str(req_num)

                try_actions([{"path": "wnd[1]/tbar[0]/btn[0]", "op": "press"}])
                tratar_popup_modal()

                mt, sb = get_statusbar()
                if sb and mt not in ("E", "A"):
                    print(f"     ✔️ SAP: {sb}")
                else:
                    print("     ✔️ SAP: Transporte associado com sucesso!")

            print("  └─ Regressando em segurança ao ecrã principal SAP Easy Access (F3)...")
            for _ in range(3):
                try_actions([{"path": "wnd[0]/tbar[0]/btn[3]", "op": "press"}])
                tratar_popup_modal()


    ###################################################################################
    # EXECUÇÃO SEQUENCIAL (FASE 1: SIMPLES, FASE 2: COMPOSTAS)
    ###################################################################################
    auditor = PFCG_AuthPage_Auditor(session)
    resultados_simples = {}
    resultados = {}

    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        session.findById("wnd[0]").sendVKey(0)
    except: pass

    # FASE 1: Funções Simples
    for idx, rr in df_proc.iterrows():
        nome = str(rr["AGR_NAME"]).strip()
        desc = str(rr.get("TEXT", "")).strip()
        
        objeto = str(rr.get("OBJETO DE AUTORIZACAO", rr.get("OBJETO", ""))).strip()
        if not objeto: objeto = "F_KNA1_GRP"
        
        print("\n======================================================================")
        print(f"▶ [{idx+1}/{len(df_proc)}] [FASE SIMPLES] LÓGICA DINÂMICA: {nome} | OBJ: {objeto}")
        print("======================================================================")
        
        try:
            print("\n[Etapa 1] Validação da Role")
            auditor.ensure_role_exists(nome, desc)
            
            print("\n[Etapa 2] Inserção MASSVAL (Modo Pull Dinâmico)")
            auditor.update_mass_values_dynamic(nome, objeto, rr)
            
            if request_transporte:
                print("\n[Etapa 3] Inserir na Request de Transporte")
                auditor.execute_transport(request_transporte, nome)

            auditor.audit_step("Voltar ao Ecrã Inicial (/N) para próxima iteração", "wnd[0]/tbar[0]/okcd", "text", "/N", silencioso=True)
            auditor.audit_step("Enter (/N)", "wnd[0]", "sendVKey", vkey=0, silencioso=True)
            
            resultados_simples[idx] = {"STATUS": "SIMPLES_OK", "MSG": "Sucesso"}
            print("\n🟢 SUCESSO! Função simples processada e ecrã preparado para a próxima.")
            
        except Exception as e:
            resultados_simples[idx] = {"STATUS": "ERRO", "MSG": str(e)}
            print(f"\n🔴 INTERRUPÇÃO DETETADA NA FUNÇÃO SIMPLES:\n{str(e)}")
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                session.findById("wnd[0]").sendVKey(0)
            except: pass

    # FASE 2: Funções Compostas
    if col_agr_composta:
        compostas_a_processar = df_proc[df_proc["AGR_NAME_COMPOSTA"].str.strip() != ""]["AGR_NAME_COMPOSTA"].unique()
        
        if len(compostas_a_processar) > 0:
            print("\n======================================================================")
            print(f"▶ INICIANDO FASE COMPOSTA: {len(compostas_a_processar)} Funções Compostas a processar")
            print("======================================================================")
            
            for nome_comp in compostas_a_processar:
                nome_comp = str(nome_comp).strip()
                
                # Reunir todas as simples associadas à composta no Excel completo (df_original)
                roles_filhas_series = df_original[df_original["AGR_NAME_COMPOSTA"] == nome_comp]["AGR_NAME"]
                roles_filhas = list(dict.fromkeys([str(r).strip().upper() for r in roles_filhas_series if str(r).strip()]))
                
                # Descrição da composta
                desc_comp_series = df_original[df_original["AGR_NAME_COMPOSTA"] == nome_comp]["TEXT_COMPOSTA"]
                desc_comp = ""
                for d in desc_comp_series:
                    if str(d).strip():
                        desc_comp = str(d).strip()
                        break
                        
                print(f"\n▶ A processar Função Composta: {nome_comp} | Componentes: {len(roles_filhas)}")
                try:
                    modo = auditor.ensure_composite_role_exists(nome_comp, desc_comp)
                    auditor.goto_roles_tab()
                    qtd_ins = auditor.add_roles(roles_filhas)
                    auditor.save_composite("  └─ Guardando Funções inseridas na Composta...")
                    
                    if request_transporte:
                        print("\n[Etapa 3] Ordem de Transporte para Composta")
                        auditor.execute_transport_composite(request_transporte, nome_comp)
                        
                    print(f"🟢 SUCESSO: Função Composta {nome_comp} tratada com sucesso!")
                    
                    # Atualizar cada linha do df_proc correspondente
                    rows_matching = df_proc[df_proc["AGR_NAME_COMPOSTA"] == nome_comp].index
                    for row_idx in rows_matching:
                        res_simp = resultados_simples.get(row_idx, {"STATUS": "ERRO", "MSG": "Não processado"})
                        if res_simp["STATUS"] == "SIMPLES_OK":
                            resultados[row_idx] = {
                                "STATUS": "CONCLUIDO",
                                "MSG": f"Sucesso ({modo}) | {qtd_ins}/{len(roles_filhas)} Componentes atribuídos."
                            }
                        else:
                            resultados[row_idx] = res_simp
                            
                except Exception as e:
                    print(f"🔴 ERRO na Função Composta {nome_comp}: {e}")
                    try:
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                        session.findById("wnd[0]").sendVKey(0)
                    except: pass
                    
                    # Atualizar erros na composta nas linhas correspondentes
                    rows_matching = df_proc[df_proc["AGR_NAME_COMPOSTA"] == nome_comp].index
                    for row_idx in rows_matching:
                        resultados[row_idx] = {"STATUS": "ERRO", "MSG": f"Erro na Composta: {str(e)}"}

    # Resolver linhas sem composta ou que não entraram na Fase 2
    for idx, rr in df_proc.iterrows():
        if idx not in resultados:
            res_simp = resultados_simples.get(idx, {"STATUS": "ERRO", "MSG": "Não processado"})
            if res_simp["STATUS"] == "SIMPLES_OK":
                resultados[idx] = {"STATUS": "CONCLUIDO", "MSG": "Sucesso simples"}
            else:
                resultados[idx] = res_simp

    ###################################################################################
    # GRAVAR EXCEL E FINALIZAR
    ###################################################################################
    try:
        col_st = col_status
        col_ms = col_msg
        col_tm = col_timestamp
        for rec in records:
            df_idx = df_proc.index[df_proc['_row'] == rec['_row']].tolist()
            if not df_idx: continue
            res = resultados.get(df_idx[0])
            if res:
                if col_st: ws.cell(row=rec["_row"], column=col_st).value = res["STATUS"]
                if col_ms: ws.cell(row=rec["_row"], column=col_ms).value = res["MSG"]
                if col_tm: ws.cell(row=rec["_row"], column=col_tm).value = now_ts()
        wb.save(caminho_ficheiro)
        wb.close()
        print("\n💾 Resultados gravados no Excel.")
    except Exception as e:
        print(f"❌ Erro ao gravar no Excel: {e}")

    tempo_decorrido_total = time.time() - tempo_inicio_total
    print(f"\n⏱️ Tempo total da operação: {formatar_tempo(tempo_decorrido_total)}")

    return True

if __name__ == "__main__":
    executar("DEV")