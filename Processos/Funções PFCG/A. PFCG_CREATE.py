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
#  - Barra de progresso por Role
#  - Etapa 1 de performance: esperas inteligentes
#  - Etapa 2 de performance: sem pandas
#  - Etapa 3 de performance: cache de IDs SAP
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

    import win32com.client
    from tkinter import filedialog
    from datetime import datetime
    from math import ceil
    from openpyxl import load_workbook
    from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn

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

    SLEEP_UI = 0.08
    SLEEP_ACTION = 0.15
    TCODE_BLOCK_SIZE = 20

    def now_ts():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _sleep(t=SLEEP_UI):
        time.sleep(t)

    def norm_col(s):
        if s is None:
            s = ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def norm_txt(s):
        if s is None:
            s = ""
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
        if not raw:
            return []
        s = str(raw).replace("\r", "\n").replace("\t", " ").strip().upper()
        parts = re.split(r"[;, \n]+", s)
        out = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            if p.startswith("/N") or p.startswith("/O"):
                p = p[2:].strip()
            if p:
                out.append(p)
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
                if name:
                    header_map[name] = idx
            break

    if not header_row:
        wb.close()
        print("\n❌ Não encontrei a linha de cabeçalho completa.")
        return

    col_agr = header_map.get("AGR_NAME")
    col_text = header_map.get("TEXT")
    col_tcode = header_map.get("TCODE")
    col_status = header_map.get("STATUS")
    col_msg = header_map.get("MSG")
    col_ts = header_map.get("TIMESTEMP")

    records = []

    for r in range(header_row + 1, ws.max_row + 1):
        agr_val = ws.cell(row=r, column=col_agr).value if col_agr else None
        agr = "" if agr_val is None else str(agr_val).strip()
        if not agr:
            continue

        text_val = ws.cell(row=r, column=col_text).value if col_text else None
        tcode_val = ws.cell(row=r, column=col_tcode).value if col_tcode else None
        status_val = ws.cell(row=r, column=col_status).value if col_status else None
        msg_val = ws.cell(row=r, column=col_msg).value if col_msg else None
        ts_val = ws.cell(row=r, column=col_ts).value if col_ts else None

        records.append({
            "_row": r,
            "AGR_NAME": agr,
            "TEXT": "" if text_val is None else str(text_val).strip(),
            "TCODE": "" if tcode_val is None else str(tcode_val).strip(),
            "STATUS": "" if status_val is None else str(status_val).strip(),
            "MSG": "" if msg_val is None else str(msg_val).strip(),
            "TIMESTEMP": "" if ts_val is None else str(ts_val).strip(),
        })

    if not records:
        wb.close()
        print("⚠️ Não encontrei linhas para processar.")
        return

    roles_map = {}

    for rec in records:
        status_norm = norm_txt(rec["STATUS"])
        if status_norm == "CONCLUIDO":
            continue

        agr = rec["AGR_NAME"].strip()
        if not agr:
            continue

        if agr not in roles_map:
            roles_map[agr] = {
                "AGR_NAME": agr,
                "TEXT": rec["TEXT"].strip(),
                "TCODE_LIST": []
            }

        if not roles_map[agr]["TEXT"] and rec["TEXT"].strip():
            roles_map[agr]["TEXT"] = rec["TEXT"].strip()

        roles_map[agr]["TCODE_LIST"].extend(split_tcodes(rec["TCODE"]))

    if not roles_map:
        wb.close()
        print("⚠️ Nada para processar (tudo CONCLUIDO).")
        return

    roles_agrupadas = []
    for item in roles_map.values():
        item["TCODE_LIST"] = list(dict.fromkeys(item["TCODE_LIST"]))
        roles_agrupadas.append(item)

    roles_agrupadas.sort(key=lambda x: x["AGR_NAME"])

    # =================================================================================
    # CAPTURA SESSÃO SAP
    # =================================================================================
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = next((sess for conn in application.Children for sess in conn.Children if sess.Info.SystemName.upper() == SISTEMA_ESPERADO), None)
    except Exception:
        session = None

    if not session:
        wb.close()
        print(f"❌ Não encontrei sessão do ambiente '{ambiente_cockpit}'.")
        return

    ###################################################################################
    # HELPERS SAP - PERFORMANCE + CACHE DE IDs
    ###################################################################################
    sap_id_cache = {}

    def _safe_find(sap_id):
        try:
            return session.findById(sap_id)
        except:
            return None

    def _sap_busy():
        try:
            return bool(getattr(session, "Busy", False))
        except:
            return False

    def _esperar_sap_livre(timeout=8.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not _sap_busy():
                return True
            time.sleep(pausa)
        return False

    def _esperar_objeto(sap_id, timeout=5.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            obj = _safe_find(sap_id)
            if obj:
                return obj
            time.sleep(pausa)
        return None

    def _esperar_sumir(sap_id, timeout=5.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not _safe_find(sap_id):
                return True
            time.sleep(pausa)
        return False

    def _send_vkey(vkey, wait_after=True):
        session.findById("wnd[0]").sendVKey(vkey)
        if wait_after:
            _esperar_sap_livre()

    def _press_if_exists(sap_id, timeout=2.0):
        obj = _esperar_objeto(sap_id, timeout=timeout)
        if not obj:
            return False
        try:
            obj.press()
            _esperar_sap_livre()
            return True
        except:
            return False

    def _resolver_id(cache_key, candidatos):
        sap_id = sap_id_cache.get(cache_key)
        if sap_id:
            obj = _safe_find(sap_id)
            if obj:
                return sap_id, obj
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
            if obj:
                return sap_id, obj
            time.sleep(pausa)
        return None, None

    def _criar_nova_request_no_sap_local(sess):
        okcd = _safe_find("wnd[0]/tbar[0]/okcd")
        if okcd:
            okcd.text = "/nSE10"
            _send_vkey(0)

        print("\nTipo da ordem:")
        print("1 - Ordem customizing")
        print("2 - Ordem workbench")

        while True:
            tipo = input("Digite a opção (1/2): ").strip()
            if tipo in ("1", "2"):
                break
            print("❌ Opção inválida. Use apenas 1 ou 2.")

        desc = input("Descrição da request (máx 60): ").strip()
        desc = desc[:60] if desc else "REQUEST CRIADA VIA SCRIPT"

        sess.findById("wnd[0]/tbar[1]/btn[6]").press()
        _esperar_objeto("wnd[1]", timeout=3.0)

        if tipo == "2":
            try:
                radio = _safe_find("wnd[1]/usr/radKO042-REQ_CONS_K")
                if radio:
                    radio.select()
            except:
                pass

        _press_if_exists("wnd[1]/tbar[0]/btn[0]", timeout=3.0)

        try:
            campo_desc = _esperar_objeto("wnd[1]/usr/txtKO013-AS4TEXT", timeout=3.0)
            if campo_desc:
                campo_desc.text = desc
        except:
            pass

        _press_if_exists("wnd[1]/tbar[0]/btn[0]", timeout=3.0)

        trkorr = None
        for sap_id in ["wnd[0]/usr/lbl[20,9]", "wnd[0]/usr/lbl[1,1]"]:
            try:
                obj = _esperar_objeto(sap_id, timeout=1.0)
                if obj:
                    txt = obj.Text
                    match = re.search(r"\b[A-Z0-9]{3,4}K\d{6,}\b", txt)
                    if match:
                        trkorr = match.group(0)
            except:
                pass
            if trkorr:
                break

        if okcd:
            okcd.text = "/n"
            _send_vkey(0)

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
                if req_input == "":
                    req_input = "4"
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

    print(f"\n📋 Roles a processar (agrupadas): {len(roles_agrupadas)}")
    for rr in roles_agrupadas:
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
        except:
            return ("", "")

    def try_actions(actions):
        for a in actions:
            try:
                ctrl = session.findById(a["path"])
                if a["op"] == "text":
                    ctrl.setFocus()
                    ctrl.text = a["val"]
                    _esperar_sap_livre()
                    return True
                elif a["op"] == "press":
                    if hasattr(ctrl, "Enabled") and not ctrl.Enabled:
                        continue
                    ctrl.press()
                    _esperar_sap_livre()
                    return True
                elif a["op"] == "select":
                    ctrl.select()
                    _esperar_sap_livre()
                    return True
            except:
                continue
        return False

    def tratar_popup_modal(max_loops=6):
        for _ in range(max_loops):
            try:
                _esperar_sap_livre(timeout=2.0, pausa=0.05)

                if session.ActiveWindow.Type != "GuiModalWindow":
                    return False

                try:
                    if session.findById("wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE", False) or \
                       session.findById("wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE1", False):
                        return False
                except:
                    pass

                candidatos = [
                    "wnd[1]/usr/btnBUTTON_1",
                    "wnd[1]/usr/btnSPOP-OPTION1",
                    "wnd[1]/tbar[0]/btn[0]",
                    "wnd[1]/tbar[0]/btn[19]",
                    "wnd[1]/tbar[0]/btn[11]"
                ]
                for p in candidatos:
                    if try_actions([{"path": p, "op": "press"}]):
                        return True
                return True
            except:
                return False
        return True

    def wait_wnd1_close(timeout=3.0):
        return _esperar_sumir("wnd[1]", timeout=timeout, pausa=0.05)

    ###################################################################################
    # BLOCO 3: Page Object PFCG
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
            _send_vkey(0)
            tratar_popup_modal()

        def set_role_name(self, nome):
            print(f"  ├─ Inserindo o nome da Role: {nome}")
            sap_id, obj = _resolver_id(
                "role_name_field",
                ["wnd[0]/usr/ctxtAGR_NAME_NEU", "wnd[0]/usr/ctxtAGR_NAME"]
            )
            if not obj:
                return False
            try:
                obj.setFocus()
                obj.text = nome
                _esperar_sap_livre()
                return True
            except:
                return False

        def open_for_edit(self):
            print("  ├─ Tentando abrir em modo de 'Criação'...")
            if not try_actions([
                {"path": "wnd[0]/usr/btn%#AUTOTEXT003", "op": "press"},
                {"path": "wnd[0]/tbar[1]/btn[5]", "op": "press"}
            ]):
                raise Exception("Não consegui clicar em Criar.")

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
            print("  ├─ Preenchendo a descrição da Role...")
            sap_id, obj = _resolver_id(
                "role_desc_field",
                ["wnd[0]/usr/txtS_AGR_TEXTS-TEXT", "wnd[0]/usr/txtS_AGR_TEXTS-TEXT1", "wnd[0]/usr/txtAGR_TEXTS-TEXT"]
            )
            if not obj:
                return False
            try:
                obj.text = desc
                _send_vkey(0)
                tratar_popup_modal()
                return True
            except:
                return False

        def save(self, log_msg="  └─ Guardando alterações..."):
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

        def goto_menu_tab(self):
            print("  ├─ Acedendo à aba 'Menu' (TAB9)...")
            sap_id, obj = _resolver_id("menu_tab", ["wnd[0]/usr/tabsTABSTRIP1/tabpTAB9"])
            if not obj:
                raise Exception("Não consegui abrir a aba Menu (TAB9).")
            try:
                obj.select()
                _esperar_sap_livre()
            except:
                raise Exception("Não consegui abrir a aba Menu (TAB9).")
            tratar_popup_modal()

        def goto_auth_tab(self):
            print("  ├─ Acedendo à aba 'Autorizações' (TAB5)...")
            sap_id, obj = _resolver_id("auth_tab", ["wnd[0]/usr/tabsTABSTRIP1/tabpTAB5"])
            if not obj:
                return False
            try:
                obj.select()
                _esperar_sap_livre()
                return True
            except:
                return False

        def _open_tcode_wizard(self):
            sap_id, obj = _resolver_id(
                "menu_shell",
                [
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubSUB1:SAPLPRGN_TREE:0321/cntlTOOL_CONTROL/shellcont/shell",
                    "wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubSUB1:SAPLPRGN_TREE:0320/cntlTOOL_CONTROL/shellcont/shell"
                ]
            )
            if obj:
                print("  ├─ Abrindo o Wizard de inserção de Transações (TB03)...")
                obj.pressButton("TB03")
                _esperar_objeto("wnd[1]", timeout=3.0)
                _esperar_sap_livre()
                return True
            raise Exception("Não encontrei o botão TB03.")

        def _fill_tcodes_fast(self, table_base, tcodes):
            if not tcodes:
                return 0
            inserted = 0
            for i, t in enumerate(tcodes):
                if i >= TCODE_BLOCK_SIZE:
                    break
                try:
                    cell_id = f"{table_base}/ctxtS_TCODES-TCODE[0,{i}]"
                    cell = self.sess.findById(cell_id)
                    cell.text = t
                    inserted += 1
                except Exception:
                    continue
            _esperar_sap_livre()
            return inserted

        def add_tcodes(self, tcodes):
            if not tcodes:
                return 0

            print(f"  ├─ Preparando a inserção rápida de {len(tcodes)} TCODE(s)...")
            inserted_total = 0
            total_blocos = max(1, ceil(len(tcodes) / TCODE_BLOCK_SIZE))

            for bloco in range(total_blocos):
                sub = tcodes[bloco * TCODE_BLOCK_SIZE: bloco * TCODE_BLOCK_SIZE + TCODE_BLOCK_SIZE]
                self._open_tcode_wizard()

                table_id, table_obj = _resolver_id(
                    "tcode_wizard_table",
                    ["wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE", "wnd[1]/usr/tblSAPLPRGN_WIZARDCTRL_TCODE1"]
                )
                if not table_id:
                    raise Exception("Não encontrei a tabela do Wizard de TCODE.")

                print(f"  ├─ Injetando bloco {bloco+1} na tabela...")
                qtd = self._fill_tcodes_fast(table_id, sub)
                inserted_total += qtd

                try:
                    self.sess.findById("wnd[1]").sendVKey(0)
                except:
                    pass

                _esperar_sap_livre()

                print("  ├─ Confirmando transações (Transferir)...")
                if not try_actions([
                    {"path": "wnd[1]/tbar[0]/btn[19]", "op": "press"},
                    {"path": "wnd[1]/tbar[0]/btn[0]", "op": "press"}
                ]):
                    pass

                wait_wnd1_close(timeout=2.0)
                tratar_popup_modal()

            return inserted_total

        def generate_authorization_profile(self):
            if not self.goto_auth_tab():
                return False
            tratar_popup_modal()

            print("  ├─ Clicando em 'Sugerir nome de perfil'...")
            try_actions([{
                "path": "wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubSUB1:SAPLPRGN_TREE:0350/btnPROFIL1",
                "op": "press"
            }])

            if _safe_find("wnd[1]"):
                print("  ├─ Confirmando a sugestão de nome de perfil no popup...")
                try_actions([{"path": "wnd[1]/tbar[0]/btn[11]", "op": "press"}])

            self.save("  ├─ Guardando a Role antes de gerar as autorizações...")

            print("  ├─ Acionando a Geração de Perfil... a aguardar...")
            try_actions([{"path": "wnd[0]/tbar[1]/btn[17]", "op": "press"}])

            if _safe_find("wnd[1]"):
                print("  ├─ Confirmando a geração de perfil na janela intermédia...")
                try_actions([{"path": "wnd[1]/usr/btnBUTTON_1", "op": "press"}])

            if _safe_find("wnd[1]"):
                print("  └─ Fechando popup de logs de autorização...")
                try:
                    self.sess.findById("wnd[1]").sendVKey(0)
                except:
                    pass
                _esperar_sap_livre()

            try:
                self.sess.findById("wnd[0]").sendVKey(0)
            except:
                pass
            _esperar_sap_livre()
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
    # BLOCO 4: EXECUÇÃO ESTRUTURADA
    ###################################################################################
    pfcg = PFCGPage(session)
    resultados = {}

    total_roles = len(roles_agrupadas)

    with Progress(
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TextColumn("({task.completed}/{task.total})"),
        TimeElapsedColumn(),
        transient=False,
    ) as progress:
        task_roles = progress.add_task("A processar roles...", total=total_roles)

        for idx, rr in enumerate(roles_agrupadas, start=1):
            nome, desc = (rr["AGR_NAME"] or "").strip(), (rr["TEXT"] or "").strip()
            tcodes = rr["TCODE_LIST"]

            progress.update(task_roles, description=f"A processar role: {nome}")

            print("\n======================================================================")
            print(f"▶ [{idx}/{len(roles_agrupadas)}] INICIANDO ROLE: {nome} | TCODEs: {len(tcodes)}")
            print("======================================================================")

            tempo_inicio_role = time.time()

            try:
                print("\n[Etapa 1] Preparação e Dados Básicos")
                pfcg.open()
                if not pfcg.set_role_name(nome):
                    raise Exception("Falha ao escrever AGR_NAME.")
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
                if mt in ("E", "A"):
                    err = sb
                resultados[nome] = {"STATUS": "ERRO", "MSG": err, "TIMESTEMP": now_ts()}

                print(f"\n🔴 ERRO: {err} ⏱️ (Tempo: {str_tempo})")
                print("----------------------------------------------------------------------")

                try:
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                    _send_vkey(0)
                except:
                    pass

            progress.advance(task_roles)

    ###################################################################################
    # BLOCO 5: GRAVAR EXCEL E TEMPO TOTAL
    ###################################################################################
    try:
        col_st, col_ms, col_tm = header_map.get("STATUS"), header_map.get("MSG"), header_map.get("TIMESTEMP")
        for rec in records:
            chave_busca = str(rec["AGR_NAME"]).strip()
            res = resultados.get(chave_busca)
            if res:
                if col_st:
                    ws.cell(row=rec["_row"], column=col_st).value = res["STATUS"]
                if col_ms:
                    ws.cell(row=rec["_row"], column=col_ms).value = res["MSG"]
                if col_tm:
                    ws.cell(row=rec["_row"], column=col_tm).value = res["TIMESTEMP"]

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