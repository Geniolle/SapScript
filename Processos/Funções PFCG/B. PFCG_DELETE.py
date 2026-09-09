# -*- coding: utf-8 -*-

###################################################################################
# BLOCO: ELIMINAÇÃO EM MASSA DE FUNÇÕES
# Por padrão (metodo="RFC") elimina via RFC (BPC_DELETE_SINGLE_ROLE, através de
# pfcg.pfcg_delete_rfc_service.bulk_delete_pfcg_roles_rfc), função a função, sem
# necessitar de sessão SAP GUI aberta. Só recorre ao SAP GUI (/NPFCGMASSDELETE)
# se o RFC falhar por motivo de infraestrutura (rede, credenciais, pyrfc
# indisponível) antes de sequer tentar eliminar alguma função, ou se
# metodo="GUI" for pedido explicitamente.
# Mantém a formatação do Excel intacta escrevendo célula a célula via openpyxl.
# Com barra de progresso Rich (modo GUI).
###################################################################################
def executar(
    ambiente_cockpit,
    pfcg_object,       # <-- Obrigatório: Cockpit deteta a aba pelo nome do script
    caminho_ficheiro,  # <-- Obrigatório: Cockpit abre a janela do Windows para o Excel
    request_transporte=None,
    modo_nao_interativo=False,
    pedir_confirmacao=True,
    metodo="RFC",       # <-- "RFC" (padrão) ou "GUI" (força o fluxo legado via /NPFCGMASSDELETE)
    **kwargs
):
    import os, time, sys, re
    import unicodedata
    from datetime import datetime
    import win32com.client
    from openpyxl import load_workbook
    from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn

    try:
        import pyperclip  # só é usado no fallback GUI (colar roles no popup do PFCGMASSDELETE)
    except ImportError:
        pyperclip = None

    # --- CORREÇÃO DA ESTRUTURA DE PASTAS PARA IMPORTAR O PESQUISAR_REQUEST ---
    dir_atual = os.path.dirname(os.path.abspath(__file__))
    dir_processos = os.path.dirname(dir_atual)
    if dir_processos not in sys.path:
        sys.path.insert(0, dir_processos)

    ###################################################################################
    # CONFIG INICIAL E LOGGING
    ###################################################################################
    tempo_inicio = time.time()
    mapa_sistema = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}
    sistema_desejado = mapa_sistema.get(ambiente_cockpit)

    metodo_normalizado = str(metodo or "RFC").strip().upper()
    if metodo_normalizado not in ("GUI", "RFC"):
        raise ValueError(f"Parâmetro 'metodo' inválido: '{metodo}'. Use 'GUI' ou 'RFC'.")

    NOME_SHEET = pfcg_object if pfcg_object else "PFCG_DELETE"
    SEARCH_HEADER_IN_FIRST_ROWS = 20

    TIMEOUT_SAP_BUSY = 180.0

    COL_ID        = "ID"
    COL_AGR_NAME  = "AGR_NAME"
    COL_TEXT      = "TEXT"
    COL_STATUS    = "STATUS"
    COL_MSG       = "MSG"
    COL_TIMESTAMP = "TIMESTEMP"

    COLUNAS_OBRIGATORIAS = {COL_ID, COL_AGR_NAME, COL_STATUS, COL_MSG, COL_TIMESTAMP}

    def agora_ts():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def log(msg):
        print(f"{agora_ts()} | {msg}", flush=True)

    ###################################################################################
    # HELPERS: NORMALIZAÇÃO
    ###################################################################################
    def norm_col(s):
        if s is None:
            return ""
        return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8").strip().upper()

    def traduzir_nome_coluna(s):
        s = norm_col(s)
        if s in ["NOME FUNCAO", "NOME FUNÇÂO", "NOME FUNÇAO"]:
            return COL_AGR_NAME
        if s in ["DESCRITIVO", "DESCRITIVO FUNCAO", "DESCRICAO", "DESCRIÇÃO"]:
            return COL_TEXT
        if s == "TIMESTAMP":
            return COL_TIMESTAMP
        return s

    ###################################################################################
    # HELPERS: SAP
    ###################################################################################
    def existe(session, obj_id):
        try:
            session.findById(obj_id)
            return True
        except:
            return False

    def esperar_sap_livre(session, timeout=120.0, pausa=0.2):
        limite = time.time() + timeout
        while time.time() < limite:
            try:
                busy = bool(getattr(session, "Busy", False))
            except:
                busy = False
            if not busy:
                return True
            time.sleep(pausa)
        return False

    def esperar_janela(session, wnd_idx, timeout=10.0, pausa=0.2):
        limite = time.time() + timeout
        while time.time() < limite:
            if existe(session, f"wnd[{wnd_idx}]"):
                return True
            time.sleep(pausa)
        return False

    def fechar_popups(session, timeout=10.0, pausa=0.2, prefer_yes=True):
        limite = time.time() + timeout
        while time.time() < limite:
            if any(existe(session, f"wnd[{i}]") for i in (1, 2, 3)):
                break
            time.sleep(pausa)

        fechou = False
        for _ in range(40):
            algum = False
            for i in (3, 2, 1):
                if not existe(session, f"wnd[{i}]"):
                    continue
                algum = True
                try:
                    if existe(session, f"wnd[{i}]/tbar[0]/btn[0]"):
                        session.findById(f"wnd[{i}]/tbar[0]/btn[0]").press()
                    elif prefer_yes and existe(session, f"wnd[{i}]/usr/btnSPOP-OPTION1"):
                        session.findById(f"wnd[{i}]/usr/btnSPOP-OPTION1").press()
                    elif existe(session, f"wnd[{i}]/usr/btnSPOP-OPTION2"):
                        session.findById(f"wnd[{i}]/usr/btnSPOP-OPTION2").press()
                    else:
                        session.findById(f"wnd[{i}]").sendVKey(0)
                    fechou = True
                except:
                    pass
                time.sleep(pausa)
            if not algum:
                break
        return fechou

    def mensagem_sem_resultado(msg):
        m = (msg or "").lower()
        return (
            ("nenhum" in m or "nenhuma" in m or "nenhumas" in m)
            and ("funç" in m or "regist" in m or "obj") and ("encontrad" in m)
        )

    ###################################################################################
    # LEITURA DO EXCEL VIA OPENPYXL (PRESERVA A FORMATAÇÃO)
    ###################################################################################
    print("\n[Etapa 1] Leitura do Excel")
    if not caminho_ficheiro or not os.path.exists(caminho_ficheiro):
        log(f"❌ Ficheiro Excel não encontrado: '{caminho_ficheiro}'.")
        return "voltar"

    try:
        wb = load_workbook(caminho_ficheiro, data_only=False)
    except Exception as e:
        log(f"❌ Erro ao abrir o ficheiro Excel: {e}")
        return "voltar"

    if NOME_SHEET in wb.sheetnames:
        ws = wb[NOME_SHEET]
    else:
        log(f"❌ Aba '{NOME_SHEET}' não encontrada no Excel.")
        log(f"💡 Abas disponíveis: {', '.join(wb.sheetnames)}")
        wb.close()
        return "voltar"

    log(f"📑 Aba (Sheet) lida com sucesso: '{NOME_SHEET}'")

    # Localizar o cabeçalho
    header_row = None
    header_map = {}
    for r in range(1, SEARCH_HEADER_IN_FIRST_ROWS + 1):
        row_vals = [traduzir_nome_coluna(c.value) for c in ws[r]]
        colunas_encontradas = set(row_vals).intersection(COLUNAS_OBRIGATORIAS)

        if len(colunas_encontradas) >= len(COLUNAS_OBRIGATORIAS):
            header_row = r
            for idx, name in enumerate(row_vals, start=1):
                if name:
                    header_map[name] = idx
            break

    if not header_row:
        wb.close()
        print(f"\n❌ Não encontrei as colunas obrigatórias nas primeiras {SEARCH_HEADER_IN_FIRST_ROWS} linhas.")
        print(f"   Esperado: {', '.join(COLUNAS_OBRIGATORIAS)}")
        return "voltar"

    # Extrair os dados para processamento
    records = []

    def get_cell(row_idx, col_name):
        if col_name not in header_map:
            return ""
        v = ws.cell(row=row_idx, column=header_map[col_name]).value
        return "" if v is None else str(v).strip()

    for r in range(header_row + 1, ws.max_row + 1):
        agr = get_cell(r, COL_AGR_NAME)
        status = get_cell(r, COL_STATUS).upper()

        # Filtra logo linhas vazias ou já concluídas
        if not agr or status == "CONCLUÍDO" or status == "CONCLUIDO":
            continue

        records.append({
            "_row": r,
            COL_AGR_NAME: agr
        })

    if not records:
        wb.close()
        log("⚠️ Nenhuma linha válida pendente encontrada na aba.")
        return "voltar"

    # Extrair lista de roles e copiar para o Clipboard
    funcoes = [rec[COL_AGR_NAME] for rec in records]
    if pyperclip is not None:
        try:
            pyperclip.copy("\r\n".join(funcoes))
        except Exception:
            pass

    ###################################################################################
    # CAPTURA SESSÃO SAP
    ###################################################################################
    print("\n[Etapa 2] Acesso ao SAP")
    try:
        log("🔌 A localizar sessão SAP...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        session = next(
            (sess for conn in app.Children for sess in conn.Children if sess.Info.SystemName.upper() == (sistema_desejado or "")),
            None
        )
    except Exception:
        session = None

    if not session:
        if metodo_normalizado == "GUI":
            log(f"❌ Sessão SAP não encontrada para '{ambiente_cockpit}' (esperado: {sistema_desejado}).")
            wb.close()
            return "voltar"
        log(
            f"ℹ️ Sessão SAP GUI não encontrada para '{ambiente_cockpit}' — segue apenas por RFC "
            f"(só será necessária se o RFC falhar)."
        )

    ###################################################################################
    # LÓGICA DA REQUEST DE TRANSPORTE
    ###################################################################################
    def _criar_nova_request_no_sap_local():
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE10"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)

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

            session.findById("wnd[0]/tbar[1]/btn[6]").press()
            time.sleep(0.4)

            if tipo == "2":
                try:
                    session.findById("wnd[1]/usr/radKO042-REQ_CONS_K").select()
                except:
                    pass

            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.4)

            try:
                session.findById("wnd[1]/usr/txtKO013-AS4TEXT").text = desc
            except:
                pass

            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.6)

            trkorr = None
            for sap_id in ["wnd[0]/usr/lbl[20,9]", "wnd[0]/usr/lbl[1,1]"]:
                try:
                    txt = session.findById(sap_id).Text
                    match = re.search(r"\b[A-Z0-9]{3,4}K\d{6,}\b", txt)
                    if match:
                        trkorr = match.group(0)
                except:
                    pass
                if trkorr:
                    break

            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            print(f"\n✅ Request criada: {trkorr}")
            return trkorr
        except Exception as e:
            print(f"❌ Falha ao criar request: {e}")
            return None

    # Parâmetros de transporte para o caminho RFC (resolvidos abaixo a partir da
    # mesma escolha de menu usada pelo caminho GUI).
    rfc_transport_mode = "LOCAL"
    rfc_request_number = ""
    rfc_request_description = ""

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
            if metodo_normalizado == "RFC":
                desc_nova = input("📝 Descrição da nova Request a criar (máx 60): ").strip()
                rfc_request_description = (desc_nova or "REQUEST CRIADA VIA SCRIPT (RFC)")[:60]
                rfc_transport_mode = "CREATE_REQUEST"
                print("ℹ️ A nova Request será criada automaticamente via RFC ao eliminar as funções.")
            else:
                request_transporte = _criar_nova_request_no_sap_local()

        elif req_input == "3":
            try:
                import pesquisar_request
                print("\n🔍 A abrir nova sessão em segundo plano para pesquisar (SE16H)...")
                resultados_pesquisa = pesquisar_request.listar_requests(
                    system_name=sistema_desejado,
                    include_requests=True,
                    use_new_mode=True,
                    minimize=True,
                    close_after=True
                )
                if resultados_pesquisa:
                    escolha = input("\n👉 Digite o número (N) da Request que deseja utilizar (ou Enter para cancelar): ").strip()
                    if escolha.isdigit() and 1 <= int(escolha) <= len(resultados_pesquisa):
                        request_transporte = resultados_pesquisa[int(escolha) - 1][0]
                        print(f"✅ Selecionou a Request: {request_transporte}")
                    else:
                        print("❌ Seleção cancelada. Não haverá transporte.")
                else:
                    print("⚠️ Não foram encontradas Requests abertas.")
            except Exception as e:
                print(f"❌ Erro na pesquisa: {e}")

        elif req_input == "4":
            print("⏭️  Nenhuma request selecionada (Transporte ignorado).")
            request_transporte = None

        print("============================================================")

    # Resolve os parâmetros de transporte para o caminho RFC a partir da mesma
    # escolha do menu acima (exceto quando já resolvido em "2 - criar nova" via RFC).
    if rfc_transport_mode != "CREATE_REQUEST":
        if request_transporte:
            rfc_transport_mode = "EXISTING_REQUEST"
            rfc_request_number = request_transporte
        else:
            rfc_transport_mode = "LOCAL"

    log(f"📂 Ficheiro: {caminho_ficheiro}")
    log(f"📋 Roles a eliminar ({len(funcoes)}):")
    for i, n in enumerate(funcoes, 1):
        print(f" {i:02d}. {n}", flush=True)

    ###################################################################################
    # EXECUÇÃO — RFC (padrão), com fallback para SAP GUI só em falha de infraestrutura
    ###################################################################################
    usar_gui = (metodo_normalizado == "GUI")
    resultado_rfc = None

    if not usar_gui:
        print("\n[Etapa 3] Eliminação das Funções via RFC")
        try:
            from pfcg.pfcg_delete_rfc_service import bulk_delete_pfcg_roles_rfc
        except Exception as e:
            log(f"⚠️ RFC indisponível (import falhou): {e}. A avançar para SAP GUI.")
            usar_gui = True
        else:
            try:
                log(f"🔎 A eliminar {len(funcoes)} função(ões) via RFC em {ambiente_cockpit}...")
                resultado_rfc = bulk_delete_pfcg_roles_rfc(
                    ambiente_cockpit, funcoes, rfc_transport_mode, rfc_request_number, rfc_request_description
                )
            except Exception as e:
                log(f"⚠️ Falha ao invocar RFC: {e}. A avançar para SAP GUI.")
                resultado_rfc = None
                usar_gui = True

        if resultado_rfc is not None and not resultado_rfc.get("items"):
            # Falhou antes de sequer tentar eliminar alguma função (infraestrutura) -> fallback GUI.
            log(
                f"⚠️ RFC falhou antes de processar qualquer função: "
                f"{resultado_rfc.get('error_type') or resultado_rfc.get('status')} - "
                f"{resultado_rfc.get('message')}. A avançar para SAP GUI."
            )
            resultado_rfc = None
            usar_gui = True

    if resultado_rfc is not None:
        # RFC produziu resultado por função (mesmo que parcial) — usa-se este resultado,
        # não se recorre ao GUI para as mesmas roles.
        itens_por_role = {
            str(item.get("role", "")).strip().upper(): item for item in resultado_rfc.get("items", [])
        }
        ts_final = agora_ts()
        concluidos = 0
        falhados = 0
        for rec in records:
            nome_role = rec[COL_AGR_NAME]
            item = itens_por_role.get(nome_role.strip().upper())
            if item is None:
                status_role = "ERRO"
                msg_role = "Função não retornada pelo RFC (verifique o nome exato)."
            elif item.get("ok"):
                status_role = "CONCLUÍDO"
                msg_role = item.get("message") or f"Eliminada via RFC ({rfc_transport_mode})."
                concluidos += 1
            else:
                status_role = "ERRO"
                msg_role = item.get("message") or f"Falha RFC: {item.get('status')}"
                falhados += 1
            rec["_STATUS_RFC"] = status_role
            rec["_MSG_RFC"] = msg_role
            rec["_TS_RFC"] = ts_final
            if status_role == "CONCLUÍDO":
                print(f"[OK] Role concluída (RFC): {nome_role}", flush=True)
            else:
                print(f"Falha ao eliminar role (RFC): {nome_role} — {msg_role}", flush=True)

        log(f"└─ RFC concluído: {concluidos} eliminada(s), {falhados} com falha, de {len(records)} função(ões).")

        print("\n[Etapa 4] Atualização do Excel")
        log("💾 A gravar resultados no Excel preservando formatações...")
        try:
            col_st = header_map.get(COL_STATUS)
            col_ms = header_map.get(COL_MSG)
            col_tm = header_map.get(COL_TIMESTAMP)
            for rec in records:
                if col_st:
                    ws.cell(row=rec["_row"], column=col_st).value = rec["_STATUS_RFC"]
                if col_ms:
                    ws.cell(row=rec["_row"], column=col_ms).value = rec["_MSG_RFC"]
                if col_tm:
                    ws.cell(row=rec["_row"], column=col_tm).value = rec["_TS_RFC"]
            wb.save(caminho_ficheiro)
            print(f"✅ Ficheiro atualizado com os resultados na aba '{NOME_SHEET}'.")
        except Exception as e:
            print(f"❌ Erro ao salvar o ficheiro: {e}")
            print("⚠️ Verifica se o ficheiro está aberto e bloqueado no Excel.")
        finally:
            wb.close()

        mm, ss = divmod(int(time.time() - tempo_inicio), 60)
        log(f"⏱️ Tempo total: {mm:02d}:{ss:02d}")
        return "voltar"

    # --- A PARTIR DAQUI: MODO GUI (usado quando metodo="GUI" explícito, ou como
    # fallback se o RFC tiver falhado por infraestrutura antes de processar alguma função) ---
    if pyperclip is None:
        log(
            "⚠️ Módulo 'pyperclip' indisponível neste Python — a colagem automática da "
            "lista de roles no popup do PFCGMASSDELETE pode falhar ou ficar vazia."
        )
    if not session:
        log(
            f"❌ Sessão SAP GUI não encontrada para '{ambiente_cockpit}' (esperado: {sistema_desejado}) — "
            f"impossível continuar em modo GUI."
        )
        wb.close()
        return "voltar"

    ###################################################################################
    # EXECUÇÃO NO SAP (Eliminação em Massa via PFCGMASSDELETE)
    ###################################################################################
    print("\n[Etapa 3] Eliminação das Funções")
    status_geral = "ERRO"
    msg_final = "Erro desconhecido."

    try:
        # session.findById("wnd[0]").maximize()
        if not esperar_sap_livre(session, timeout=TIMEOUT_SAP_BUSY):
            raise RuntimeError("SAP bloqueado antes de iniciar.")

        log("├─ A abrir transação /NPFCGMASSDELETE...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NPFCGMASSDELETE"
        session.findById("wnd[0]").sendVKey(0)

        try:
            sb = session.findById("wnd[0]/sbar").Text.strip()
            if sb:
                log(f"   ✔️ SAP: {sb}")
        except:
            pass

        if existe(session, "wnd[0]/usr/radMOD_EXE"):
            session.findById("wnd[0]/usr/radMOD_EXE").select()

        # Abrir lista múltipla e colar as funções
        log(f"├─ A carregar lista de {len(funcoes)} funções no campo de seleção SAP...")
        session.findById("wnd[0]/usr/btn%_ROLE_%_APP_%-VALU_PUSH").press()
        time.sleep(0.5)
        session.findById("wnd[1]").sendVKey(24)  # Shift+F12 (Colar)
        time.sleep(0.3)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(0.3)
        log(f"│  └─ {len(funcoes)} funções copiadas do clipboard e confirmadas no SAP.")

        log("├─ A executar a eliminação em massa (PFCGMASSDELETE)...")
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Gestão de Popups com feedback
        timeout = time.time() + 15.0
        ultimo_log_heartbeat = time.time()
        popup_count = 0
        while time.time() < timeout:
            time.sleep(0.5)
            elapsed = time.time() - (timeout - 15.0)

            # Heartbeat a cada 3 segundos durante o processamento SAP
            if time.time() - ultimo_log_heartbeat >= 3.0:
                log(f"│  ⏳ SAP a processar... ({elapsed:.0f}s decorridos)")
                ultimo_log_heartbeat = time.time()

            if existe(session, "wnd[1]/usr/ctxtKO008-TRKORR"):
                popup_count += 1
                if request_transporte:
                    log(f"│  ├─ Popup de transporte detetado. A injetar Request: {request_transporte}")
                    session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = request_transporte
                else:
                    log("│  ├─ Popup de transporte detetado. A ignorar (sem transporte).")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                continue
            if existe(session, "wnd[1]/usr/btnSPOP-OPTION1"):
                popup_count += 1
                log("│  ├─ Popup de confirmação SAP. A aceitar...")
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                continue
            if existe(session, "wnd[1]/tbar[0]/btn[0]"):
                popup_count += 1
                log("│  ├─ Popup genérico SAP. A fechar...")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                continue
            if existe(session, "wnd[0]/usr/cntlGRID1/shellcont/shell"):
                log("│  └─ ALV de resultado detetado. A ler resultados...")
                break

        # Ler resultado do ALV
        msg_alv = ""
        alv_rows = 0
        try:
            grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            alv_rows = int(grid.RowCount) if grid.RowCount else 0
            log(f"├─ ALV de resultado contém {alv_rows} linha(s).")
            if alv_rows > 0:
                for col in ["MESSAGE", "TEXT", "MSG"]:
                    try:
                        v = str(grid.GetCellValue(0, col)).strip()
                        if v:
                            msg_alv = v
                            log(f"│  └─ Mensagem ALV: {v}")
                            break
                    except:
                        pass
        except:
            pass

        msg_barra = ""
        try:
            msg_barra = session.findById("wnd[0]/sbar").Text.strip()
            if msg_barra:
                log(f"├─ SAP Status Bar: {msg_barra}")
        except:
            pass

        msg_final = msg_alv or msg_barra or "Execução concluída (Verificada no Log ALV)"
        msg_transporte = f" [Req: {request_transporte}]" if request_transporte else ""

        # Voltar ao ecrã base
        try:
            if existe(session, "wnd[0]/tbar[0]/btn[3]"):
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass

        if mensagem_sem_resultado(msg_final):
            status_geral = "ERRO"
            msg_final = f"{msg_final} - SAP não encontrou as roles informadas."
            log(f"└─ ❌ {msg_final}")
        else:
            status_geral = "CONCLUÍDO"
            msg_final = f"{msg_final}{msg_transporte}"
            log(f"└─ ✅ SAP status final: {msg_final}")

    except Exception as e:
        status_geral = "ERRO"
        msg_final = f"Erro no processo SAP: {e}"
        log(f"❌ Erro crítico no SAP: {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass

    ###################################################################################
    # GRAVAÇÃO CÉLULA A CÉLULA VIA OPENPYXL (Com barra de progresso)
    ###################################################################################
    print("\n[Etapa 4] Atualização do Excel")
    log("💾 A gravar resultados no Excel preservando formatações...")
    try:
        col_st = header_map.get(COL_STATUS)
        col_ms = header_map.get(COL_MSG)
        col_tm = header_map.get(COL_TIMESTAMP)

        ts_final = agora_ts()

        with Progress(
            TextColumn("[bold cyan]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TextColumn("({task.completed}/{task.total})"),
            TimeElapsedColumn(),
            transient=False,
        ) as progress:
            task_excel = progress.add_task("A atualizar Excel...", total=len(records))

            for rec in records:
                linha_excel = rec["_row"]
                nome_role = rec[COL_AGR_NAME]

                progress.update(task_excel, description=f"A atualizar Excel: {nome_role}")

                if col_st:
                    ws.cell(row=linha_excel, column=col_st).value = status_geral
                if col_ms:
                    ws.cell(row=linha_excel, column=col_ms).value = msg_final
                if col_tm:
                    ws.cell(row=linha_excel, column=col_tm).value = ts_final

                # Relatar conclusão de cada função no dashboard
                if status_geral == "CONCLUÍDO":
                    print(f"[OK] Role concluída: {nome_role}", flush=True)
                else:
                    print(f"Falha ao eliminar role: {nome_role}", flush=True)

                progress.advance(task_excel)

        wb.save(caminho_ficheiro)
        print(f"✅ Ficheiro atualizado com os resultados na aba '{NOME_SHEET}'.")
    except Exception as e:
        print(f"❌ Erro ao salvar o ficheiro: {e}")
        print("⚠️ Verifica se o ficheiro está aberto e bloqueado no Excel.")
    finally:
        wb.close()

    ###################################################################################
    # TEMPO TOTAL
    ###################################################################################
    mm, ss = divmod(int(time.time() - tempo_inicio), 60)
    log(f"⏱️ Tempo total: {mm:02d}:{ss:02d}")
    return "voltar"