# -*- coding: utf-8 -*-
###################################################################################
# SCRIPT: Criar e Atribuir Cadeias.py
#
# Objetivo:
#  - Ler um Excel com colunas normalizadas:
#       NOME CADEIA DE PESQUISA | EMPRESA | BANCO | ID CONTA | OPERAÇÃO EXTERNA | ... | STATUS | MSG
#  - Executar em duas fases sequenciais:
#       Fase 1: Criar a cadeia no SAP (OTPM) caso não exista na tabela TPAMA.
#       Fase 2: Atribuir a cadeia no SAP (OTPM) caso a atribuição não exista na tabela T028P.
#  - Capturar mensagens reais do SAP e atualizar o STATUS/MSG no próprio Excel.
#
# Regras:
#  - Integra pesquisa e criação automática de request de transporte.
#  - Mantém compatibilidade com o Cockpit (GUI e Web Worker).
###################################################################################

def executar(
    ambiente_cockpit,
    request_ctx,                 # OBRIGATÓRIO para compatibilidade com o cockpit
    request_transporte=None,
    request_description="",
    caminho_ficheiro=None,
    modo_nao_interativo=False,
    pedir_confirmacao=True,
    cliente_sap=None,
    chamado_pelo_main=False
):
    import re
    import os
    import subprocess
    import time
    import warnings
    import unicodedata
    import pandas as pd
    import win32com.client
    import tkinter as tk
    from tkinter import filedialog
    import sys

    registro_fase1 = {}
    registro_fase2 = {}
    evidencias_dir = ""
    timestamp = time.strftime("%Y%m%d_%H%M%S")

    def _parse_env_line(line: str):
        raw = str(line or "").strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            return None, None
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and (
            (value.startswith('"') and value.endswith('"'))
            or (value.startswith("'") and value.endswith("'"))
        ):
            value = value[1:-1]
        return key, value

    def _load_dotenv_manual_local():
        base_script = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(os.getcwd(), ".env"),
            os.path.join(base_script, ".env"),
            os.path.join(base_script, "..", "..", ".env"),
            os.path.join(base_script, "..", "..", "..", ".env"),
        ]
        seen = set()
        for path in candidates:
            path_abs = os.path.abspath(path)
            if path_abs in seen:
                continue
            seen.add(path_abs)
            if not os.path.exists(path_abs):
                continue
            with open(path_abs, "r", encoding="utf-8-sig") as file_obj:
                for raw in file_obj:
                    key, value = _parse_env_line(raw)
                    if key and key not in os.environ:
                        os.environ[key] = value or ""
            return path_abs
        return None

    _load_dotenv_manual_local()

    warnings.simplefilter("ignore", UserWarning)
    warnings.simplefilter("ignore", FutureWarning)

    called_by_main = bool(chamado_pelo_main) or str(os.getenv("SAP_CALLED_BY_MAIN", "")).strip().lower() in {
        "1", "true", "yes", "on", "sim", "s"
    }
    keep_validation_screen = str(os.getenv("WORKFLOW_CAPTURE_VALIDATION_SCREEN", "")).strip().lower() in {
        "1", "true", "yes", "on", "sim", "s"
    }
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if called_by_main:
        modo_nao_interativo = True
        pedir_confirmacao = False

    MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    SISTEMA_ESPERADO = MAPA_SISTEMA.get(ambiente_cockpit)
    CLIENTE_ESPERADO = str(cliente_sap or "").strip()

    if not SISTEMA_ESPERADO:
        print(f"❌ Ambiente inválido: {ambiente_cockpit}")
        return

    # ----- cabeçalhos com sinónimos -----
    COL_NECESSARIAS = {
        "NOME CADEIA DE PESQUISA": {"NOME CADEIA DE PESQUISA", "NOME DA CADEIA DE PESQUISA", "NOME CADEIA PESQUISA"},
        "EMPRESA": {"EMPRESA", "BUKRS", "COMPANY CODE"},
        "BANCO": {"BANCO", "HBKID", "ID BANCO"},
        "ID CONTA": {"ID CONTA", "HKTID", "ID DA CONTA", "ID-CONTA"},
        "OPERAÇÃO EXTERNA": {"OPERAÇÃO EXTERNA", "OPERACAO EXTERNA", "VGEXT", "OP EXTERNA"},
        "SINAL (+/-)": {"SINAL (+/-)", "SINAL", "VOZPM", "SINAL +-"},
        "NOME DO CAMPO DESTINO": {"NOME DO CAMPO DESTINO", "NOME CAMPO DESTINO", "ALVO", "DESTINO"},
        "CAMPO DESTINO": {"CAMPO DESTINO", "TARGFI", "CAMPO DE DESTINO"},
        "BASE DE MAPEAMENTO": {"BASE DE MAPEAMENTO", "PREFIX", "BASE MAPEAMENTO"},
    }

    # ----- Tabela fixa BASE -> sinal (EBVGINT) -----
    _base_sign_pairs = [
        ("CP0100000000000", "-"),
        ("Z001", "+"),
        ("Z002", "+"), ("Z002", "-"),
        ("Z006", "+"),
        ("Z007", "-"),
        ("Z021", "-"),
        ("Z022", "+"),
        ("Z030", "+"),
        ("Z031", "-"),
        ("Z032", "+"),
        ("Z033", "-"),
        ("Z050", "+"),
        ("Z051", "-"),
        ("Z056", "+"),
        ("Z057", "-"),
        ("Z060", "+"),
        ("Z061", "-"),
        ("Z062", "+"),
        ("Z063", "-"),
        ("Z118", "+"),
        ("Z119", "-"),
        ("ZDD3", "+"),
        ("ZDD4", "-"),
    ]
    BASE_TO_ALLOWED_SIGNS = {}
    for b, s in _base_sign_pairs:
        BASE_TO_ALLOWED_SIGNS.setdefault(b.upper(), set()).add(s)

    ###################################################################################
    # HELPERS GERAIS
    ###################################################################################
    def norm(s: str) -> str:
        if s is None:
            return ""
        s = str(s)
        s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
        s = s.replace("\u00A0", " ")
        s = re.sub(r"\s+", " ", s.strip())
        return s.upper()

    def safe_value(val):
        return "" if pd.isna(val) or str(val).strip().lower() in {"nan", "none"} else str(val).strip()

    def series_to_upper_clean(s: pd.Series) -> pd.Series:
        return s.fillna("").astype(str).str.strip().str.upper()

    def map_sign_free(val: str) -> str:
        v = (val or "").strip().upper()
        if v in {"+", "PLUS", "POS", "P"}:
            return "+"
        if v in {"-", "MINUS", "NEG", "M"}:
            return "-"
        return ""

    def aplicar_modo_janela_sap(sess) -> None:
        modo = str(os.getenv("SAP_WINDOW_MODE", "") or "").strip().lower()
        if not modo:
            minimizar_bool = str(os.getenv("SAP_WINDOW_MINIMIZE", "false") or "").strip().lower()
            if minimizar_bool in {"1", "true", "yes", "on", "sim", "s"}:
                modo = "minimize"
        if not modo:
            modo = "show"

        try:
            wnd0 = sess.findById("wnd[0]")
        except Exception:
            return

        try:
            if modo in {"minimize", "minimizar", "hidden", "hide", "ocultar", "quiet"}:
                wnd0.iconify()
            elif modo in {"show", "mostrar", "visible", "visivel", "exibir"}:
                wnd0.maximize()
        except Exception:
            return

    def selecionar_ficheiro_excel() -> str:
        print("📂 Selecione o ficheiro Excel (janela foi colocada em primeiro plano)...")
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.attributes("-topmost", True)
        root.focus_force()
        root.update()
        try:
            caminho = filedialog.askopenfilename(
                parent=root,
                title="Selecione o ficheiro de Cadeias de Pesquisa",
                filetypes=[("Ficheiros Excel", "*.xlsx"), ("Todos os ficheiros", "*.*")],
            )
        finally:
            try:
                root.attributes("-topmost", False)
            except Exception:
                pass
            root.destroy()
        return caminho

    def abrir_excel_para_dataframe(caminho: str) -> pd.DataFrame:
        try:
            return pd.read_excel(caminho, sheet_name="Folha2", dtype=str).fillna("")
        except PermissionError:
            print("\n❌ ERRO: O ficheiro Excel está ABERTO.")
            print("👉 Por favor, feche o ficheiro Excel e execute o script novamente.")
            sys.exit(1)
        except Exception:
            try:
                return pd.read_excel(caminho, dtype=str).fillna("")
            except PermissionError:
                print("\n❌ ERRO: O ficheiro Excel está ABERTO.")
                print("👉 Por favor, feche o ficheiro Excel e execute o script novamente.")
                sys.exit(1)

    def sanitize_paname(v: str, maxlen: int = 20) -> str:
        vv = safe_value(v)
        vv = "".join(c for c in unicodedata.normalize("NFKD", vv) if not unicodedata.combining(c))
        vv = re.sub(r"\s+", " ", vv.strip())
        vv = re.sub(r"[^A-Za-z0-9_\-\.\/ ]", "", vv)
        vv = vv.upper()
        return vv[:maxlen]

    def resolver_colunas(df: pd.DataFrame):
        cols_norm = {c: norm(c) for c in df.columns}
        inv = {}
        for orig, n in cols_norm.items():
            inv.setdefault(n, orig)

        resolvidas, faltantes = {}, []
        for canonico, sinonimos in COL_NECESSARIAS.items():
            s_norm = {norm(x) for x in sinonimos}
            match = next((inv[nc] for nc in s_norm if nc in inv), None)
            if match:
                resolvidas[canonico] = match
            else:
                faltantes.append(canonico)

        return resolvidas, faltantes

    ###################################################################################
    # HELPERS SAP
    ###################################################################################
    def _safe_find(sess, sap_id):
        try:
            return sess.findById(sap_id)
        except Exception:
            return None

    def _sap_busy(sess):
        try:
            return bool(getattr(sess, "Busy", False))
        except Exception:
            return False

    def _esperar_sap_livre(sess, timeout=8.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not _sap_busy(sess):
                return True
            time.sleep(pausa)
        return False

    def _esperar_objeto(sess, sap_id, timeout=4.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            obj = _safe_find(sess, sap_id)
            if obj:
                return obj
            time.sleep(pausa)
        return None

    def _send_vkey(sess, vkey, wait_after=True):
        sess.findById("wnd[0]").sendVKey(vkey)
        if wait_after:
            _esperar_sap_livre(sess)

    def get_statusbar(sess):
        try:
            sbar = sess.findById("wnd[0]/sbar")
            tipo = str(getattr(sbar, "MessageType", "") or "").strip().upper()
            texto = str(getattr(sbar, "Text", "") or "").strip()
            if texto:
                print(f"[SAP_SBAR] {texto}")
            return tipo, texto
        except Exception:
            return "", ""

    def sbar_error(sess) -> str:
        tipo, texto = get_statusbar(sess)
        if tipo in ("E", "A"):
            return texto
        return ""

    def fechar_popup_se_existir(sess):
        caminhos = [
            "wnd[1]/tbar[0]/btn[0]",
            "wnd[1]/tbar[0]/btn[11]",
            "wnd[1]/tbar[0]/btn[12]",
        ]
        for p in caminhos:
            try:
                obj = sess.findById(p)
                obj.press()
                _esperar_sap_livre(sess)
                return True
            except Exception:
                pass
        try:
            sess.findById("wnd[1]").sendVKey(0)
            _esperar_sap_livre(sess)
            return True
        except Exception:
            return False

    def limpar_para_home(sess):
        try:
            sess.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            _send_vkey(sess, 0)
        except Exception:
            pass

    def try_press(sess, caminhos):
        for p in caminhos:
            try:
                obj = sess.findById(p)
                obj.press()
                _esperar_sap_livre(sess)
                return True
            except Exception:
                continue
        return False

    def _try_set_text(sess, caminhos, valor):
        for p in caminhos:
            try:
                obj = sess.findById(p)
                obj.text = valor
                return True
            except Exception:
                continue
        return False

    def set_cell_any(sess, table_base_id, field_name, col_index, row_index, value):
        value = "" if value is None else str(value)
        variants = [
            (f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]", "key",  value),
            (f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]", "text", value),
            (f"{table_base_id}/txt{field_name}[{col_index},{row_index}]",  "text", value),
        ]
        for ctrl_id, prop, val in variants:
            try:
                ctrl = sess.findById(ctrl_id)
                setattr(ctrl, prop, val)
                return True
            except Exception:
                continue
        return False

    def get_cell_text(sess, table_base_id, field_name, col_index, row_index):
        for ctrl_id, attr in [
            (f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]", "key"),
            (f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]", "text"),
            (f"{table_base_id}/txt{field_name}[{col_index},{row_index}]",  "text"),
        ]:
            try:
                ctrl = sess.findById(ctrl_id)
                return getattr(ctrl, attr)
            except Exception:
                continue
        return ""

    def set_combo_key_then_check(sess, table_base_id, field_name, col_index, row_index, key_expected):
        try:
            ctrl = sess.findById(f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]")
            ctrl.key = str(key_expected)
            time.sleep(0.1)
            ok = (ctrl.key or "").strip().upper() == str(key_expected).strip().upper()
            if not ok:
                ctrl = sess.findById(f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]")
                ctrl.text = str(key_expected)
                time.sleep(0.1)
                ok = (ctrl.text or "").strip().upper() == str(key_expected).strip().upper()
            return ok
        except Exception:
            return set_cell_any(sess, table_base_id, field_name, col_index, row_index, key_expected)

    ###################################################################################
    # SE16H / BD FETCH HELPERS
    ###################################################################################
    def _se16h_tem_resultados(sess, max_nodes=7000):
        raiz = _safe_find(sess, "wnd[0]")
        if not raiz:
            return False

        pilha = [raiz]
        visitados = 0
        while pilha and visitados < max_nodes:
            obj = pilha.pop()
            visitados += 1
            try:
                row_count = int(getattr(obj, "RowCount", 0))
                if row_count > 0:
                    return True
            except Exception:
                pass
            try:
                filhos = int(obj.Children.Count)
            except Exception:
                filhos = 0
            for idx in range(filhos):
                try:
                    pilha.append(obj.Children(idx))
                except Exception:
                    continue
        return False

    def _extract_first_int(texto):
        raw = str(texto or "")
        match = re.search(r"(\d[\d\.\, ]*)", raw)
        if not match:
            return None
        digits = re.sub(r"\D", "", match.group(1))
        if not digits:
            return None
        try:
            return int(digits)
        except Exception:
            return None

    def _norm_key_tuple(values):
        return tuple(safe_value(v).strip().upper() for v in values)

    def _find_first_grid_with_rows(sess):
        raiz = _safe_find(sess, "wnd[0]")
        if not raiz:
            return None
        pilha = [raiz]
        visitados = 0
        while pilha and visitados < 9000:
            obj = pilha.pop()
            visitados += 1
            try:
                row_count = int(getattr(obj, "RowCount", 0))
            except Exception:
                row_count = 0
            if row_count > 0:
                try:
                    getattr(obj, "GetCellValue")
                    return obj
                except Exception:
                    pass
            try:
                filhos = int(obj.Children.Count)
            except Exception:
                filhos = 0
            for idx in range(filhos):
                try:
                    pilha.append(obj.Children(idx))
                except Exception:
                    continue
        return None

    def _grid_get_cell(grid, row_idx, candidates):
        for col in candidates:
            try:
                value = grid.GetCellValue(int(row_idx), str(col))
                value = "" if value is None else str(value).strip()
                if value:
                    return value
            except Exception:
                continue
        return ""

    def _cadeia_existe_em_tpama(sess, nome_cadeia):
        nome_cadeia = (nome_cadeia or "").strip()[:20]
        if not nome_cadeia:
            return False, "Nome da cadeia vazio.", ""

        try:
            sess.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16H"
            _send_vkey(sess, 0)

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                return None, sb, ""

            if not _try_set_text(sess, [
                "wnd[0]/usr/ctxtGD-TAB",
                "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
                "wnd[0]/usr/ctxtTABNAME",
            ], "TPAMA"):
                return None, "Nao consegui preencher a tabela TPAMA na SE16H.", ""

            _send_vkey(sess, 0)
            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                return None, sb, ""

            campo_panam = _esperar_objeto(
                sess,
                "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]",
                timeout=2.5,
            )
            if not campo_panam:
                return None, "Nao consegui localizar o campo PANAM na TPAMA.", ""

            campo_panam.text = nome_cadeia

            if not try_press(sess, ["wnd[0]/tbar[1]/btn[8]"]):
                return None, "Nao consegui executar a pesquisa na TPAMA.", ""

            nao_encontrou_tokens = ("NO VALUES", "NOT FOUND", "NENHUM", "NAO ENCONTR", "SEM REGIST")
            encontrou_tokens = ("SELECION", "SELECTED", "REGIST", "ENTRAD", "VALUES")

            scr = ""
            for _ in range(20):
                mt, sb = get_statusbar(sess)
                msg_norm = norm(sb)

                if msg_norm and any(t in msg_norm for t in nao_encontrou_tokens):
                    return False, sb or f"Cadeia '{nome_cadeia}' nao encontrada na TPAMA.", ""
                if msg_norm and any(t in msg_norm for t in encontrou_tokens):
                    scr = capturar_print_sap(sess, f"Validacao_TPAMA_{nome_cadeia}")
                    return True, sb, scr

                wnd0 = _safe_find(sess, "wnd[0]")
                wnd_text = norm(getattr(wnd0, "Text", "") if wnd0 else "")
                if "ENTRADAS ENCONTRADAS" in wnd_text or "ENTRIES FOUND" in wnd_text:
                    scr = capturar_print_sap(sess, f"Validacao_TPAMA_{nome_cadeia}")
                    return True, sb or f"Cadeia '{nome_cadeia}' encontrada na TPAMA.", scr

                if _se16h_tem_resultados(sess):
                    scr = capturar_print_sap(sess, f"Validacao_TPAMA_{nome_cadeia}")
                    return True, sb or f"Cadeia '{nome_cadeia}' encontrada na TPAMA.", scr

                time.sleep(0.2)

            return False, sb or f"Cadeia '{nome_cadeia}' nao encontrada na TPAMA.", ""
        except Exception as e:
            return None, str(e), ""
        finally:
            if not keep_validation_screen:
                limpar_para_home(sess)

    def _t028p_fetch_keys_by_panam(sess, panam):
        try:
            sess.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16H"
            sess.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)

            if not _try_set_text(sess, [
                "wnd[0]/usr/ctxtGD-TAB",
                "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
                "wnd[0]/usr/ctxtTABNAME",
            ], "T028P"):
                return None, "Nao consegui definir T028P na SE16H.", None, set(), ""

            sess.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)

            base = "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC"
            panam = safe_value(panam)
            if not panam:
                return None, "PANAM vazio para validacao da T028P.", None, set(), ""

            ctrl_id = f"{base}/ctxtGS_SELFIELDS-LOW[2,5]"
            obj = _safe_find(sess, ctrl_id)
            if not obj:
                return None, "Campo PANAM nao encontrado na selecao da SE16H.", None, set(), ""
            obj.text = panam

            sess.findById("wnd[0]/tbar[1]/btn[8]").press()

            not_found_tokens = ("NO VALUES", "NOT FOUND", "NENHUM", "NAO ENCONTR", "SEM REGIST")
            found_tokens = ("SELECION", "SELECTED", "REGIST", "ENTRAD", "VALUES")
            reported_count = None
            status_msg = ""

            for _ in range(20):
                mt, sb = get_statusbar(sess)
                status_msg = sb or status_msg
                sbar_norm = norm(sb)
                if sbar_norm and any(token in sbar_norm for token in not_found_tokens):
                    return False, sb, 0, set(), ""
                if sbar_norm and any(token in sbar_norm for token in found_tokens):
                    reported_count = _extract_first_int(sb)
                    break

                wnd0 = _safe_find(sess, "wnd[0]")
                wnd_txt = norm(str(getattr(wnd0, "Text", "") or ""))
                if "ENTRADAS ENCONTRADAS" in wnd_txt or "ENTRIES FOUND" in wnd_txt:
                    reported_count = _extract_first_int(sb or wnd_txt)
                    break

                if _se16h_tem_resultados(sess):
                    reported_count = _extract_first_int(sb or wnd_txt)
                    break

                time.sleep(0.2)

            extracted_keys = set()
            grid = _find_first_grid_with_rows(sess)
            if grid:
                try:
                    row_count = int(getattr(grid, "RowCount", 0))
                except Exception:
                    row_count = 0
                field_cols = {
                    "BUKRS": ["BUKRS", "T028P-BUKRS", "V_T028P-BUKRS"],
                    "HBKID": ["HBKID", "T028P-HBKID", "V_T028P-HBKID"],
                    "HKTID": ["HKTID", "T028P-HKTID", "V_T028P-HKTID"],
                    "VGEXT": ["VGEXT", "T028P-VGEXT", "V_T028P-VGEXT"],
                    "VOZPM": ["VOZPM", "T028P-VOZPM", "V_T028P-VOZPM"],
                    "PANAM": ["PANAM", "T028P-PANAM", "V_T028P-PANAM"],
                    "TARGFI": ["TARGFI", "T028P-TARGFI", "V_T028P-TARGFI"],
                    "PREFIX": ["PREFIX", "T028P-PREFIX", "V_T028P-PREFIX"],
                }
                for idx in range(max(0, row_count)):
                    item = _norm_key_tuple(
                        [
                            _grid_get_cell(grid, idx, field_cols["BUKRS"]),
                            _grid_get_cell(grid, idx, field_cols["HBKID"]),
                            _grid_get_cell(grid, idx, field_cols["HKTID"]),
                            _grid_get_cell(grid, idx, field_cols["VGEXT"]),
                            _grid_get_cell(grid, idx, field_cols["VOZPM"]),
                            _grid_get_cell(grid, idx, field_cols["PANAM"]),
                            _grid_get_cell(grid, idx, field_cols["TARGFI"]),
                            _grid_get_cell(grid, idx, field_cols["PREFIX"]),
                        ]
                    )
                    if any(item):
                        extracted_keys.add(item)

            scr = ""
            if extracted_keys or (reported_count and int(reported_count) > 0):
                scr = capturar_print_sap(sess, f"Validacao_T028P_{panam}")

            if extracted_keys:
                return True, status_msg or f"Registos lidos na T028P para PANAM={panam}.", reported_count, extracted_keys, scr

            if reported_count and int(reported_count) > 0:
                return True, status_msg or f"{reported_count} registo(s) encontrado(s) na T028P.", reported_count, set(), scr

            return False, status_msg or "Sem entradas encontradas na T028P.", reported_count or 0, set(), ""
        except Exception as e:
            return None, str(e), None, set(), ""
        finally:
            if not keep_validation_screen:
                try:
                    sess.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                    sess.findById("wnd[0]").sendVKey(0)
                except Exception:
                    pass

    def _validar_cadeia_na_tpama(sess, nome_cadeia, tentativas=3, pausa=1.0):
        ultima_msg = ""
        for _ in range(max(1, int(tentativas))):
            existe, msg, _ = _cadeia_existe_em_tpama(sess, nome_cadeia)
            ultima_msg = msg or ultima_msg
            if existe is True:
                return True, msg or f"Cadeia '{nome_cadeia}' confirmada na TPAMA."
            if existe is None:
                time.sleep(max(0.0, float(pausa)))
                continue
            time.sleep(max(0.0, float(pausa)))
        return False, ultima_msg or f"Cadeia '{nome_cadeia}' nao foi confirmada na TPAMA."

    ###################################################################################
    # REQUEST DE TRANSPORTE RESOLUTION
    ###################################################################################
    def validar_request(valor: str) -> str:
        v = (valor or "").strip().upper().replace(" ", "")
        if not v:
            return ""
        if re.match(r"^[A-Z0-9]{3,4}K\d{6,}$", v):
            return v
        return ""

    def resolver_request_recebida(request_transporte_param, request_ctx_param):
        numero = validar_request(request_transporte_param)
        desc = ""
        if isinstance(request_ctx_param, dict):
            if not numero:
                numero = validar_request(request_ctx_param.get("request_number", ""))
            desc = str(request_ctx_param.get("request_desc", "") or "").strip()
        return numero, desc

    def _extrair_request_do_output(output_texto):
        marker = re.search(r"REQUEST_NUMBER=([A-Z0-9]{3,4}K\d{6,})", output_texto or "")
        if marker:
            return validar_request(marker.group(1))
        fallback = re.search(r"\b([A-Z0-9]{3,4}K\d{6,})\b", output_texto or "")
        if fallback:
            return validar_request(fallback.group(1))
        return ""

    def _garantir_request(sess, motivo):
        nonlocal request_number
        if request_number:
            return request_number

        system_name = str(getattr(sess.Info, "SystemName", "") or SISTEMA_ESPERADO).strip().upper()
        client = CLIENTE_ESPERADO or str(getattr(sess.Info, "Client", "") or "").strip()
        if not system_name or not client:
            raise RuntimeError("Nao foi possivel resolver sistema/cliente para criar request.")

        criar_request_path = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "criar_request.py")
        )
        if not os.path.exists(criar_request_path):
            raise RuntimeError(f"Script de criacao de request nao encontrado: {criar_request_path}")

        descricao = (request_description or "").strip()
        if not os.path.exists(criar_request_path):
            raise RuntimeError(f"Script de criacao de request nao encontrado: {criar_request_path}")

        if not descricao:
            descricao = (motivo or "").strip()
        if not descricao:
            descricao = "Cadeias de Pesquisa"
        comando = [
            sys.executable,
            criar_request_path,
            "--auto-create",
            "--order-type",
            "customizing",
            "--description",
            descricao[:72],
            "--system-name",
            system_name,
            "--client",
            client,
        ]

        run = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        output = f"{run.stdout or ''}\n{run.stderr or ''}"
        req = _extrair_request_do_output(output)
        if run.returncode != 0 or not req:
            raise RuntimeError(
                "Falha ao criar request automaticamente. "
                f"returncode={run.returncode} | output={output.strip()}"
            )

        request_number = req
        print(f"REQUEST_NUMBER={request_number}")
        return request_number

    def localizar_pesquisar_request():
        base_script = os.path.dirname(os.path.abspath(__file__))
        base_pai = os.path.dirname(base_script)
        base_avo = os.path.dirname(base_pai)
        candidatos = [
            os.environ.get("SAP_PESQUISAR_REQUEST_PATH", "").strip(),
            os.path.join(base_script, "pesquisar_request.py"),
            os.path.join(base_pai, "pesquisar_request.py"),
            os.path.join(base_avo, "pesquisar_request.py"),
            os.path.join(os.getcwd(), "pesquisar_request.py"),
        ]
        vistos = set()
        for caminho in candidatos:
            if not caminho:
                continue
            caminho_abs = os.path.abspath(caminho)
            if caminho_abs in vistos:
                continue
            vistos.add(caminho_abs)
            if os.path.exists(caminho_abs) and os.path.isfile(caminho_abs):
                return caminho_abs
        return None

    def carregar_pesquisar_request():
        caminho = localizar_pesquisar_request()
        if not caminho:
            raise FileNotFoundError("Não encontrei o ficheiro pesquisar_request.py.")
        import importlib.util
        spec = importlib.util.spec_from_file_location("pesquisar_request", caminho)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod, caminho

    def escolher_request_por_linha(lista_resultados):
        if not lista_resultados:
            return ("", "")
        while True:
            raw = input(f"\nDigite o número da linha da request (1-{len(lista_resultados)}): ").strip()
            if not raw.isdigit():
                print("❌ Digite apenas números.")
                continue
            idx = int(raw)
            if idx < 1 or idx > len(lista_resultados):
                print("❌ Número fora do intervalo.")
                continue
            trkorr, as4text = lista_resultados[idx - 1]
            trkorr = validar_request(trkorr)
            as4text = "" if as4text is None else str(as4text).strip()
            return (trkorr, as4text)

    def listar_requests_via_modulo():
        mod, caminho_modulo = carregar_pesquisar_request()
        try:
            lista = mod.listar_requests(
                system_name=SISTEMA_ESPERADO,
                max_rows="5000",
                include_requests=False,
                use_new_mode=True,
                minimize=True,
                close_after=True
            )
        except TypeError:
            lista = mod.listar_requests(
                system_name=SISTEMA_ESPERADO,
                max_rows="5000"
            )
        return lista, caminho_modulo

    def perguntar_request_transporte():
        if modo_nao_interativo:
            return "", ""
        print("\n======================================================================")
        print("OPÇÕES DE TRANSPORTE")
        print("1 - Escrever o número da Request")
        print("2 - Pesquisar suas requests criadas")
        print("3 - Prima [Enter] vazio para NÃO selecionar agora")
        print("======================================================================")
        while True:
            opc = input("\nOpção: ").strip()
            if opc in ("1", "2", "3", ""):
                if opc == "" or opc == "3":
                    return "", ""
                if opc == "1":
                    while True:
                        num_raw = input("Número da Request (ex: S4QK900396): ").strip()
                        num = validar_request(num_raw)
                        if num:
                            return num, ""
                        print("❌ Request inválida. Exemplo válido: S4QK900396")
                if opc == "2":
                    try:
                        lista, caminho_modulo = listar_requests_via_modulo()
                        print(f"✅ Módulo de pesquisa carregado: {caminho_modulo}")
                    except Exception as e:
                        print(f"❌ Falha ao pesquisar requests: {e}")
                        continue
                    if not lista:
                        print("⚠️ Nenhuma request encontrada.")
                        continue
                    trkorr, as4text = escolher_request_por_linha(lista)
                    if trkorr:
                        print(f"✅ Request selecionada: {trkorr} | {as4text}")
                        return trkorr, as4text
                    continue
            print("❌ Opção inválida. Use 1, 2, 3 ou apenas pressione Enter.")

    def preencher_request_no_popup(sess, request_numero):
        request_numero = validar_request(request_numero)
        if not request_numero:
            return False
        caminhos_campo = [
            "wnd[1]/usr/ctxtKO008-TRKORR",
            "wnd[1]/usr/ctxtTRWBO_REQUEST-TRKORR",
            "wnd[1]/usr/ctxtE070-TRKORR",
        ]
        for caminho in caminhos_campo:
            try:
                campo = sess.findById(caminho)
                campo.text = request_numero
                try:
                    campo.setFocus()
                except Exception:
                    pass
                try:
                    campo.caretPosition = len(request_numero)
                except Exception:
                    pass
                return True
            except Exception:
                continue
        return False

    def obter_sessao_sap():
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            for conn in application.Children:
                for sess in conn.Children:
                    if str(sess.Info.SystemName).strip().upper() != SISTEMA_ESPERADO:
                        continue
                    if CLIENTE_ESPERADO and str(sess.Info.Client).strip() != CLIENTE_ESPERADO:
                        continue
                    return sess
            return None
        except Exception:
            return None

    def capturar_print_sap(sess, name: str) -> str:
        try:
            if not evidencias_dir:
                return ""
            os.makedirs(evidencias_dir, exist_ok=True)
            path = os.path.join(evidencias_dir, f"{name}.bmp")
            wnd = sess.findById("wnd[0]")
            try:
                wnd.hardCopy(path, 2)
            except Exception:
                try:
                    wnd.HardCopy(path, 2)
                except Exception:
                    try:
                        wnd.hardCopy(path)
                    except Exception:
                        wnd.HardCopy(path)
            if os.path.exists(path):
                return path
        except Exception as e:
            print(f"  ├─ Aviso: Falha ao capturar screenshot: {e}")
        return ""

    ###################################################################################
    # LÓGICA SAP - CRIAR CADEIA
    ###################################################################################
    def criar_cadeia_sap(sess, nome_cadeia):
        nonlocal request_number
        nome_cadeia = (nome_cadeia or "").strip()
        nome_cadeia_limite = nome_cadeia[:20]
        screenshot_path = ""

        try:
            print(f"  |- A validar existencia de '{nome_cadeia_limite}' na TPAMA (SE16H)...")
            existe_tpama, msg_tpama, scr_val = _cadeia_existe_em_tpama(sess, nome_cadeia_limite)
            if existe_tpama is True:
                msg = msg_tpama or f"Cadeia '{nome_cadeia_limite}' ja existe na TPAMA."
                registro_fase1[nome_cadeia] = {"status": "Já existia", "msg": msg, "screenshot": scr_val}
                return {"ok": True, "msg": msg}
            if existe_tpama is None:
                print(f"  |- Aviso: validacao TPAMA indisponivel ({msg_tpama}). Vou seguir com a criacao.")

            if not request_number and modo_nao_interativo:
                print("  |- Sem request no contexto. A criar request automaticamente...")
                _garantir_request(sess, f"Cadeia Pesquisa | {nome_cadeia_limite}")

            print(f"  ├─ A abrir transação OTPM...")
            sess.findById("wnd[0]/tbar[0]/okcd").text = "/NOTPM"
            _send_vkey(sess, 0)

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                registro_fase1[nome_cadeia] = {"status": "Erro", "msg": sb, "screenshot": ""}
                return {"ok": False, "msg": sb}

            print(f"  ├─ A clicar em Criar...")
            if not try_press(sess, ["wnd[0]/tbar[1]/btn[25]", "wnd[0]/tbar[1]/btn[5]"]):
                raise Exception("Não consegui clicar no botão Criar.")

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                registro_fase1[nome_cadeia] = {"status": "Erro", "msg": sb, "screenshot": ""}
                return {"ok": False, "msg": sb}

            print(f"  ├─ A ativar modo de edição...")
            try_press(sess, ["wnd[0]/tbar[1]/btn[5]"])

            campo_nome = _esperar_objeto(sess, "wnd[0]/usr/txtV_TPAMA-PANAM", timeout=3.0)
            campo_desc = _esperar_objeto(sess, "wnd[0]/usr/txtV_TPAMA-NOTE", timeout=3.0)
            campo_regex = _esperar_objeto(sess, "wnd[0]/usr/txtV_TPAMA-REGEX", timeout=3.0)

            if not campo_nome or not campo_desc or not campo_regex:
                raise Exception("Não consegui localizar os campos principais da cadeia no SAP.")

            print(f"  ├─ A preencher dados principais...")
            campo_nome.text = nome_cadeia_limite
            campo_desc.text = nome_cadeia
            campo_regex.text = nome_cadeia
            campo_regex.setFocus()
            try:
                campo_regex.caretPosition = len(nome_cadeia)
            except Exception:
                pass

            # CAPTURA PRINT SCREEN ANTES DE DAR ENTER
            screenshot_path = capturar_print_sap(sess, f"Criar_Cadeia_{nome_cadeia_limite}")

            _send_vkey(sess, 0)

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                registro_fase1[nome_cadeia] = {"status": "Erro", "msg": sb, "screenshot": screenshot_path}
                return {"ok": False, "msg": sb}

            total_linhas_limpeza = max(20, len(nome_cadeia))
            print(f"  ├─ A limpar até {total_linhas_limpeza} linhas da tabela de mapeamento...")
            for i in range(total_linhas_limpeza):
                campo = f"wnd[0]/usr/subSUB_PAMA:SAPLPAMI:0210/tblSAPLPAMITC_MAP/txtT_MAP-MXCHAR[3,{i}]"
                try:
                    obj = sess.findById(campo)
                    obj.text = ""
                except Exception:
                    pass

            _send_vkey(sess, 0)

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A") and sb:
                registro_fase1[nome_cadeia] = {"status": "Erro", "msg": sb, "screenshot": screenshot_path}
                return {"ok": False, "msg": sb}

            print(f"  ├─ A guardar...")
            if not try_press(sess, ["wnd[0]/tbar[0]/btn[11]"]):
                raise Exception("Não consegui clicar em Guardar.")

            popup_req = _safe_find(sess, "wnd[1]")
            if popup_req:
                print(f"  ├─ Popup SAP detetado após guardar...")
                if not request_number and not modo_nao_interativo:
                    print(f"  ├─ Nenhuma request pré-selecionada. A pedir seleção agora...")
                    req_num_p, _ = perguntar_request_transporte()
                    request_number = req_num_p or request_number

                if request_number:
                    print(f"  ├─ A preencher request: {request_number}")
                    preencher_request_no_popup(sess, request_number)

                if not try_press(sess, ["wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[11]"]):
                    fechar_popup_se_existir(sess)

            mt, sb = get_statusbar(sess)
            if mt in ("E", "A"):
                msg = sb or f"Erro ao criar cadeia '{nome_cadeia_limite}'."
                limpar_para_home(sess)
                registro_fase1[nome_cadeia] = {"status": "Erro", "msg": msg, "screenshot": screenshot_path}
                return {"ok": False, "msg": msg}

            msg_criacao = sb or f"Cadeia '{nome_cadeia_limite}' criada com sucesso."
            print("  |- A validar criacao na TPAMA (SE16H)...")
            validada, msg_validacao = _validar_cadeia_na_tpama(sess, nome_cadeia_limite, tentativas=4, pausa=1.0)
            if not validada:
                msg = f"Criacao executada, mas sem confirmacao na TPAMA. Criacao SAP='{msg_criacao}' | Validacao='{msg_validacao}'"
                registro_fase1[nome_cadeia] = {"status": "Erro de Validação", "msg": msg, "screenshot": screenshot_path}
                return {"ok": False, "msg": msg}

            res_msg = msg_validacao or msg_criacao
            registro_fase1[nome_cadeia] = {"status": "Criado com Sucesso", "msg": res_msg, "screenshot": screenshot_path}
            return {"ok": True, "msg": res_msg}

        except Exception as e:
            mt, sb = get_statusbar(sess)
            msg = sb if sb else str(e)
            try:
                fechar_popup_se_existir(sess)
            except Exception:
                pass
            limpar_para_home(sess)
            registro_fase1[nome_cadeia] = {"status": "Erro", "msg": msg, "screenshot": screenshot_path}
            return {"ok": False, "msg": msg}



    ###################################################################################
    # RESOLUÇÃO DOS DADOS DE ENTRADA
    ###################################################################################
    if not caminho_ficheiro:
        caminho_ficheiro = selecionar_ficheiro_excel()
        if not caminho_ficheiro:
            print("❌ Operação cancelada pelo utilizador.")
            return

    if not os.path.exists(caminho_ficheiro):
        print(f"❌ Ficheiro não encontrado: {caminho_ficheiro}")
        return

    excel_dir = os.path.dirname(caminho_ficheiro)
    evidencias_dir = os.path.join(excel_dir, f"Evidencias_Cadeias_{timestamp}")

    df = abrir_excel_para_dataframe(caminho_ficheiro)
    cols_map, faltantes = resolver_colunas(df)
    if faltantes:
        print("❌ Colunas obrigatórias em falta após normalização:")
        for c in faltantes:
            print(f"   - {c}")
        print("🔍 Cabeçalhos detetados:", ", ".join(df.columns))
        return

    C_NOME = cols_map["NOME CADEIA DE PESQUISA"]
    C_EMP = cols_map["EMPRESA"]
    C_BANCO = cols_map["BANCO"]
    C_IDCT = cols_map["ID CONTA"]
    C_OPEXT = cols_map["OPERAÇÃO EXTERNA"]
    C_SINAL = cols_map["SINAL (+/-)"]
    C_NDEST = cols_map["NOME DO CAMPO DESTINO"]
    C_DEST = cols_map["CAMPO DESTINO"]
    C_BASE = cols_map["BASE DE MAPEAMENTO"]

    def _row_t028p_key(row):
        return {
            "BUKRS": safe_value(row[C_EMP]),
            "HBKID": safe_value(row[C_BANCO]),
            "HKTID": safe_value(row[C_IDCT]),
            "PANAM": safe_value(row["PANAM_EFETIVO"]),
            "TARGFI": safe_value(row[C_DEST]),
            "VGEXT": safe_value(row[C_OPEXT]),
            "VOZPM": safe_value(row[C_SINAL]),
            "PREFIX": safe_value(row[C_BASE]),
        }

    def _key_to_label(key):
        return (
            f"BUKRS={key['BUKRS']}|HBKID={key['HBKID']}|HKTID={key['HKTID']}|"
            f"VGEXT={key['VGEXT']}|VOZPM={key['VOZPM']}|PANAM={key['PANAM']}|"
            f"TARGFI={key['TARGFI']}|PREFIX={key['PREFIX']}"
        )

    def _subset_exists_in_t028p(session, subset):
        expected_by_panam = {}
        for _, row in subset.iterrows():
            key = _row_t028p_key(row)
            panam = key.get("PANAM", "")
            key_tuple = tuple(
                key.get(k, "")
                for k in ("BUKRS", "HBKID", "HKTID", "VGEXT", "VOZPM", "PANAM", "TARGFI", "PREFIX")
            )
            if not panam:
                continue
            expected_by_panam.setdefault(panam, set()).add(key_tuple)

        if not expected_by_panam:
            return False, "Nao foi possivel resolver PANAM para validacao na T028P.", []

        missing = []
        details = []
        total_expected = 0
        screenshots = []
        for panam, keys in expected_by_panam.items():
            expected_set = set(
                _norm_key_tuple(
                    [
                        k[0],  # BUKRS
                        k[1],  # HBKID
                        k[2],  # HKTID
                        k[3],  # VGEXT
                        k[4],  # VOZPM
                        k[5],  # PANAM
                        k[6],  # TARGFI
                        k[7],  # PREFIX
                    ]
                )
                for k in keys
            )
            expected_count = len(expected_set)
            total_expected += expected_count
            exists, msg, qtd, extracted, scr = _t028p_fetch_keys_by_panam(session, panam)
            if exists is None:
                return None, msg or f"Validacao T028P indisponivel para PANAM={panam}.", []
            if scr:
                screenshots.append(scr)
            if not exists:
                missing.append(f"PANAM={panam} (esperado>={expected_count})")
                details.append(f"{panam}: esperado={expected_count} encontrado=0")
                continue

            if extracted:
                missing_keys = [item for item in expected_set if item not in extracted]
                if missing_keys:
                    missing.append(f"PANAM={panam} (faltam {len(missing_keys)} chave(s))")
                details.append(
                    f"{panam}: esperado={expected_count} encontrado={len(extracted)}"
                )
            else:
                if qtd is not None and int(qtd) < int(expected_count):
                    missing.append(f"PANAM={panam} (esperado>={expected_count}, encontrado={qtd})")
                details.append(f"{panam}: esperado>={expected_count} encontrado={qtd if qtd is not None else '?'}")

        if missing:
            preview = "; ".join(missing[:3])
            return False, f"Entradas nao encontradas na T028P ({len(missing)}): {preview}", screenshots

        return True, f"Validacao T028P por PANAM concluida ({total_expected} entrada(s) esperadas). " + "; ".join(details[:3]), screenshots

    def _confirm_subset_in_t028p(sess, subset, tentativas=4, pausa=1.0):
        last_msg = ""
        for _ in range(max(1, int(tentativas))):
            exists, msg, scrs = _subset_exists_in_t028p(sess, subset)
            last_msg = msg or last_msg
            if exists is True:
                return True, msg or "Subconjunto confirmado na T028P.", scrs
            time.sleep(max(0.0, float(pausa)))
        return False, last_msg or "Nao foi possivel confirmar o subconjunto na T028P.", []

    if "STATUS" not in df.columns:
        df["STATUS"] = ""
    if "MSG" not in df.columns:
        df["MSG"] = ""

    # Filtrar registos pendentes
    if called_by_main:
        df_filtrado = df.copy()
    else:
        df_filtrado = df[df["STATUS"].astype(str).map(norm) != "CONCLUIDO"].copy()

    if df_filtrado.empty:
        print("✅ Nenhuma linha para processar. Tudo concluído.")
        return

    df_filtrado = df_filtrado.apply(lambda x: x.astype(str).apply(safe_value))
    df_filtrado["PANAM_EFETIVO"] = df_filtrado[C_NOME].apply(lambda x: sanitize_paname(x, 20))

    # Identificar chaves únicas para a criação de cadeias
    nomes_unicos = df_filtrado[C_NOME].unique().tolist()

    if not nomes_unicos:
        print("✅ Nenhuma cadeia nova para validar ou criar.")
        return

    # ----- Resolução do Request -----
    request_number, request_desc = resolver_request_recebida(request_transporte, request_ctx)
    if request_number:
        print(f"✅ Request selecionada: {request_number}")
    else:
        print("ℹ️ Nenhuma request foi pré-selecionada. Se o SAP pedir, será tratado no momento do popup.")

    # ----- Conexão SAP -----
    session = obter_sessao_sap()
    if not session:
        print(f"❌ Não encontrei sessão ativa do ambiente '{ambiente_cockpit}' ({SISTEMA_ESPERADO}).")
        return

    print(f"✅ Sessão SAP encontrada: {session.Info.SystemName} | User: {session.Info.User} | Cliente: {session.Info.Client}")
    aplicar_modo_janela_sap(session)

    ###################################################################################
    # FASE 1: CRIAR CADEIAS DE PESQUISA EM FALTA
    ###################################################################################
    print("\n======================================================================")
    print("FASE 1: CRIAÇÃO DE CADEIAS DE PESQUISA")
    print("======================================================================")

    cadeias_criadas_com_sucesso = set()
    cadeias_erros = {}

    for idx, nome_cadeia in enumerate(nomes_unicos, start=1):
        print(f"▶ [{idx}/{len(nomes_unicos)}] A processar cadeia: {nome_cadeia}")
        resultado = criar_cadeia_sap(session, nome_cadeia)
        if resultado["ok"]:
            print(f"  ├─ Sucesso: {resultado['msg']}")
            cadeias_criadas_com_sucesso.add(nome_cadeia)
        else:
            print(f"  ├─ Erro: {resultado['msg']}")
            cadeias_erros[nome_cadeia] = resultado["msg"]
            # Marcar erro na tabela original para estas linhas
            mask_cadeia = df[C_NOME].astype(str).str.strip() == nome_cadeia.strip()
            df.loc[mask_cadeia, "STATUS"] = "Erro na criação"
            df.loc[mask_cadeia, "MSG"] = resultado["msg"]

    ###################################################################################
    # FASE 2: ATRIBUIR CADEIAS DE PESQUISA (T028P)
    ###################################################################################
    print("\n======================================================================")
    print("FASE 2: ATRIBUIÇÃO DE CADEIAS DE PESQUISA")
    print("======================================================================")

    # Excluir linhas onde a criação da cadeia falhou na Fase 1
    df_filtrado_atrib = df_filtrado[df_filtrado[C_NOME].isin(cadeias_criadas_com_sucesso)].copy()

    if df_filtrado_atrib.empty:
        print("⚠️ Sem atribuições para processar (todas as cadeias falharam ou não existiam).")
        # Salvar o Excel com os erros reportados na Fase 1
        df.to_excel(caminho_ficheiro, index=False)
        return

    # Construção de blocos baseado nas regras
    col_nome_dest_norm = df_filtrado_atrib[C_NDEST].apply(norm)

    is_bp = col_nome_dest_norm == "BP"
    is_centro = col_nome_dest_norm == "CENTRO"
    is_centro_lucro = col_nome_dest_norm == "CENTRO DE LUCRO"
    is_regra_cont = col_nome_dest_norm.str.contains("REGRA DE CONTABILIZACAO")
    is_xref1 = col_nome_dest_norm == "CHAVE REF 1"

    df_bp = df_filtrado_atrib[is_bp].copy()
    df_centro = df_filtrado_atrib[is_centro].copy()
    df_centro_lucro = df_filtrado_atrib[is_centro_lucro].copy()
    df_regra_cont = df_filtrado_atrib[is_regra_cont].copy()
    df_xref1 = df_filtrado_atrib[is_xref1].copy()
    df_outros = df_filtrado_atrib[~(is_bp | is_centro | is_centro_lucro | is_regra_cont | is_xref1)].copy()

    blocos = []

    # 1. OUTROS
    if not df_outros.empty:
        blocos.append(df_outros)

    # 2. BP -> EBAVKOA (+/-) / EBAVKON (+/-)
    if not df_bp.empty:
        dd_keys_with_D = set()
        if not df_regra_cont.empty:
            df_regra_cont_chk = df_regra_cont.copy()
            df_regra_cont_chk[C_BASE] = series_to_upper_clean(df_regra_cont_chk[C_BASE])
            mask_dd = df_regra_cont_chk[C_BASE].isin({"ZDD5", "ZDD6"})
            if mask_dd.any():
                for _, r in df_regra_cont_chk[mask_dd].iterrows():
                    dd_keys_with_D.add((
                        safe_value(r[C_NOME]),
                        safe_value(r[C_EMP]),
                        safe_value(r[C_BANCO]),
                        safe_value(r[C_IDCT]),
                        safe_value(r[C_OPEXT]),
                    ))
        for _, grupo in df_bp.groupby([C_NOME, C_EMP, C_BANCO, C_IDCT, C_OPEXT], dropna=False):
            t = grupo.iloc[0].copy()
            grp_key = (safe_value(t[C_NOME]), safe_value(t[C_EMP]), safe_value(t[C_BANCO]), safe_value(t[C_IDCT]), safe_value(t[C_OPEXT]))
            base_original = t[C_BASE]
            prefix_for_k = "D" if grp_key in dd_keys_with_D else "K"

            k = t.copy()
            k[C_DEST] = "EBAVKOA"
            k[C_BASE] = prefix_for_k

            v = t.copy()
            v[C_DEST] = "EBAVKON"
            v[C_BASE] = base_original

            blocos.append(pd.DataFrame(
                [dict(k, **{C_SINAL: "+"}), dict(k, **{C_SINAL: "-"}),
                 dict(v, **{C_SINAL: "+"}), dict(v, **{C_SINAL: "-"})],
                index=[t.name] * 4
            ))

    # 3. CENTRO -> EBFNAM1 / EBFVAL1
    if not df_centro.empty:
        for _, grupo in df_centro.groupby([C_NOME, C_EMP, C_BANCO, C_IDCT, C_OPEXT], dropna=False):
            t = grupo.iloc[0].copy()
            base_original = t[C_BASE]

            fx = t.copy()
            fx[C_DEST] = "EBFNAM1"
            fx[C_BASE] = "COBL-WERKS"

            vr = t.copy()
            vr[C_DEST] = "EBFVAL1"
            vr[C_BASE] = base_original

            blocos.append(pd.DataFrame(
                [dict(fx, **{C_SINAL: "+"}), dict(fx, **{C_SINAL: "-"}),
                 dict(vr, **{C_SINAL: "+"}), dict(vr, **{C_SINAL: "-"})],
                index=[t.name] * 4
            ))

    # 4. CHAVE REF 1 -> EBFNAM2 / EBFVAL2
    if not df_xref1.empty:
        for _, grupo in df_xref1.groupby([C_NOME, C_EMP, C_BANCO, C_IDCT, C_OPEXT], dropna=False):
            t = grupo.iloc[0].copy()
            base_original = t[C_BASE]

            fx = t.copy()
            fx[C_DEST] = "EBFNAM2"
            fx[C_BASE] = "BSEG-XREF1"

            vr = t.copy()
            vr[C_DEST] = "EBFVAL2"
            vr[C_BASE] = base_original

            blocos.append(pd.DataFrame(
                [dict(fx, **{C_SINAL: "+"}), dict(fx, **{C_SINAL: "-"}),
                 dict(vr, **{C_SINAL: "+"}), dict(vr, **{C_SINAL: "-"})],
                index=[t.name] * 4
            ))

    # 5. CENTRO DE LUCRO -> EBPRCTR
    if not df_centro_lucro.empty:
        for _, grupo in df_centro_lucro.groupby([C_NOME, C_EMP, C_BANCO, C_IDCT, C_OPEXT], dropna=False):
            t = grupo.iloc[0].copy()
            t[C_DEST] = "EBPRCTR"
            blocos.append(pd.DataFrame(
                [dict(t, **{C_SINAL: "+"}), dict(t, **{C_SINAL: "-"})],
                index=[t.name] * 2
            ))

    # 6. REGRA DE CONTABILIZAÇÃO (EBVGINT)
    if not df_regra_cont.empty:
        tmp = df_regra_cont.copy()
        tmp[C_DEST] = "EBVGINT"
        def decide_sign_from_base(base_val, sinal_ficheiro):
            b = safe_value(base_val).upper()
            allowed = BASE_TO_ALLOWED_SIGNS.get(b)
            if allowed:
                if len(allowed) == 1:
                    return next(iter(allowed))
                sf = map_sign_free(sinal_ficheiro)
                return sf if (sf in allowed and sf) else "+"
            sf = map_sign_free(sinal_ficheiro)
            return sf if sf else "+"
        tmp[C_SINAL] = [decide_sign_from_base(b, s) for b, s in zip(tmp[C_BASE], tmp[C_SINAL])]
        blocos.append(tmp)

    if not blocos:
        print("❌ Nenhum grupo válido encontrado para atribuição.")
        df.to_excel(caminho_ficheiro, index=False)
        return

    df_final = pd.concat(blocos, ignore_index=False)
    CHAVE = [C_EMP, C_BANCO, C_IDCT, C_OPEXT, C_SINAL, "PANAM_EFETIVO", C_DEST, C_BASE]
    df_final = df_final[(df_final[C_DEST].astype(str) != "") & (df_final[C_BASE].astype(str) != "")]
    df_final = df_final.drop_duplicates(subset=CHAVE, keep="first").reset_index()

    # Executar Atribuição no SAP
    TBL = "wnd[0]/usr/tblSAPLPAMVTCTRL_V_T028P"

    for funcao in df_final[C_NOME].unique():
        subset = df_final[df_final[C_NOME] == funcao]
        print(f"\n▶ A atribuir função '{funcao}' ({len(subset)} linha(s))")

        print("  |- A validar existencia na T028P (SE16H)...")
        pre_exists, pre_msg, pre_scrs = _subset_exists_in_t028p(session, subset)
        if pre_exists is True:
            msg_done = pre_msg or "Entradas ja existem na T028P."
            df.loc[subset["index"], "STATUS"] = "CONCLUIDO"
            df.loc[subset["index"], "MSG"] = msg_done
            print(f"  |- CONCLUIDO sem lancamento: {msg_done}")
            scr_val = pre_scrs[0] if pre_scrs else ""
            registro_fase2[funcao] = {
                "status": "Já existia",
                "msg": msg_done,
                "screenshot": scr_val,
                "linhas": subset.to_dict(orient="records")
            }
            continue
        if pre_exists is None:
            print(f"  |- Aviso: validacao inicial T028P indisponivel ({pre_msg}). Vou seguir com o lancamento.")

        if not request_number:
            print("  |- Sem request no contexto. A criar request automaticamente...")
            _garantir_request(session, f"Atribuir Cadeia Pesquisa | {funcao}")

        screenshot_path = ""
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nOTPM"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/shellcont/shell").selectItem("02", "Column1")
            session.findById("wnd[0]/shellcont/shell").doubleClickItem("02", "Column1")
            session.findById("wnd[0]/tbar[1]/btn[25]").press()
            session.findById("wnd[0]/tbar[1]/btn[5]").press()

            total_linhas_lote = len(subset)
            for i, (_, row) in enumerate(subset.iterrows()):
                set_cell_any(session, TBL, "V_T028P-BUKRS", 0, i, safe_value(row[C_EMP]))
                set_cell_any(session, TBL, "V_T028P-HBKID", 1, i, safe_value(row[C_BANCO]))
                set_cell_any(session, TBL, "V_T028P-HKTID", 2, i, safe_value(row[C_IDCT]))
                set_cell_any(session, TBL, "V_T028P-VGEXT", 3, i, safe_value(row[C_OPEXT]))
                set_cell_any(session, TBL, "V_T028P-VOZPM", 4, i, safe_value(row[C_SINAL]))

                targfi = safe_value(row[C_DEST])
                ok_combo = set_combo_key_then_check(session, TBL, "V_T028P-TARGFI", 7, i, targfi)
                if not ok_combo:
                    lido = get_cell_text(session, TBL, "V_T028P-TARGFI", 7, i)
                    raise RuntimeError(f"Nao foi possivel definir TARGFI='{targfi}' (ficou '{lido}')")

                set_cell_any(session, TBL, "V_T028P-PREFIX", 9, i, safe_value(row[C_BASE]))
                panam_val = safe_value(row["PANAM_EFETIVO"])
                set_cell_any(session, TBL, "V_T028P-PANAM", 6, i, panam_val)

                try:
                    session.findById(f"{TBL}/chkV_T028P-ENABLED[8,{i}]").selected = True
                except Exception:
                    pass

            # CAPTURA PRINT SCREEN ANTES DE DAR ENTER
            screenshot_path = capturar_print_sap(session, f"Atribuir_Cadeia_{sanitize_paname(funcao, 12)}")

            print(f"  |- Lote preenchido ({total_linhas_lote} linha(s)). A validar consistencia no SAP...")
            session.findById("wnd[0]").sendVKey(0)
            err = sbar_error(session)
            if err:
                raise RuntimeError(f"Erro de validacao do lote '{funcao}': {err}")

            # Grava + TRKORR
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = request_number
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            msg_gravacao = session.findById("wnd[0]/sbar").Text.strip()
            print("  |- A validar gravacao na T028P (SE16H)...")
            validado, msg_validacao, scrs_val = _confirm_subset_in_t028p(session, subset, tentativas=4, pausa=1.0)
            if not validado:
                raise RuntimeError(f"Lancamento executado sem confirmacao na T028P. Gravacao SAP='{msg_gravacao}' | Validacao='{msg_validacao}'")

            msg_final = msg_validacao or msg_gravacao
            df.loc[subset["index"], "STATUS"] = "CONCLUIDO"
            df.loc[subset["index"], "MSG"] = msg_final
            print(f"  |- CONCLUIDO: {msg_final}")

            registro_fase2[funcao] = {
                "status": "Atribuído com Sucesso",
                "msg": msg_final,
                "screenshot": screenshot_path,
                "linhas": subset.to_dict(orient="records")
            }

        except Exception as e:
            msg_erro = f"Erro ao lancar '{funcao}': {e}"
            df.loc[subset["index"], "STATUS"] = "Erro no processamento"
            df.loc[subset["index"], "MSG"] = msg_erro
            print(f"ERRO: {msg_erro}")

            registro_fase2[funcao] = {
                "status": "Erro",
                "msg": msg_erro,
                "screenshot": screenshot_path,
                "linhas": subset.to_dict(orient="records")
            }

    # Salvar o Excel com todas as alterações
    def gerar_documento_configuracao():
        if not registro_fase1 and not registro_fase2:
            return
        
        md_path = os.path.join(excel_dir, f"Configuracao_Cadeias_{timestamp}.md")
        try:
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(f"# Documento de Configuração e Validação SAP - Cadeias de Pesquisa\n\n")
                f.write(f"**Data/Hora:** {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"**Ambiente:** {ambiente_cockpit} ({SISTEMA_ESPERADO})\n")
                f.write(f"**Cliente/Mandante:** {CLIENTE_ESPERADO or '-'}\n")
                if request_number:
                    f.write(f"**Ordem de Transporte (Request):** {request_number}\n")
                f.write(f"\n## Transações Utilizadas\n")
                f.write(f"1. **OTPM**: Manutenção de cadeias de pesquisa (TPAMA) e atribuição de regras (T028P).\n")
                f.write(f"2. **SE16H**: Consulta direta às tabelas para validação física da gravação.\n\n")
                
                if registro_fase1:
                    f.write(f"## 1. Configuração de Cadeias de Pesquisa (TPAMA)\n")
                    for c_nome, info_c in registro_fase1.items():
                        f.write(f"### Cadeia: {c_nome}\n")
                        f.write(f"- **Estado:** {info_c['status']}\n")
                        f.write(f"- **Mensagem:** {info_c['msg']}\n")
                        if info_c['screenshot']:
                            rel_path = os.path.basename(info_c['screenshot'])
                            label_print = "Evidência de validação (existente no SAP)" if info_c['status'] == "Já existia" else "Screenshot antes de gravar"
                            f.write(f"- **{label_print}:** [BMP](Evidencias_Cadeias_{timestamp}/{rel_path})\n")
                        f.write(f"\n")
                
                if registro_fase2:
                    f.write(f"## 2. Atribuição de Cadeias (T028P)\n")
                    for grp_nome, info_a in registro_fase2.items():
                        f.write(f"### Grupo/Função: {grp_nome}\n")
                        f.write(f"- **Estado:** {info_a['status']}\n")
                        f.write(f"- **Mensagem:** {info_a['msg']}\n")
                        if info_a['screenshot']:
                            rel_path = os.path.basename(info_a['screenshot'])
                            label_print = "Evidência de validação (existente no SAP)" if info_a['status'] == "Já existia" else "Screenshot antes de gravar"
                            f.write(f"- **{label_print}:** [BMP](Evidencias_Cadeias_{timestamp}/{rel_path})\n")
                        
                        f.write(f"\n| Empresa | Banco | Conta | Op. Externa | Sinal | Alvo (TARGFI) | Base (PREFIX) |\n")
                        f.write(f"|---|---|---|---|---|---|---|\n")
                        for r in info_a['linhas']:
                            f.write(f"| {r.get(C_EMP, '')} | {r.get(C_BANCO, '')} | {r.get(C_IDCT, '')} | {r.get(C_OPEXT, '')} | {r.get(C_SINAL, '')} | {r.get(C_DEST, '')} | {r.get(C_BASE, '')} |\n")
                        f.write(f"\n")
            print(f"  ├─ Documento Markdown gerado: {md_path}")
        except Exception as e:
            print(f"  ├─ Aviso: Erro ao gerar Markdown: {e}")

        docx_path = os.path.join(excel_dir, f"Configuracao_Cadeias_{timestamp}.docx")
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = None
            doc = None
            try:
                app = win32com.client.DispatchEx("Word.Application")
                app.Visible = False
                doc = app.Documents.Add()
                sel = app.Selection
                
                def _line(text: str = "", bold=False):
                    sel.Font.Bold = bold
                    sel.TypeText(str(text))
                    sel.TypeParagraph()
                
                _line("Documento de Configuração e Validação SAP - Cadeias de Pesquisa", bold=True)
                _line(f"Data/Hora: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                _line(f"Ambiente: {ambiente_cockpit} ({SISTEMA_ESPERADO})")
                _line(f"Cliente/Mandante: {CLIENTE_ESPERADO or '-'}")
                if request_number:
                    _line(f"Ordem de Transporte (Request): {request_number}")
                _line()
                
                _line("Transações Utilizadas:", bold=True)
                _line("1. OTPM - Manutenção de Cadeias de Pesquisa (TPAMA) e Regras (T028P)")
                _line("2. SE16H - Consulta direta para verificação e validação física")
                _line()
                
                if registro_fase1:
                    _line("1. Configuração de Cadeias de Pesquisa (TPAMA)", bold=True)
                    for c_nome, info_c in registro_fase1.items():
                        _line(f"Cadeia: {c_nome}", bold=True)
                        _line(f"Estado: {info_c['status']}")
                        _line(f"Mensagem: {info_c['msg']}")
                        scr = info_c['screenshot']
                        if scr and os.path.exists(scr):
                            try:
                                label_print = "Evidência de validação (existente no SAP)" if info_c['status'] == "Já existia" else "Screenshot antes de gravar"
                                _line(label_print)
                                sel.InlineShapes.AddPicture(FileName=str(os.path.abspath(scr)), LinkToFile=False, SaveWithDocument=True)
                                sel.TypeParagraph()
                            except Exception as e:
                                _line(f"Falha ao adicionar imagem ao Word: {e}")
                        else:
                            _line("Captura de ecrã: Não aplicável ou já existente")
                        _line()
                
                if registro_fase2:
                    _line("2. Atribuição de Cadeias de Pesquisa (T028P)", bold=True)
                    for grp_nome, info_a in registro_fase2.items():
                        _line(f"Grupo / Função: {grp_nome}", bold=True)
                        _line(f"Estado: {info_a['status']}")
                        _line(f"Mensagem: {info_a['msg']}")
                        scr = info_a['screenshot']
                        if scr and os.path.exists(scr):
                            try:
                                label_print = "Evidência de validação (existente no SAP)" if info_a['status'] == "Já existia" else "Screenshot antes de gravar"
                                _line(label_print)
                                sel.InlineShapes.AddPicture(FileName=str(os.path.abspath(scr)), LinkToFile=False, SaveWithDocument=True)
                                sel.TypeParagraph()
                            except Exception as e:
                                _line(f"Falha ao adicionar imagem ao Word: {e}")
                        else:
                            _line("Captura de ecrã: Não aplicável ou já existente")
                        _line()
                
                doc.SaveAs(str(os.path.abspath(docx_path)), FileFormat=12)
                print(f"  ├─ Documento Word (.docx) gerado com sucesso: {docx_path}")
            finally:
                if doc is not None:
                    doc.Close(False)
                if app is not None:
                    app.Quit()
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"  ├─ Aviso: Word não disponível ou erro na geração do docx: {e}")

    try:
        df.to_excel(caminho_ficheiro, index=False)
        print(f"\n💾 Ficheiro de controlo atualizado com sucesso: {caminho_ficheiro}")
    except Exception as e:
        print(f"❌ Erro ao guardar o ficheiro de controlo: {e}")

    try:
        gerar_documento_configuracao()
    except Exception as e:
        print(f"⚠️ Erro ao gerar documentação: {e}")

    if request_number:
        print(f"REQUEST_NUMBER={request_number}")
    print("🔁 Fim.")
    return True


if __name__ == "__main__":
    import argparse
    import os

    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"], required=True)
    parser.add_argument("--request", help="Request opcional para execução direta fora do cockpit.")
    parser.add_argument("--request-description", default="", help="Descricao da request.")
    parser.add_argument("--xlsx")
    parser.add_argument("--auto", action="store_true")
    parser.add_argument("--no-confirm", action="store_true")
    parser.add_argument("--client")
    parser.add_argument("--from-main", action="store_true")
    args = parser.parse_args()

    from_main_cli = bool(args.from_main) or str(os.getenv("SAP_CALLED_BY_MAIN", "")).strip().lower() in {
        "1", "true", "yes", "on", "sim", "s"
    }

    if from_main_cli and not args.xlsx:
        raise SystemExit("Erro: em modo --from-main informe --xlsx.")
    if args.auto and not args.xlsx:
        raise SystemExit("Erro: em modo --auto, informe tambem --xlsx.")

    ctx = {
        "request_option": "1" if args.request else "4",
        "request_number": (args.request or "").strip().upper(),
        "request_desc": "",
        "search_text": "",
    }

    executar(
        ambiente_cockpit=args.ambiente,
        request_ctx=ctx,
        request_transporte=args.request,
        request_description=args.request_description,
        caminho_ficheiro=args.xlsx,
        modo_nao_interativo=(bool(args.auto) or from_main_cli),
        pedir_confirmacao=((not args.no_confirm) and (not from_main_cli)),
        cliente_sap=args.client,
        chamado_pelo_main=from_main_cli,
    )
