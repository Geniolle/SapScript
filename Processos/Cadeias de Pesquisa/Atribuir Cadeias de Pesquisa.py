# -*- coding: utf-8 -*-
###################################################################################
# SCRIPT: LanÃ§ar Cadeias de Pesquisa (OTPM) â€” EBVGINT, PANAM 20 chars
# VERSÃƒO ATUALIZADA:
# 1. Z062 Ã© estritamente POSITIVO (+).
# 2. Z063 Ã© estritamente NEGATIVO (-).
# 3. Inclui lÃ³gica "CHAVE REF 1" -> EBFNAM2/EBFVAL2.
# 4. NÃ£o pede mais a request por input(); usa o processo padrÃ£o de request do cockpit.
###################################################################################

def executar(
    ambiente_cockpit,
    request_ctx,                 # OBRIGATRIO: fora o cockpit a chamar o processo da request
    request_transporte=None,
    request_description="",
    caminho_ficheiro=None,
    modo_nao_interativo=False,
    pedir_confirmacao=True,
    cliente_sap=None,
    chamado_pelo_main=False
):
    # BLOCO 1: IMPORTAÃ‡Ã•ES / CONFIG GERAL
    ###################################################################################
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
    if called_by_main:
        modo_nao_interativo = True
        pedir_confirmacao = False
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    SISTEMA_DESEJADO = MAPA_SISTEMA.get(ambiente_cockpit)
    CLIENTE_ESPERADO = str(cliente_sap or "").strip()

    if not SISTEMA_DESEJADO:
        print(f"âŒ Ambiente invÃ¡lido: {ambiente_cockpit}")
        return

    ###################################################################################
    # BLOCO 2: HELPERS / UTILITÃRIOS
    ###################################################################################
    def norm(s: str) -> str:
        """Remove acentos, normaliza espaÃ§os, trim e upper (para strings simples)."""
        if s is None:
            return ""
        s = str(s)
        s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
        s = s.replace("\u00A0", " ")
        s = re.sub(r"\s+", " ", s.strip())
        return s.upper()

    def safe_value(val):
        """Converte para string limpa; vazio quando NaN/None/'nan'."""
        return "" if pd.isna(val) or str(val).strip().lower() in {"nan", "none"} else str(val).strip()

    def series_to_upper_clean(s: pd.Series) -> pd.Series:
        """NormalizaÃ§Ã£o robusta para Series: NaN-safe, strip e upper."""
        return s.fillna("").astype(str).str.strip().str.upper()

    def map_sign_free(val: str) -> str:
        """Normaliza sinais de livre-tecla para '+' ou '-'."""
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

    def validar_request(valor: str) -> str:
        v = (valor or "").strip().upper().replace(" ", "")
        if not v:
            return ""
        if re.match(r"^[A-Z0-9]{3,4}K\d{6,}$", v):
            return v
        return ""

    def resolver_request_recebida(request_transporte_param, request_ctx_param):
        """
        Prioridade:
        1) request_transporte recebido diretamente
        2) request_ctx['request_number']
        """
        numero = validar_request(request_transporte_param)

        desc = ""
        if isinstance(request_ctx_param, dict):
            if not numero:
                numero = validar_request(request_ctx_param.get("request_number", ""))
            desc = str(request_ctx_param.get("request_desc", "") or "").strip()

        return numero, desc

    # ----- Tabela fixa BASE -> sinal (apenas para EBVGINT) -----
    # ATENÃ‡ÃƒO: Z062 fixo em + e Z063 fixo em -
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

    # ----- cabeÃ§alhos com sinÃ³nimos -----
    COL_NECESSARIAS = {
        "NOME CADEIA DE PESQUISA": {"NOME CADEIA DE PESQUISA", "NOME DA CADEIA DE PESQUISA", "NOME CADEIA PESQUISA"},
        "EMPRESA": {"EMPRESA", "BUKRS", "COMPANY CODE"},
        "BANCO": {"BANCO", "HBKID", "ID BANCO"},
        "ID CONTA": {"ID CONTA", "HKTID", "ID DA CONTA", "ID-CONTA"},
        "OPERAÃ‡ÃƒO EXTERNA": {"OPERAÃ‡ÃƒO EXTERNA", "OPERACAO EXTERNA", "VGEXT", "OP EXTERNA"},
        "SINAL (+/-)": {"SINAL (+/-)", "SINAL", "VOZPM", "SINAL +-"},
        "NOME DO CAMPO DESTINO": {"NOME DO CAMPO DESTINO", "NOME CAMPO DESTINO", "ALVO", "DESTINO"},
        "CAMPO DESTINO": {"CAMPO DESTINO", "TARGFI", "CAMPO DE DESTINO"},
        "BASE DE MAPEAMENTO": {"BASE DE MAPEAMENTO", "PREFIX", "BASE MAPEAMENTO"},
    }

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

    def selecionar_ficheiro_excel() -> str:
        print("ðŸ“ Selecione o ficheiro Excel (janela foi colocada em primeiro plano)...")
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
        # VersÃ£o segura para avisar se o ficheiro estiver aberto
        try:
            return pd.read_excel(caminho, sheet_name="Folha2", dtype=str).fillna("")
        except PermissionError:
            print("\nâŒ ERRO: O ficheiro Excel estÃ¡ ABERTO.")
            print("ðŸ‘‰ Por favor, feche o ficheiro Excel e execute o script novamente.")
            sys.exit()
        except Exception:
            try:
                return pd.read_excel(caminho, dtype=str).fillna("")
            except PermissionError:
                print("\nâŒ ERRO: O ficheiro Excel estÃ¡ ABERTO.")
                print("ðŸ‘‰ Por favor, feche o ficheiro Excel e execute o script novamente.")
                sys.exit()

    def sanitize_paname(v: str, maxlen: int = 20) -> str:
        """Nome da cadeia (PANAM) normalizado para CHAR20."""
        vv = safe_value(v)
        vv = "".join(c for c in unicodedata.normalize("NFKD", vv) if not unicodedata.combining(c))
        vv = re.sub(r"\s+", " ", vv.strip())
        vv = re.sub(r"[^A-Za-z0-9_\-\.\/ ]", "", vv)
        vv = vv.upper()
        return vv[:maxlen]

    # ----- helpers SAP -----
    def set_cell_any(session, table_base_id, field_name, col_index, row_index, value):
        value = "" if value is None else str(value)
        variants = [
            (f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]", "key",  value),
            (f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]", "text", value),
            (f"{table_base_id}/txt{field_name}[{col_index},{row_index}]",  "text", value),
        ]
        for ctrl_id, prop, val in variants:
            try:
                ctrl = session.findById(ctrl_id)
                setattr(ctrl, prop, val)
                return True
            except Exception:
                continue
        return False

    def get_cell_text(session, table_base_id, field_name, col_index, row_index):
        for ctrl_id, attr in [
            (f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]", "key"),
            (f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]", "text"),
            (f"{table_base_id}/txt{field_name}[{col_index},{row_index}]",  "text"),
        ]:
            try:
                ctrl = session.findById(ctrl_id)
                return getattr(ctrl, attr)
            except Exception:
                continue
        return ""

    def set_combo_key_then_check(session, table_base_id, field_name, col_index, row_index, key_expected):
        """Define .key numa combo, aguarda e confirma."""
        try:
            ctrl = session.findById(f"{table_base_id}/cmb{field_name}[{col_index},{row_index}]")
            ctrl.key = str(key_expected)
            time.sleep(0.1)
            ok = (ctrl.key or "").strip().upper() == str(key_expected).strip().upper()
            if not ok:
                ctrl = session.findById(f"{table_base_id}/ctxt{field_name}[{col_index},{row_index}]")
                ctrl.text = str(key_expected)
                time.sleep(0.1)
                ok = (ctrl.text or "").strip().upper() == str(key_expected).strip().upper()
            return ok
        except Exception:
            return set_cell_any(session, table_base_id, field_name, col_index, row_index, key_expected)

    def sbar_error(session) -> str:
        try:
            bar = session.findById("wnd[0]/sbar")
            if getattr(bar, "MessageType", "") in ("E", "A"):
                return bar.Text.strip()
        except Exception:
            pass
        return ""

    def _safe_find(session, sap_id):
        try:
            return session.findById(sap_id)
        except Exception:
            return None

    def _set_text_any(session, ids, value):
        for cid in ids:
            try:
                obj = session.findById(cid)
                obj.text = value
                return True
            except Exception:
                continue
        return False

    def _se16h_tem_resultados(session, max_nodes=7000):
        root = _safe_find(session, "wnd[0]")
        if not root:
            return False

        stack = [root]
        seen = 0
        while stack and seen < max_nodes:
            obj = stack.pop()
            seen += 1

            try:
                row_count = int(getattr(obj, "RowCount", 0))
                if row_count > 0:
                    return True
            except Exception:
                pass

            try:
                child_count = int(obj.Children.Count)
            except Exception:
                child_count = 0

            for idx in range(child_count):
                try:
                    stack.append(obj.Children(idx))
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

    def _find_first_grid_with_rows(session):
        root = _safe_find(session, "wnd[0]")
        if not root:
            return None

        stack = [root]
        seen = 0
        while stack and seen < 9000:
            obj = stack.pop()
            seen += 1
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
                child_count = int(obj.Children.Count)
            except Exception:
                child_count = 0
            for idx in range(child_count):
                try:
                    stack.append(obj.Children(idx))
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

    def _t028p_fetch_keys_by_panam(session, panam):
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16H"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)

            if not _set_text_any(session, [
                "wnd[0]/usr/ctxtGD-TAB",
                "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
                "wnd[0]/usr/ctxtTABNAME",
            ], "T028P"):
                return None, "Nao consegui definir T028P na SE16H."

            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.2)

            base = "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC"
            panam = safe_value(panam)
            if not panam:
                return None, "PANAM vazio para validacao da T028P.", None, set()

            ctrl_id = f"{base}/ctxtGS_SELFIELDS-LOW[2,5]"
            obj = _safe_find(session, ctrl_id)
            if not obj:
                return None, "Campo PANAM nao encontrado na selecao da SE16H.", None, set()
            obj.text = panam

            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            not_found_tokens = ("NO VALUES", "NOT FOUND", "NENHUM", "NAO ENCONTR", "SEM REGIST")
            found_tokens = ("SELECION", "SELECTED", "REGIST", "ENTRAD", "VALUES")
            reported_count = None
            status_msg = ""

            for _ in range(20):
                sbar = _safe_find(session, "wnd[0]/sbar")
                sbar_txt = str(getattr(sbar, "Text", "") or "").strip()
                status_msg = sbar_txt or status_msg
                sbar_norm = norm(sbar_txt)
                if sbar_norm and any(token in sbar_norm for token in not_found_tokens):
                    return False, sbar_txt, 0, set()
                if sbar_norm and any(token in sbar_norm for token in found_tokens):
                    reported_count = _extract_first_int(sbar_txt)
                    break

                wnd0 = _safe_find(session, "wnd[0]")
                wnd_txt = norm(str(getattr(wnd0, "Text", "") or ""))
                if "ENTRADAS ENCONTRADAS" in wnd_txt or "ENTRIES FOUND" in wnd_txt:
                    reported_count = _extract_first_int(sbar_txt or wnd_txt)
                    break

                if _se16h_tem_resultados(session):
                    reported_count = _extract_first_int(sbar_txt or wnd_txt)
                    break

                time.sleep(0.2)

            # Leitura unica do conteudo apresentado na SE16H para comparar com o esperado.
            extracted_keys = set()
            grid = _find_first_grid_with_rows(session)
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

            if extracted_keys:
                return True, status_msg or f"Registos lidos na T028P para PANAM={panam}.", reported_count, extracted_keys

            if reported_count and int(reported_count) > 0:
                return True, status_msg or f"{reported_count} registo(s) encontrado(s) na T028P.", reported_count, set()

            return False, status_msg or "Sem entradas encontradas na T028P.", reported_count or 0, set()
        except Exception as e:
            return None, str(e), None, set()
        finally:
            if not keep_validation_screen:
                try:
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                    session.findById("wnd[0]").sendVKey(0)
                except Exception:
                    pass

    ###################################################################################
    # BLOCO 2.1: REQUEST RECEBIDA DO PROCESSO PADRÃƒO
    ###################################################################################
    request_number, request_desc = resolver_request_recebida(request_transporte, request_ctx)

    if request_desc:
        print(f"INFO: Request recebida do cockpit: {request_number} | {request_desc}")
    elif request_number:
        print(f"INFO: Request recebida do cockpit: {request_number}")
    else:
        print("INFO: Request nao recebida no contexto. Sera criada automaticamente apenas se necessario.")

    def _extrair_request_do_output(output_texto: str) -> str:
        marker_req = re.search(r"REQUEST_NUMBER=([A-Z0-9]{3,4}K\d{6,})", output_texto or "")
        if marker_req:
            return validar_request(marker_req.group(1))
        fallback = re.search(r"\b([A-Z0-9]{3,4}K\d{6,})\b", output_texto or "")
        if fallback:
            return validar_request(fallback.group(1))
        return ""

    def _garantir_request(sess, motivo: str) -> str:
        nonlocal request_number
        if request_number:
            return request_number

        system_name = str(getattr(sess.Info, "SystemName", "") or SISTEMA_DESEJADO).strip().upper()
        client = CLIENTE_ESPERADO or str(getattr(sess.Info, "Client", "") or "").strip()
        if not system_name or not client:
            raise RuntimeError("Nao foi possivel resolver sistema/cliente para criar request.")

        criar_request_path = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "criar_request.py")
        )
        if not os.path.exists(criar_request_path):
            raise RuntimeError(f"Script de criacao de request nao encontrado: {criar_request_path}")

        descricao = (request_description or "").strip()
        if not descricao:
            descricao = (motivo or "").strip()
        if not descricao:
            descricao = "Atribuir Cadeia de Pesquisa"
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

    ###################################################################################
    # BLOCO 3: LEITURA / PREPARAÃ‡ÃƒO
    ###################################################################################
    if modo_nao_interativo and not caminho_ficheiro:
        print("âŒ Em modo nao-interativo, informe --xlsx.")
        return

    if not caminho_ficheiro:
        caminho_ficheiro = selecionar_ficheiro_excel()
        if not caminho_ficheiro:
            print("âŒ Nenhum ficheiro selecionado. A execuÃ§Ã£o foi cancelada.")
            return

    print(f"âœ… Ficheiro a processar: {caminho_ficheiro}")

    df = abrir_excel_para_dataframe(caminho_ficheiro)
    cols_map, faltantes = resolver_colunas(df)
    if faltantes:
        print("âŒ Colunas obrigatÃ³rias em falta apÃ³s normalizaÃ§Ã£o:")
        for c in faltantes:
            print(f"   - {c}")
        print("ðŸ”Ž CabeÃ§alhos detetados:", ", ".join(df.columns))
        return

    C_NOME = cols_map["NOME CADEIA DE PESQUISA"]
    C_EMP = cols_map["EMPRESA"]
    C_BANCO = cols_map["BANCO"]
    C_IDCT = cols_map["ID CONTA"]
    C_OPEXT = cols_map["OPERAÃ‡ÃƒO EXTERNA"]
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
            return False, "Nao foi possivel resolver PANAM para validacao na T028P."

        missing = []
        details = []
        total_expected = 0
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
            exists, msg, qtd, extracted = _t028p_fetch_keys_by_panam(session, panam)
            if exists is None:
                return None, msg or f"Validacao T028P indisponivel para PANAM={panam}."
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
                # fallback quando a leitura de colunas nao for suportada pelo controlo SAP.
                if qtd is not None and int(qtd) < int(expected_count):
                    missing.append(f"PANAM={panam} (esperado>={expected_count}, encontrado={qtd})")
                details.append(f"{panam}: esperado>={expected_count} encontrado={qtd if qtd is not None else '?'}")

        if missing:
            preview = "; ".join(missing[:3])
            return False, f"Entradas nao encontradas na T028P ({len(missing)}): {preview}"

        return True, f"Validacao T028P por PANAM concluida ({total_expected} entrada(s) esperadas). " + "; ".join(details[:3])

    def _confirm_subset_in_t028p(session, subset, tentativas=4, pausa=1.0):
        last_msg = ""
        for _ in range(max(1, int(tentativas))):
            exists, msg = _subset_exists_in_t028p(session, subset)
            last_msg = msg or last_msg
            if exists is True:
                return True, msg or "Subconjunto confirmado na T028P."
            time.sleep(max(0.0, float(pausa)))

        return False, last_msg or "Nao foi possivel confirmar o subconjunto na T028P."

    if "STATUS" not in df.columns:
        df["STATUS"] = ""
    if "MSG" not in df.columns:
        df["MSG"] = ""

    if called_by_main:
        # No fluxo main/workflow validamos sempre no SAP, independentemente do STATUS atual no ficheiro.
        df_filtrado = df.copy()
    else:
        df_filtrado = df[df["STATUS"].astype(str).map(norm) != "CONCLUIDO"].copy()
    if df_filtrado.empty:
        if called_by_main:
            print("INFO: Nenhuma linha com dados para validar/processar.")
        else:
            print("INFO: Nenhuma linha nova para processar. Tudo concluido.")
        return

    # normalizaÃ§Ãµes mÃ­nimas
    df_filtrado = df_filtrado.apply(lambda x: x.astype(str).apply(safe_value))

    # PANAM efetivo (20 chars) para gravaÃ§Ã£o
    df_filtrado["PANAM_EFETIVO"] = df_filtrado[C_NOME].apply(lambda x: sanitize_paname(x, 20))
    truncados = df_filtrado[df_filtrado[C_NOME].str.upper() != df_filtrado["PANAM_EFETIVO"]]
    if not truncados.empty:
        print("\nâš ï¸  Nomes normalizados/truncados para 20 chars (PANAM):")
        for _, r in truncados.iterrows():
            print(f"   - '{r[C_NOME]}'  âžœ  '{r['PANAM_EFETIVO']}'")

    ###################################################################################
    # BLOCO 4: CONSTRUÃ‡ÃƒO DAS LINHAS (REGRAS)
    ###################################################################################
    col_nome_dest_norm = df_filtrado[C_NDEST].apply(norm)

    is_bp = col_nome_dest_norm == "BP"
    is_centro = col_nome_dest_norm == "CENTRO"
    is_centro_lucro = col_nome_dest_norm == "CENTRO DE LUCRO"
    is_regra_cont = col_nome_dest_norm.str.contains("REGRA DE CONTABILIZACAO")
    is_xref1 = col_nome_dest_norm == "CHAVE REF 1"

    df_bp = df_filtrado[is_bp].copy()
    df_centro = df_filtrado[is_centro].copy()
    df_centro_lucro = df_filtrado[is_centro_lucro].copy()
    df_regra_cont = df_filtrado[is_regra_cont].copy()
    df_xref1 = df_filtrado[is_xref1].copy()
    df_outros = df_filtrado[~(is_bp | is_centro | is_centro_lucro | is_regra_cont | is_xref1)].copy()

    blocos = []

    # 4.1 OUTROS (standard)
    if not df_outros.empty:
        blocos.append(df_outros)

    # 4.2 BP â†’ pares +/- (NOVA REGRA)
    if not df_bp.empty:
        dd_keys_with_D = set()
        if not df_regra_cont.empty:
            df_regra_cont_chk = df_regra_cont.copy()

            if C_BASE not in df_regra_cont_chk.columns:
                raise KeyError(
                    f"Coluna '{C_BASE}' nÃ£o encontrada em df_regra_cont (Regra de ContabilizaÃ§Ã£o). "
                    f"CabeÃ§alhos: {list(df_regra_cont_chk.columns)}"
                )

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
            grp_key = (
                safe_value(t[C_NOME]),
                safe_value(t[C_EMP]),
                safe_value(t[C_BANCO]),
                safe_value(t[C_IDCT]),
                safe_value(t[C_OPEXT]),
            )
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

    # 4.3 CENTRO â†’ pares +/-
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

    # 4.3b CHAVE REF 1 -> EBFNAM2 / EBFVAL2
    if not df_xref1.empty:
        print(f"ðŸ”Ž Processando 'Chave Ref 1' para EBFNAM2/EBFVAL2 ({len(df_xref1)} registos de origem)...")
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

    # 4.4 CENTRO DE LUCRO â†’ pares +/-
    if not df_centro_lucro.empty:
        for _, grupo in df_centro_lucro.groupby([C_NOME, C_EMP, C_BANCO, C_IDCT, C_OPEXT], dropna=False):
            t = grupo.iloc[0].copy()
            t[C_DEST] = "EBPRCTR"
            blocos.append(pd.DataFrame(
                [dict(t, **{C_SINAL: "+"}), dict(t, **{C_SINAL: "-"})],
                index=[t.name] * 2
            ))

    # 4.5 EBVGINT sem duplicar; sinal decidido pela BASE
    if not df_regra_cont.empty:
        print("\nðŸ”Ž Processando 'Regra de ContabilizaÃ§Ã£o' (EBVGINT) com sinal pela BASE...")
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
        print(f"âœ… EBVGINT: {len(tmp)} linha(s) preparadas (sem duplicar).")

    if not blocos:
        print("âŒ Nenhum grupo vÃ¡lido encontrado para processar.")
        return

    df_final = pd.concat(blocos, ignore_index=False)

    ###################################################################################
    # BLOCO 5: DEDUP + PRÃ‰-VISUALIZAÃ‡ÃƒO
    ###################################################################################
    CHAVE = [C_EMP, C_BANCO, C_IDCT, C_OPEXT, C_SINAL, "PANAM_EFETIVO", C_DEST, C_BASE]
    df_final = df_final[(df_final[C_DEST].astype(str) != "") & (df_final[C_BASE].astype(str) != "")]
    df_final = df_final.drop_duplicates(subset=CHAVE, keep="first").reset_index()

    print("-" * 70)
    print(f"\nðŸ“Š Total de linhas no ficheiro original: {len(df)}")
    print(f"ðŸ“‹ Linhas a processar apÃ³s regras (antes da gravaÃ§Ã£o): {len(df_final)}")

    cols_out = [
        (C_NOME, 26),
        ("PANAM_EFETIVO", 22),
        (C_EMP, 5),
        (C_BANCO, 8),
        (C_IDCT, 12),
        (C_OPEXT, 6),
        (C_SINAL, 2),
        (C_NDEST, 25),
        (C_DEST, 9),
        (C_BASE, 15),
    ]

    def _fmt(val, width):
        v = safe_value(val)
        if v == "":
            v = "-"
        v = str(v)
        if len(v) > width:
            return v[:width]
        return v.ljust(width)

    for funcao in df_final[C_NOME].unique():
        subset = df_final[df_final[C_NOME] == funcao].copy()

        print(f"\n- {funcao}")
        for _, r in subset.iterrows():
            line = " ".join(_fmt(r[c], w) for c, w in cols_out)
            print(line)

        marcados = subset[
            (subset[C_DEST].astype(str).str.upper() == "EBAVKOA") &
            (subset[C_BASE].astype(str).str.upper() == "D")
        ]
        if not marcados.empty:
            uniq = set()
            for _, r in marcados.iterrows():
                k = (
                    safe_value(r[C_EMP]),
                    safe_value(r[C_BANCO]),
                    safe_value(r[C_IDCT]),
                    safe_value(r[C_OPEXT]),
                )
                uniq.add(k)

            print("")
            print("   âš™ï¸  Regra aplicada (EBAVKOA com PREFIX = 'D') para as chaves:")
            for emp, bco, idc, opx in sorted(uniq):
                emp = emp if emp else "-"
                bco = bco if bco else "-"
                idc = idc if idc else "-"
                opx = opx if opx else "-"
                print(f"      â€¢ EMP={emp} | BANCO={bco} | IDCONTA={idc} | OPEXT={opx}")

    if pedir_confirmacao and not modo_nao_interativo:
        resp = input("\nâž¡ï¸  Confirmar lanÃ§amento no SAP com a TABELA FINAL acima? [S/N]: ").strip().upper()
        if resp != "S":
            print("âŒ LanÃ§amento cancelado pelo utilizador.")
            return

    ###################################################################################
    # BLOCO 6: CONEXÃƒO SAP
    ###################################################################################
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = None
        for conn in application.Children:
            for sess in conn.Children:
                if sess.Info.SystemName.upper() != SISTEMA_DESEJADO:
                    continue
                if CLIENTE_ESPERADO and str(sess.Info.Client).strip() != CLIENTE_ESPERADO:
                    continue
                session = sess
                break
            if session:
                break

        if not session:
            print(f"âŒ Nenhuma sessÃ£o encontrada para o ambiente '{ambiente_cockpit}'.")
            return

        print(f"\nâœ… Conectado ao SAP: {session.Info.SystemName} (ambiente {ambiente_cockpit})")
        print(f"ðŸ‘¤ Utilizador SAP: {session.Info.User} | Cliente: {session.Info.Client}")
        aplicar_modo_janela_sap(session)
    except Exception as e:
        print(f"âŒ Erro ao conectar SAP: {e}")
        return

    ###################################################################################
    # BLOCO 7: LANÃ‡AMENTO NO SAP (TARGFI ANTES DE PANAM + VALIDAÃ‡Ã•ES)
    ###################################################################################
    TBL = "wnd[0]/usr/tblSAPLPAMVTCTRL_V_T028P"

    for funcao in df_final[C_NOME].unique():
        subset = df_final[df_final[C_NOME] == funcao]
        print(f"\nAtribuindo funcao '{funcao}' ({len(subset)} linha(s))")

        print("  |- A validar existencia na T028P (SE16H)...")
        pre_exists, pre_msg = _subset_exists_in_t028p(session, subset)
        if pre_exists is True:
            msg_done = pre_msg or "Entradas ja existem na T028P."
            df.loc[subset["index"], "STATUS"] = "CONCLUIDO"
            df.loc[subset["index"], "MSG"] = msg_done
            print(f"  |- CONCLUIDO sem lancamento: {msg_done}")
            continue
        if pre_exists is None:
            print(f"  |- Aviso: validacao inicial T028P indisponivel ({pre_msg}). Vou seguir com o lancamento.")
        if not request_number:
            print("  |- Sem request no contexto. A criar request automaticamente...")
            _garantir_request(session, f"Atribuir Cadeia Pesquisa | {funcao}")

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

            # Validacao unica do lote para reduzir o tempo do preenchimento linha-a-linha.
            print(f"  |- Lote preenchido ({total_linhas_lote} linha(s)). A validar consistencia no SAP...")
            session.findById("wnd[0]").sendVKey(0)
            err = sbar_error(session)
            if err:
                raise RuntimeError(f"Erro de validacao do lote '{funcao}': {err}")

            # Grava + TRKORR vindo do cockpit
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = request_number
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            msg_gravacao = session.findById("wnd[0]/sbar").Text.strip()
            print("  |- A validar gravacao na T028P (SE16H)...")
            validado, msg_validacao = _confirm_subset_in_t028p(session, subset, tentativas=4, pausa=1.0)
            if not validado:
                raise RuntimeError(
                    f"Lancamento executado sem confirmacao na T028P. "
                    f"Gravacao SAP='{msg_gravacao}' | Validacao='{msg_validacao}'"
                )

            msg_final = msg_validacao or msg_gravacao
            df.loc[subset["index"], "STATUS"] = "CONCLUIDO"
            df.loc[subset["index"], "MSG"] = msg_final
            print(f"  |- CONCLUIDO: {msg_final}")

        except Exception as e:
            msg_erro = f"Erro ao lancar '{funcao}': {e}"
            df.loc[subset["index"], "STATUS"] = "Erro no processamento"
            df.loc[subset["index"], "MSG"] = msg_erro
            print(f"ERRO: {msg_erro}")

    ###################################################################################
    # BLOCO 8: GUARDAR CONTROLO
    ###################################################################################
    try:
        df.to_excel(caminho_ficheiro, index=False)
        print(f"\nðŸ’¾ Ficheiro de controlo atualizado com sucesso: {caminho_ficheiro}")
    except Exception as e:
        print(f"âŒ Erro ao guardar o ficheiro de controlo: {e}")
    if request_number:
        print(f"REQUEST_NUMBER={request_number}")


if __name__ == "__main__":
    import argparse
    import os

    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"], required=True)
    parser.add_argument("--request", help="Request opcional para execuÃ§Ã£o direta fora do cockpit.")
    parser.add_argument("--request-description", default="", help="Descricao da request. Ex: Chave | Resumo")
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



