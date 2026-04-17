###################################################################################
# SCRIPT: Lançar Cadeias de Pesquisa (OTPM) — EBVGINT, PANAM 20 chars
# VERSÃO ATUALIZADA:
# 1. Z062 é estritamente POSITIVO (+).
# 2. Z063 é estritamente NEGATIVO (-).
# 3. Inclui lógica "CHAVE REF 1" -> EBFNAM2/EBFVAL2.
# 4. Não pede mais a request por input(); usa o processo padrão de request do cockpit.
###################################################################################

def executar(
    ambiente_cockpit,
    request_ctx,                 # OBRIGATÓRIO: força o cockpit a chamar o processo da request
    request_transporte=None
):
    ###################################################################################
    # BLOCO 1: IMPORTAÇÕES / CONFIG GERAL
    ###################################################################################
    import re
    import time
    import warnings
    import unicodedata
    import pandas as pd
    import win32com.client
    import tkinter as tk
    from tkinter import filedialog
    import sys

    warnings.simplefilter("ignore", UserWarning)
    warnings.simplefilter("ignore", FutureWarning)

    MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    SISTEMA_DESEJADO = MAPA_SISTEMA.get(ambiente_cockpit)

    if not SISTEMA_DESEJADO:
        print(f"❌ Ambiente inválido: {ambiente_cockpit}")
        return

    ###################################################################################
    # BLOCO 2: HELPERS / UTILITÁRIOS
    ###################################################################################
    def norm(s: str) -> str:
        """Remove acentos, normaliza espaços, trim e upper (para strings simples)."""
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
        """Normalização robusta para Series: NaN-safe, strip e upper."""
        return s.fillna("").astype(str).str.strip().str.upper()

    def map_sign_free(val: str) -> str:
        """Normaliza sinais de livre-tecla para '+' ou '-'."""
        v = (val or "").strip().upper()
        if v in {"+", "PLUS", "POS", "P"}:
            return "+"
        if v in {"-", "MINUS", "NEG", "M"}:
            return "-"
        return ""

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
    # ATENÇÃO: Z062 fixo em + e Z063 fixo em -
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
        print("📁 Selecione o ficheiro Excel (janela foi colocada em primeiro plano)...")
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
        # Versão segura para avisar se o ficheiro estiver aberto
        try:
            return pd.read_excel(caminho, sheet_name="Folha2", dtype=str).fillna("")
        except PermissionError:
            print("\n❌ ERRO: O ficheiro Excel está ABERTO.")
            print("👉 Por favor, feche o ficheiro Excel e execute o script novamente.")
            sys.exit()
        except Exception:
            try:
                return pd.read_excel(caminho, dtype=str).fillna("")
            except PermissionError:
                print("\n❌ ERRO: O ficheiro Excel está ABERTO.")
                print("👉 Por favor, feche o ficheiro Excel e execute o script novamente.")
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

    ###################################################################################
    # BLOCO 2.1: REQUEST RECEBIDA DO PROCESSO PADRÃO
    ###################################################################################
    request_number, request_desc = resolver_request_recebida(request_transporte, request_ctx)

    if not request_number:
        print("❌ Nenhuma request foi recebida do processo de transporte.")
        print("👉 Execute este script pelo cockpit ou informe uma request válida.")
        return

    if request_desc:
        print(f"✅ Request recebida do cockpit: {request_number} | {request_desc}")
    else:
        print(f"✅ Request recebida do cockpit: {request_number}")

    ###################################################################################
    # BLOCO 3: LEITURA / PREPARAÇÃO
    ###################################################################################
    caminho_ficheiro = selecionar_ficheiro_excel()
    if not caminho_ficheiro:
        print("❌ Nenhum ficheiro selecionado. A execução foi cancelada.")
        return
    print(f"✅ Ficheiro a processar: {caminho_ficheiro}")

    df = abrir_excel_para_dataframe(caminho_ficheiro)
    cols_map, faltantes = resolver_colunas(df)
    if faltantes:
        print("❌ Colunas obrigatórias em falta após normalização:")
        for c in faltantes:
            print(f"   - {c}")
        print("🔎 Cabeçalhos detetados:", ", ".join(df.columns))
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

    if "STATUS" not in df.columns:
        df["STATUS"] = ""
    if "MSG" not in df.columns:
        df["MSG"] = ""

    df_filtrado = df[df["STATUS"].astype(str).map(norm) != "CONCLUIDO"].copy()
    if df_filtrado.empty:
        print("✅ Nenhuma linha nova para processar. Tudo concluído.")
        return

    # normalizações mínimas
    df_filtrado = df_filtrado.apply(lambda x: x.astype(str).apply(safe_value))

    # PANAM efetivo (20 chars) para gravação
    df_filtrado["PANAM_EFETIVO"] = df_filtrado[C_NOME].apply(lambda x: sanitize_paname(x, 20))
    truncados = df_filtrado[df_filtrado[C_NOME].str.upper() != df_filtrado["PANAM_EFETIVO"]]
    if not truncados.empty:
        print("\n⚠️  Nomes normalizados/truncados para 20 chars (PANAM):")
        for _, r in truncados.iterrows():
            print(f"   - '{r[C_NOME]}'  ➜  '{r['PANAM_EFETIVO']}'")

    ###################################################################################
    # BLOCO 4: CONSTRUÇÃO DAS LINHAS (REGRAS)
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

    # 4.2 BP → pares +/- (NOVA REGRA)
    if not df_bp.empty:
        dd_keys_with_D = set()
        if not df_regra_cont.empty:
            df_regra_cont_chk = df_regra_cont.copy()

            if C_BASE not in df_regra_cont_chk.columns:
                raise KeyError(
                    f"Coluna '{C_BASE}' não encontrada em df_regra_cont (Regra de Contabilização). "
                    f"Cabeçalhos: {list(df_regra_cont_chk.columns)}"
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

    # 4.3 CENTRO → pares +/-
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
        print(f"🔎 Processando 'Chave Ref 1' para EBFNAM2/EBFVAL2 ({len(df_xref1)} registos de origem)...")
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

    # 4.4 CENTRO DE LUCRO → pares +/-
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
        print("\n🔎 Processando 'Regra de Contabilização' (EBVGINT) com sinal pela BASE...")
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
        print(f"✅ EBVGINT: {len(tmp)} linha(s) preparadas (sem duplicar).")

    if not blocos:
        print("❌ Nenhum grupo válido encontrado para processar.")
        return

    df_final = pd.concat(blocos, ignore_index=False)

    ###################################################################################
    # BLOCO 5: DEDUP + PRÉ-VISUALIZAÇÃO
    ###################################################################################
    CHAVE = [C_EMP, C_BANCO, C_IDCT, C_OPEXT, C_SINAL, "PANAM_EFETIVO", C_DEST, C_BASE]
    df_final = df_final[(df_final[C_DEST].astype(str) != "") & (df_final[C_BASE].astype(str) != "")]
    df_final = df_final.drop_duplicates(subset=CHAVE, keep="first").reset_index()

    print("-" * 70)
    print(f"\n📊 Total de linhas no ficheiro original: {len(df)}")
    print(f"📋 Linhas a processar após regras (antes da gravação): {len(df_final)}")

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
            print("   ⚙️  Regra aplicada (EBAVKOA com PREFIX = 'D') para as chaves:")
            for emp, bco, idc, opx in sorted(uniq):
                emp = emp if emp else "-"
                bco = bco if bco else "-"
                idc = idc if idc else "-"
                opx = opx if opx else "-"
                print(f"      • EMP={emp} | BANCO={bco} | IDCONTA={idc} | OPEXT={opx}")

    resp = input("\n➡️  Confirmar lançamento no SAP com a TABELA FINAL acima? [S/N]: ").strip().upper()
    if resp != "S":
        print("❌ Lançamento cancelado pelo utilizador.")
        return

    ###################################################################################
    # BLOCO 6: CONEXÃO SAP
    ###################################################################################
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = None
        for conn in application.Children:
            for sess in conn.Children:
                if sess.Info.SystemName.upper() == SISTEMA_DESEJADO:
                    session = sess
                    break
            if session:
                break

        if not session:
            print(f"❌ Nenhuma sessão encontrada para o ambiente '{ambiente_cockpit}'.")
            return

        print(f"\n✅ Conectado ao SAP: {session.Info.SystemName} (ambiente {ambiente_cockpit})")
        print(f"👤 Utilizador SAP: {session.Info.User} | Cliente: {session.Info.Client}")
    except Exception as e:
        print(f"❌ Erro ao conectar SAP: {e}")
        return

    ###################################################################################
    # BLOCO 7: LANÇAMENTO NO SAP (TARGFI ANTES DE PANAM + VALIDAÇÕES)
    ###################################################################################
    TBL = "wnd[0]/usr/tblSAPLPAMVTCTRL_V_T028P"

    for funcao in df_final[C_NOME].unique():
        subset = df_final[df_final[C_NOME] == funcao]
        print(f"\n🔧 Atribuindo função '{funcao}' ({len(subset)} linha(s))")

        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nOTPM"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/shellcont/shell").selectItem("02", "Column1")
            session.findById("wnd[0]/shellcont/shell").doubleClickItem("02", "Column1")
            session.findById("wnd[0]/tbar[1]/btn[25]").press()
            session.findById("wnd[0]/tbar[1]/btn[5]").press()

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
                    raise RuntimeError(f"Não foi possível definir TARGFI='{targfi}' (ficou '{lido}')")

                lido = get_cell_text(session, TBL, "V_T028P-TARGFI", 7, i)
                if (lido or "").strip().upper() not in {targfi.strip().upper()}:
                    raise RuntimeError(f"TARGFI não persistiu: esperado '{targfi}', lido '{lido}'")

                set_cell_any(session, TBL, "V_T028P-PREFIX", 9, i, safe_value(row[C_BASE]))

                panam_val = safe_value(row["PANAM_EFETIVO"])
                set_cell_any(session, TBL, "V_T028P-PANAM", 6, i, panam_val)
                session.findById("wnd[0]").sendVKey(0)

                err = sbar_error(session)
                if err:
                    raise RuntimeError(f"PANAM inválido '{panam_val}' no lote '{funcao}': {err}")

                try:
                    session.findById(f"{TBL}/chkV_T028P-ENABLED[8,{i}]").selected = True
                except Exception:
                    pass

            # Grava + TRKORR vindo do cockpit
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/ctxtKO008-TRKORR").text = request_number
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            msg = session.findById("wnd[0]/sbar").Text.strip()
            df.loc[subset["index"], "STATUS"] = "PROCESSADO"
            df.loc[subset["index"], "MSG"] = msg
            print(f"✅ Função '{funcao}' atribuída: {msg}")

        except Exception as e:
            msg_erro = f"Erro ao lançar '{funcao}': {e}"
            df.loc[subset["index"], "STATUS"] = "Erro no processamento"
            df.loc[subset["index"], "MSG"] = msg_erro
            print(f"❌ {msg_erro}")

    ###################################################################################
    # BLOCO 8: GUARDAR CONTROLO
    ###################################################################################
    try:
        df.to_excel(caminho_ficheiro, index=False)
        print(f"\n💾 Ficheiro de controlo atualizado com sucesso: {caminho_ficheiro}")
    except Exception as e:
        print(f"❌ Erro ao guardar o ficheiro de controlo: {e}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"], required=True)
    parser.add_argument("--request", help="Request opcional para execução direta fora do cockpit.")
    args = parser.parse_args()

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
    )