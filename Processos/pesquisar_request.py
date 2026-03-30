# -*- coding: utf-8 -*-
"""
pesquisar_request.py

Objetivo:
- Abrir SE16H em NOVO modo (/ose16h)
- Minimizar a janela desse novo modo enquanto executa
- Ler resultados da E070 e imprimir lista NUMERADA:
    N | TRKORR | AS4TEXT
- Regra atualizada: APENAS listar as linhas cujo valor da coluna STRKORR for diferente de vazio.
- Guardar automaticamente a lista num ficheiro JSON para uso posterior (seleção por número da linha)

Execução:
& "C:/SAP Script/.venv/Scripts/python.exe" "C:/SAP Script/Processos/pesquisar_request.py" --system "S4Q" --max "5000"

Opcional:
--no-new-mode        (não usa /o; usa a sessão atual)
--no-minimize        (não minimiza)
--no-close           (não fecha a janela ao terminar)
"""

import sys
import time
import json
import os
import win32com.client

MSG_RZ11_SCRIPTING = 'Ativar na transação RZ11 o nome do parametro "sapgui/user_scripting" alterar para "TRUE"'


def _base_dir():
    # este ficheiro está em ...\SAP Script\Processos\pesquisar_request.py
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _cache_dir():
    d = os.path.join(_base_dir(), "cache")
    os.makedirs(d, exist_ok=True)
    return d


def _cache_file_path():
    return os.path.join(_cache_dir(), "last_e070_list.json")


def _save_results(results, system_name, user):
    """
    results: list[tuple(trkorr, as4text)]
    Guarda:
      - índice (1..N)
      - trkorr
      - as4text
      - metadata
    """
    payload = {
        "meta": {
            "system": system_name,
            "user": user,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        },
        "items": [
            {"idx": i + 1, "TRKORR": trkorr, "AS4TEXT": as4text}
            for i, (trkorr, as4text) in enumerate(results)
        ],
    }

    path = _cache_file_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    os.environ["SAP_LAST_E070_LIST_FILE"] = path
    os.environ["SAP_LAST_E070_LIST_COUNT"] = str(len(results))

    print(f"\n💾 Lista guardada em: {path}")
    return path


def _log_alerta_rz11():
    print(f"⚠️  {MSG_RZ11_SCRIPTING}")


def _erro_scripting_inativo(e=None):
    print("❌ O scripting do SAP GUI não está ativo ou não foi possível inicializar o objeto SAPGUI.")
    _log_alerta_rz11()
    if e:
        print(f"🔧 Detalhes técnicos: {e}")
    sys.exit(1)


def _get_application():
    try:
        sap = win32com.client.GetObject("SAPGUI")
        app = sap.GetScriptingEngine
        if not app:
            raise RuntimeError("GetScriptingEngine retornou vazio/None.")
        return app
    except Exception as e:
        _erro_scripting_inativo(e)


def _iter_sessions(application):
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            try:
                for j in range(conn.Children.Count):
                    yield conn.Children(j)
            except Exception:
                continue
    except Exception:
        return


def _pick_session(application, system_name=None):
    candidates = []
    for sess in _iter_sessions(application):
        try:
            sysname = (sess.Info.SystemName or "").upper()
        except Exception:
            sysname = ""
        try:
            user = (sess.Info.User or "").strip()
        except Exception:
            user = ""
        candidates.append((sysname, bool(user), sess))

    if not candidates:
        print("❌ Nenhuma sessão SAP encontrada. Abra o SAP Logon e faça login.")
        sys.exit(1)

    if system_name:
        target = system_name.upper()
        in_sys = [c for c in candidates if c[0] == target]
        if in_sys:
            logged = [c for c in in_sys if c[1]]
            return logged[0][2] if logged else in_sys[0][2]

    logged_any = [c for c in candidates if c[1]]
    return logged_any[0][2] if logged_any else candidates[0][2]


def _wait_not_busy(session, timeout_s=12):
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        try:
            if not session.Busy:
                return True
        except Exception:
            return True
        time.sleep(0.1)
    return False


def _try_set_text(session, id_path, value):
    try:
        session.findById(id_path).text = value
        return True
    except Exception:
        return False


def _try_press(session, id_path):
    try:
        session.findById(id_path).press()
        return True
    except Exception:
        return False


def _iconify(session):
    try:
        session.findById("wnd[0]").iconify()
        return True
    except Exception:
        return False


def _close_window(session):
    try:
        session.findById("wnd[0]").close()
    except Exception:
        return

    time.sleep(0.3)

    # confirmação padrão de popup (se existir)
    try:
        if _try_press(session, "wnd[1]/usr/btnSPOP-OPTION1"):
            return
        if _try_press(session, "wnd[1]/tbar[0]/btn[0]"):
            return
    except Exception:
        pass


def _set_table_e070(session):
    candidates = [
        "wnd[0]/usr/ctxtGD-TAB",
        "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
        "wnd[0]/usr/ctxtTABNAME",
    ]
    for cid in candidates:
        if _try_set_text(session, cid, "E070"):
            try:
                session.findById("wnd[0]").sendVKey(0)
            except Exception:
                pass
            _wait_not_busy(session, 10)
            time.sleep(0.2)
            return True
    return False


def _set_max_ocorrencias(session, max_rows="5000"):
    candidates = [
        "wnd[0]/usr/txtMAX_SEL",
        "wnd[0]/usr/txtGD-MAXROWS",
        "wnd[0]/usr/txtMAX_HITS",
    ]
    for cid in candidates:
        if _try_set_text(session, cid, str(max_rows)):
            return True
    return False


def _aplicar_filtros_base(session, user):
    ok1 = _try_set_text(
        session,
        "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]",
        "D",
    )
    ok2 = _try_set_text(
        session,
        "wnd[0]/usr/subTAB_SUB:SAPLSE16N:0121/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,5]",
        user,
    )
    return ok1 and ok2


def _press_execute(session):
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_not_busy(session, 12)
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"❌ Falha ao executar (F8) no SE16H: {e}")
        return False


def _walk_children(root, max_nodes=8000):
    stack = [root]
    seen = 0
    while stack and seen < max_nodes:
        obj = stack.pop()
        seen += 1
        yield obj
        try:
            cnt = obj.Children.Count
        except Exception:
            continue
        for i in range(cnt - 1, -1, -1):
            try:
                stack.append(obj.Children(i))
            except Exception:
                continue


def _score_grid_candidate(obj):
    try:
        rc = int(obj.RowCount)
        if rc < 0:
            return -1
    except Exception:
        return -1

    score = 0
    if rc > 0:
        score += 5

    # assinatura de E070
    for col in ("TRKORR", "STRKORR", "AS4TEXT"):
        try:
            _ = obj.GetCellValue(0, col)
            score += 10
        except Exception:
            pass

    return score


def _find_best_grid(session):
    roots = []
    try:
        roots.append(session.findById("wnd[0]/usr"))
    except Exception:
        pass
    try:
        roots.append(session.findById("wnd[0]"))
    except Exception:
        pass

    candidates = []
    for root in roots:
        for obj in _walk_children(root):
            s = _score_grid_candidate(obj)
            if s >= 0:
                candidates.append((s, obj))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _get_cell(grid, row, col):
    try:
        return str(grid.GetCellValue(row, col)).strip()
    except Exception:
        return ""


def _open_se16h_new_mode(session):
    """
    Abre /ose16h em novo modo e devolve (new_session, created_flag).
    """
    try:
        connection = session.Parent
        before = int(connection.Children.Count)
    except Exception:
        connection = None
        before = None

    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ose16h"
        session.findById("wnd[0]").sendVKey(0)
    except Exception:
        return session, False

    if connection is not None and before is not None:
        t0 = time.time()
        while time.time() - t0 <= 6:
            try:
                now = int(connection.Children.Count)
                if now > before:
                    new_sess = connection.Children(now - 1)
                    _wait_not_busy(new_sess, 12)
                    time.sleep(0.2)
                    return new_sess, True
            except Exception:
                pass
            time.sleep(0.2)

    _wait_not_busy(session, 12)
    time.sleep(0.2)
    return session, False


def listar_requests(
    system_name=None,
    max_rows="5000",
    include_requests=False,
    use_new_mode=True,
    minimize=True,
    close_after=True,
):
    app = _get_application()
    base_session = _pick_session(app, system_name=system_name)

    try:
        user = (base_session.Info.User or "").strip()
    except Exception:
        user = ""

    if not user:
        print("❌ Sessão encontrada, mas não está logada (Info.User vazio). Faça login no SAP e execute novamente.")
        return []

    work_session = base_session
    created_new = False

    if use_new_mode:
        work_session, created_new = _open_se16h_new_mode(base_session)
    else:
        try:
            base_session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16h"
            base_session.findById("wnd[0]").sendVKey(0)
            _wait_not_busy(base_session, 12)
            time.sleep(0.2)
        except Exception:
            pass

    if minimize:
        _iconify(work_session)

    _set_table_e070(work_session)
    _set_max_ocorrencias(work_session, max_rows=max_rows)

    if not _aplicar_filtros_base(work_session, user):
        print("⚠️ Não consegui aplicar TRSTATUS/AS4USER pelos IDs esperados (layout SE16H diferente). Vou executar mesmo assim.")

    if not _press_execute(work_session):
        if created_new and close_after:
            _close_window(work_session)
        return []

    grid = _find_best_grid(work_session)
    if not grid:
        print("❌ Não encontrei a grelha de resultados para ler TRKORR/STRKORR/AS4TEXT.")
        if created_new and close_after:
            _close_window(work_session)
        return []

    try:
        row_count = int(grid.RowCount)
    except Exception:
        print("❌ Não consegui obter RowCount da grelha.")
        if created_new and close_after:
            _close_window(work_session)
        return []

    results = []
    for r in range(row_count):
        trkorr = _get_cell(grid, r, "TRKORR")
        strkorr = _get_cell(grid, r, "STRKORR")
        as4text = _get_cell(grid, r, "AS4TEXT") or _get_cell(grid, r, "TXT_BREVE") or _get_cell(grid, r, "TEXT")

        # Ajuste exigido: Ignorar todas as linhas onde a coluna STRKORR for vazia
        if not strkorr:
            continue

        if trkorr or as4text:
            results.append((trkorr, as4text))

    # impressão NUMERADA
    try:
        sysname = work_session.Info.SystemName
    except Exception:
        sysname = ""

    print(f"\n✅ Resultados: {len(results)} | Sistema={sysname} | User={user}")
    print("N | TRKORR | AS4TEXT")
    print("-" * 90)
    for i, (trkorr, as4text) in enumerate(results, start=1):
        print(f"{i} | {trkorr} | {as4text}")

    # guarda para seleção futura por número
    _save_results(results, system_name=sysname, user=user)

    if created_new and close_after:
        _close_window(work_session)

    return results


def _parse_args(argv):
    system = None
    max_rows = "5000"
    include_requests = False
    use_new_mode = True
    minimize = True
    close_after = True

    i = 0
    while i < len(argv):
        a = argv[i].strip()
        if a == "--system" and i + 1 < len(argv):
            system = argv[i + 1]
            i += 2
            continue
        if a == "--max" and i + 1 < len(argv):
            max_rows = argv[i + 1]
            i += 2
            continue
        if a == "--include-requests":
            include_requests = True
            i += 1
            continue
        if a == "--no-new-mode":
            use_new_mode = False
            i += 1
            continue
        if a == "--no-minimize":
            minimize = False
            i += 1
            continue
        if a == "--no-close":
            close_after = False
            i += 1
            continue
        i += 1

    return system, max_rows, include_requests, use_new_mode, minimize, close_after


if __name__ == "__main__":
    system, max_rows, include_requests, use_new_mode, minimize, close_after = _parse_args(sys.argv[1:])
    listar_requests(
        system_name=system,
        max_rows=max_rows,
        include_requests=include_requests,
        use_new_mode=use_new_mode,
        minimize=minimize,
        close_after=close_after,
    )