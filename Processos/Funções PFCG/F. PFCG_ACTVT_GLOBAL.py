# -*- coding: utf-8 -*-

###################################################################################
# F. PFCG_ACTVT_GLOBAL.py
# PFCG - Aplicar Autorização Global (valor '*', Full Auth) a um campo, via
# PFCGMASSVAL — suporta execução individual (uma função) e em massa (várias
# funções, seleção múltipla, numa única execução da transação).
#
# A transação PFCGMASSVAL tem duas opções de entrada distintas consoante o
# tipo de campo, e a escolha errada não produz o efeito desejado:
#   - Campos NORMAIS (residem só em AGR_1251, ex.: ACTVT) -> opção
#     "Campo para vários objetos" (radSEL_FLD), modo 'T - Substituir tudo'.
#   - Campos de NÍVEL ORGANIZACIONAL (residem também em AGR_1252, ex.:
#     BUKRS, WERKS, EKORG, VKORG, KOART, FIKRS/FM_FIKRS...) -> opção
#     "Modificar níveis organizacionais" (radSEL_ORG), modo 'A - Inserir'.
#     Usar radSEL_FLD nestes campos não garante escrever o valor real em
#     AGR_1252 (ver memória do projeto "PFCG BUKRS org-field").
#
# A deteção do tipo de campo é automática (`campo_e_organizacional`): lê
# AGR_1252 via RFC (somente leitura) à procura de qualquer linha com
# VARBL = '$<CAMPO>' em qualquer função do sistema — se existir, o campo é
# tratado como organizacional.
#
# Em ambos os casos, depois de abrir "Vals." (btnPFLDN / btnPORGN) pode
# aparecer um de dois botões, consoante o domínio de valores do campo:
#   - btnGES1 "Autorização global para" (campos de valor aberto, ex. WERKS,
#     BUKRS, ACTVT não usa este) -> abre popup de confirmação "Inserir
#     autorização global" -> botão "Sim".
#   - btnGES2 "Autorização global" (campos de domínio fixo, ex. ACTVT,
#     KOART) -> marca diretamente as checkboxes de valor, sem confirmação
#     extra.
# Se nenhum dos dois aparecer, a função já tem autorização global atribuída
# a esse campo — é saltada (mass) ou reportada como skip (individual), sem
# ser tratada como falha.
#
# Seleção de funções:
#   - 1 função -> preenche ctxtROLE-LOW diretamente.
#   - 2+ funções -> popup de "Seleção múltipla para ROLE", populado via
#     upload do clipboard (Shift+F12), não pela scrollbar (ver memória do
#     projeto "PFCG BUKRS org-field" para o porquê). Uma única execução da
#     transação processa todas as funções selecionadas de uma vez — não é
#     um loop função a função.
#
# Verificação (RFC, somente leitura, sempre depois de cada execução):
#   - Campo organizacional: AGR_1252, LOW = '*' para o VARBL do campo.
#   - Campo normal: AGR_1251, LOW = '*' em todas as linhas não eliminadas
#     desse campo.
# Nunca confiar apenas no ecrã da transação — o resultado é sempre
# reconfirmado por leitura RFC independente.
###################################################################################

import sys
import time
from pathlib import Path
from typing import Any, Dict, List

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
CAMINHO_EXCEL_PADRAO = PROJECT_ROOT / "sap_script_uploads" / "S4H_Perfis de autorização.xlsx"
NOME_SHEET_OBJETOS_ATUALIZADOS = "OBJETOS ATUALIZADOS"


def registar_objetos_atualizados(role_name: str, objetos_valores: List[tuple], caminho_excel: str = None) -> None:
    """Regista (FUNÇÃO, OBJETO, VALOR) na sheet 'OBJETOS ATUALIZADOS' do ficheiro Excel oficial."""
    from openpyxl import load_workbook

    if not objetos_valores:
        return

    caminho = Path(caminho_excel) if caminho_excel else CAMINHO_EXCEL_PADRAO
    wb = load_workbook(caminho)
    if NOME_SHEET_OBJETOS_ATUALIZADOS not in wb.sheetnames:
        ws = wb.create_sheet(NOME_SHEET_OBJETOS_ATUALIZADOS)
        ws.append(["FUNÇÃO", "OBJETO", "VALOR"])
    else:
        ws = wb[NOME_SHEET_OBJETOS_ATUALIZADOS]

    for objeto, valor in objetos_valores:
        ws.append([role_name, objeto, valor])
    wb.save(caminho)


def _get_session(ambiente: str):
    import win32com.client

    sistema_esperado = MAPA_SISTEMA.get(str(ambiente or "").strip().upper())
    if not sistema_esperado:
        raise ValueError(f"Ambiente inválido: '{ambiente}'. Use DEV, QAD ou PRD.")

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    session = next(
        (sess for conn in application.Children for sess in conn.Children
         if sess.Info.SystemName.upper() == sistema_esperado),
        None,
    )
    if not session:
        raise Exception(f"Não encontrei sessão SAP GUI aberta para '{ambiente}' ({sistema_esperado}).")
    return session


def campo_e_organizacional(campo: str, ambiente: str = "PRD") -> bool:
    """Somente leitura (RFC): True se o campo existir em AGR_1252 (nível
    organizacional) em qualquer função do sistema, False caso contrário
    (campo normal, só em AGR_1251)."""
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, read_table,
    )
    from pyrfc import Connection

    campo = campo.strip().upper()
    load_project_env(find_project_root())
    params = build_connection_params_for(ambiente)
    guard = make_read_only_guard(("AGR_1252",))
    conn = Connection(**params)
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_1252",
            fields=["AGR_NAME", "VARBL"],
            options=[{"TEXT": f"VARBL = '${campo}'"}],
            rowcount=1,
        )
    finally:
        conn.close()
    return bool(rows)


def obter_objetos_com_campo(role_name: str, campo: str, ambiente: str = "PRD") -> List[str]:
    """Somente leitura (RFC): objetos de autorização da função que têm o campo indicado (AGR_1251)."""
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, make_option_eq, read_table,
    )
    from pyrfc import Connection

    campo = campo.strip().upper()
    load_project_env(find_project_root())
    params = build_connection_params_for(ambiente)
    guard = make_read_only_guard(("AGR_1251",))
    conn = Connection(**params)
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_1251",
            fields=["AGR_NAME", "OBJECT", "FIELD", "DELETED"],
            options=make_option_eq("AGR_NAME", role_name.strip().upper()),
            rowcount=0,
        )
    finally:
        conn.close()

    objetos = sorted({
        obj.strip()
        for _agr, obj, campo_linha, deleted in rows
        if campo_linha.strip().upper() == campo and deleted.strip().upper() != "X"
    })
    return objetos


class AutorizacaoGlobalPage:
    def __init__(self, sess):
        self.sess = sess

    def _find(self, sap_id):
        try:
            return self.sess.findById(sap_id)
        except Exception:
            return None

    def _exists(self, sap_id):
        return self._find(sap_id) is not None

    def _busy(self):
        try:
            return bool(getattr(self.sess, "Busy", False))
        except Exception:
            return False

    def _esperar_livre(self, timeout=8.0, pausa=0.05):
        limite = time.time() + timeout
        while time.time() < limite:
            if not self._busy():
                return True
            time.sleep(pausa)
        return False

    def _sbar(self):
        try:
            sbar = self.sess.findById("wnd[0]/sbar")
            return getattr(sbar, "MessageType", "").strip().upper(), (sbar.Text or "").strip()
        except Exception:
            return "", ""

    def _step(self, descricao, path, acao="press", valor=None, vkey=None):
        print(f"  🔎 {descricao}")
        try:
            elem = self.sess.findById(path) if path else None
            if acao == "text":
                elem.text = valor
            elif acao == "press":
                elem.press()
            elif acao == "select":
                if hasattr(elem, "selected"):
                    elem.selected = True
                else:
                    elem.select()
            elif acao == "unselect":
                elem.selected = False
            elif acao == "key":
                elem.key = valor
            elif acao == "sendVKey":
                (elem if path else self.sess.findById("wnd[0]")).sendVKey(vkey)
            self._esperar_livre()
        except Exception as e:
            raise Exception(f"Falha no passo '{descricao}' (ID: {path}): {e}")

    def _selecionar_roles_multipla(self, roles: List[str]) -> None:
        """Popula o popup 'Seleção múltipla para ROLE' via upload do
        clipboard — funciona para 1+ funções, sem limite de linhas visíveis."""
        import win32clipboard

        self._step("Abrir seleção múltipla de ROLE", "wnd[0]/usr/btn%_ROLE_%_APP_%-VALU_PUSH", "press")
        time.sleep(0.6)
        self._step("Eliminar completamente seleção", "wnd[1]/tbar[0]/btn[16]", "press")
        time.sleep(0.4)

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText("\r\n".join(roles), win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        time.sleep(0.3)

        self._step("Upload do clipboard", "wnd[1]/tbar[0]/btn[24]", "press")
        time.sleep(0.6)
        self._step("Transferir", "wnd[1]/tbar[0]/btn[8]", "press")
        time.sleep(0.4)

    def _aplicar_global_no_popup_vals(self, btn_vals_id: str) -> Dict[str, str]:
        """Abre 'Vals.' e clica no botão de autorização global disponível
        (btnGES1 ou btnGES2). Retorna {'status': 'aplicado'|'skip', ...}."""
        if not self._exists(f"wnd[0]/usr/{btn_vals_id}"):
            return {"status": "erro", "message": "Botão 'Vals.' não encontrado."}

        self._step("Abrir 'Vals.'", f"wnd[0]/usr/{btn_vals_id}", "press")
        time.sleep(0.4)

        tem_ges2 = self._exists("wnd[1]/usr/btnGES2")
        tem_ges1 = self._exists("wnd[1]/usr/btnGES1")

        if not tem_ges2 and not tem_ges1:
            if self.sess.Children.Count > 1:
                self._step("Cancelar popup de valores", "wnd[1]/tbar[0]/btn[12]", "press")
            return {"status": "skip", "message": "Botão de autorização global não apareceu (pode já estar atribuída)."}

        if tem_ges2:
            # Campo de domínio fixo (ex.: ACTVT, KOART) — marca as checkboxes diretamente.
            self._step("Autorização global (domínio fixo)", "wnd[1]/usr/btnGES2", "press")
        else:
            # Campo de valor aberto (ex.: WERKS, BUKRS) — abre confirmação extra.
            self._step("Autorização global para (valor aberto)", "wnd[1]/usr/btnGES1", "press")
            time.sleep(0.4)
            if self._exists("wnd[2]/usr/btnBUTTON_1"):
                self._step("Confirmar 'Inserir autorização global' (Sim)", "wnd[2]/usr/btnBUTTON_1", "press")

        self._step("Confirmar popup de valores", "wnd[1]/tbar[0]/btn[0]", "press")
        return {"status": "aplicado"}

    def _executar_e_gerar_perfis(self) -> Dict[str, str]:
        self._step("Executar (relógio)", "wnd[0]/tbar[1]/btn[8]", "press")
        time.sleep(1.0)

        mt, sb = self._sbar()
        if mt in ("E", "A"):
            return {"ok": False, "message": sb or "Erro ao executar."}

        if self._exists("wnd[0]/tbar[1]/btn[20]"):
            self._step("Gerar perfis", "wnd[0]/tbar[1]/btn[20]", "press")
            time.sleep(1.0)
            if self.sess.Children.Count > 1:
                self._step("Fechar popup de geração de perfil", "wnd[1]/tbar[0]/btn[0]", "press")

        return {"ok": True, "message": sb or "Autorização global aplicada e perfil gerado."}

    def aplicar_campo_normal(self, roles: List[str], campo: str) -> Dict[str, Any]:
        """radSEL_FLD, modo 'T - Substituir tudo' — campos normais (só AGR_1251)."""
        campo = campo.strip().upper()
        print(f"\n▶ Campo NORMAL {campo} — {len(roles)} função(ões), radSEL_FLD / Substituir tudo")

        self._step("Chamar /nPFCGMASSVAL", "wnd[0]/tbar[0]/okcd", "text", "/nPFCGMASSVAL")
        self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
        self._step("Selecionar Execução Direta", "wnd[0]/usr/radMOD_EXE", "select")
        self._step("Selecionar 'Campo para vários objetos'", "wnd[0]/usr/radSEL_FLD", "select")
        self._step("Enter (atualizar ecrã)", "wnd[0]", "sendVKey", vkey=0)

        if len(roles) == 1:
            self._step("Preencher ROLE-LOW", "wnd[0]/usr/ctxtROLE-LOW", "text", roles[0])
        else:
            self._selecionar_roles_multipla(roles)

        self._step("Modo 'T - Substituir tudo'", "wnd[0]/usr/cmbFLDACT", "key", "T")
        self._step("Objeto em branco (todos os objetos)", "wnd[0]/usr/ctxtFLDOBJ", "text", "")
        self._step(f"Nome do campo = {campo}", "wnd[0]/usr/ctxtFLDFLD", "text", campo)
        self._step("Desmarcar 'Nenhuma mudança para o status Modificado'", "wnd[0]/usr/chkS_ACHD", "unselect")
        self._step("Enter (validar seleção)", "wnd[0]", "sendVKey", vkey=0)

        r = self._aplicar_global_no_popup_vals("btnPFLDN")
        if r["status"] != "aplicado":
            self._step("Voltar ao início (/N)", "wnd[0]/tbar[0]/okcd", "text", "/N")
            self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
            return {"ok": r["status"] == "skip", "skip": r["status"] == "skip", "campo": campo, "roles": roles, "message": r["message"]}

        resultado = self._executar_e_gerar_perfis()
        self._step("Voltar ao início (/N)", "wnd[0]/tbar[0]/okcd", "text", "/N")
        self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
        return {"ok": resultado["ok"], "skip": False, "campo": campo, "roles": roles, "message": resultado["message"]}

    def aplicar_campo_organizacional(self, roles: List[str], campo: str) -> Dict[str, Any]:
        """radSEL_ORG, modo 'A - Inserir' — campos de nível organizacional (também em AGR_1252)."""
        campo = campo.strip().upper()
        print(f"\n▶ Campo ORGANIZACIONAL {campo} — {len(roles)} função(ões), radSEL_ORG / Inserir")

        self._step("Chamar /nPFCGMASSVAL", "wnd[0]/tbar[0]/okcd", "text", "/nPFCGMASSVAL")
        self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
        self._step("Selecionar Execução Direta", "wnd[0]/usr/radMOD_EXE", "select")
        self._step("Selecionar 'Modificar níveis organizacionais'", "wnd[0]/usr/radSEL_ORG", "select")
        self._step("Enter (atualizar ecrã)", "wnd[0]", "sendVKey", vkey=0)

        if len(roles) == 1:
            self._step("Preencher ROLE-LOW", "wnd[0]/usr/ctxtROLE-LOW", "text", roles[0])
        else:
            self._selecionar_roles_multipla(roles)

        self._step(f"Nível organizacional = {campo}", "wnd[0]/usr/ctxtORGFLD", "text", campo)
        self._step("Modo 'A - Inserir'", "wnd[0]/usr/cmbORGACT", "key", "A")
        self._step("Enter (validar seleção)", "wnd[0]", "sendVKey", vkey=0)

        r = self._aplicar_global_no_popup_vals("btnPORGN")
        if r["status"] != "aplicado":
            self._step("Voltar ao início (/N)", "wnd[0]/tbar[0]/okcd", "text", "/N")
            self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
            return {"ok": r["status"] == "skip", "skip": r["status"] == "skip", "campo": campo, "roles": roles, "message": r["message"]}

        resultado = self._executar_e_gerar_perfis()
        self._step("Voltar ao início (/N)", "wnd[0]/tbar[0]/okcd", "text", "/N")
        self._step("Enter", "wnd[0]", "sendVKey", vkey=0)
        return {"ok": resultado["ok"], "skip": False, "campo": campo, "roles": roles, "message": resultado["message"]}


def verificar_full_auth_campo_normal(role_name: str, campo: str, ambiente: str = "PRD") -> Dict[str, Any]:
    """Somente leitura (RFC): confirma que todos os objetos com este campo (AGR_1251) ficaram com LOW = '*'."""
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, make_option_eq, read_table,
    )
    from pyrfc import Connection

    role_name = role_name.strip().upper()
    campo = campo.strip().upper()
    load_project_env(find_project_root())
    params = build_connection_params_for(ambiente)
    guard = make_read_only_guard(("AGR_1251",))
    conn = Connection(**params)
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_1251",
            fields=["OBJECT", "AUTH", "FIELD", "LOW", "DELETED"],
            options=make_option_eq("AGR_NAME", role_name),
            rowcount=0,
        )
    finally:
        conn.close()

    valores_por_objeto: Dict[str, set] = {}
    for obj, _auth, campo_linha, low, deleted in rows:
        if deleted.strip().upper() == "X":
            continue
        if campo_linha.strip().upper() != campo:
            continue
        valores_por_objeto.setdefault(obj.strip(), set()).add(low.strip())

    problemas = [
        {"objeto": obj, "valores": sorted(vals)}
        for obj, vals in sorted(valores_por_objeto.items())
        if vals != {"*"}
    ]
    objetos_ok = [obj for obj, vals in sorted(valores_por_objeto.items()) if vals == {"*"}]

    return {
        "ok": len(problemas) == 0,
        "role": role_name,
        "campo": campo,
        "total_objetos": len(valores_por_objeto),
        "objetos_ok": objetos_ok,
        "problemas": problemas,
    }


def verificar_full_auth_campo_organizacional(role_name: str, campo: str, ambiente: str = "PRD") -> Dict[str, Any]:
    """Somente leitura (RFC): confirma que a função ficou com LOW = '*' em AGR_1252 para este campo."""
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root,
        make_read_only_guard, read_table,
    )
    from pyrfc import Connection

    role_name = role_name.strip().upper()
    campo = campo.strip().upper()
    load_project_env(find_project_root())
    params = build_connection_params_for(ambiente)
    guard = make_read_only_guard(("AGR_1252",))
    conn = Connection(**params)
    try:
        rows = read_table(
            conn, guard,
            table_name="AGR_1252",
            fields=["AGR_NAME", "VARBL", "LOW", "HIGH"],
            options=[{"TEXT": f"AGR_NAME = '{role_name}' AND VARBL = '${campo}'"}],
            rowcount=1,
        )
    finally:
        conn.close()

    if not rows:
        return {"ok": False, "role": role_name, "campo": campo, "message": "Sem entrada em AGR_1252 para este campo."}

    low = rows[0][2].strip()
    return {"ok": low == "*", "role": role_name, "campo": campo, "low": low}


def verificar_full_auth_campo(role_name: str, campo: str, ambiente: str = "PRD") -> Dict[str, Any]:
    """Escolhe automaticamente a verificação certa (AGR_1252 vs AGR_1251) consoante o tipo de campo."""
    if campo_e_organizacional(campo, ambiente):
        return verificar_full_auth_campo_organizacional(role_name, campo, ambiente)
    return verificar_full_auth_campo_normal(role_name, campo, ambiente)


def aplicar_autorizacao_global(ambiente: str, roles: List[str], campo: str) -> Dict[str, Any]:
    """Ponto de entrada único: aplica autorização global ao campo indicado,
    para 1 ou várias funções, numa única execução da transação PFCGMASSVAL.
    Deteta automaticamente se o campo é organizacional (AGR_1252) ou normal
    (AGR_1251) e usa a opção certa da transação. Depois de executar, verifica
    por RFC (somente leitura) cada função individualmente."""
    campo = campo.strip().upper()
    roles = [r.strip().upper() for r in roles]

    session = _get_session(ambiente)
    e_org = campo_e_organizacional(campo, ambiente)
    print(f"Campo '{campo}' classificado como {'ORGANIZACIONAL (AGR_1252)' if e_org else 'NORMAL (AGR_1251)'}.")

    page = AutorizacaoGlobalPage(session)
    try:
        if e_org:
            r = page.aplicar_campo_organizacional(roles, campo)
        else:
            r = page.aplicar_campo_normal(roles, campo)
    except Exception as e:
        r = {"ok": False, "skip": False, "campo": campo, "roles": roles, "message": str(e)}

    resultado: Dict[str, Any] = {
        "ok": r["ok"], "skip": r.get("skip", False), "campo": campo, "roles": roles,
        "e_organizacional": e_org, "message": r["message"], "verificacao": {},
    }

    if r.get("skip") or not r["ok"]:
        return resultado

    print("A verificar via RFC (somente leitura) o resultado, função a função...")
    verificacoes = {}
    problematicas = []
    for role_name in roles:
        v = verificar_full_auth_campo(role_name, campo, ambiente)
        verificacoes[role_name] = v
        if not v["ok"]:
            problematicas.append(role_name)
        elif not e_org:
            registar_objetos_atualizados(role_name, [(obj, "*") for obj in v["objetos_ok"]])

    resultado["verificacao"] = verificacoes
    resultado["ok"] = len(problematicas) == 0
    if problematicas:
        resultado["message"] = f"{r['message']} — mas {len(problematicas)} função(ões) falharam na verificação RFC: {problematicas}"
    return resultado


def executar_individual(ambiente: str, role_name: str, campo: str = "ACTVT") -> Dict[str, Any]:
    """Compatibilidade: aplica autorização global a uma única função."""
    return aplicar_autorizacao_global(ambiente, [role_name], campo)


def executar_massa(ambiente: str, roles: List[str], campo: str = "ACTVT") -> Dict[str, Any]:
    """Aplica autorização global a várias funções numa única execução (seleção múltipla)."""
    return aplicar_autorizacao_global(ambiente, roles, campo)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--ambiente", choices=["DEV", "QAD", "PRD"], required=True)
    parser.add_argument("--role", action="append", default=[], help="Função a processar (repetível para lote).")
    parser.add_argument("--campo", default="ACTVT", help="Campo a aplicar Full Auth (ACTVT, WERKS, EKORG, BUKRS, ...).")
    args = parser.parse_args()

    roles = args.role or ["Z_PRODUCTION_ORDER_CREATE"]
    resultado = aplicar_autorizacao_global(args.ambiente, roles, args.campo)

    print("\n" + "=" * 75)
    print("RESULTADO FINAL")
    print("=" * 75)
    print(resultado)
