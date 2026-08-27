"""Atribuição de uma role/perfil PFCG a uma Request de transporte via SAP GUI Scripting.

Motivo desta abordagem: confirmado empiricamente (ver `pfcg_transport_service.py` e os testes
documentados no histórico do projeto) que não existe caminho por RFC/BAPI capaz de atribuir uma
role já criada a uma Request de transporte:
  - `PRGN_RFC_CREATE_ACTIVITY_GROUP` aceita um parâmetro REQUEST mas nunca grava o objeto em
    E071/E071K, em nenhum cenário testado.
  - `TR_OBJECT_INSERT` (inserir objeto existente numa request) não é remote-enabled.
  - `PFCG_MASS_TRANSPORT` (o report por trás do botão "Transportar" da PFCG) bloqueia
    explicitamente a própria execução em background, mesmo via BAPI_XBP_* (job real, testado
    e confirmado: mensagem "Não é possível chamar report PFCG_MASS_TRANSPORT em job de
    background").

A única via funcional restante é reproduzir, via SAP GUI Scripting, o mesmo caminho que um
utilizador faria manualmente: executar o report PFCG_MASS_TRANSPORT em primeiro plano (SE38),
preenchendo o ecrã de seleção com a role e a Request pretendidas. Se não houver uma sessão SAP
GUI ativa e ligada ao ambiente pedido, a função devolve um status que sinaliza ao chamador para
usar o fallback manual (instruir o utilizador a fazer a atribuição via GUI, ele próprio).
"""
from __future__ import annotations

import os
import re
import time
from typing import Any

from sap_rfc._rfc_common import find_project_root, load_project_env

REPORT_NAME = "PFCG_MASS_TRANSPORT"

# Ambientes com escrita permitida via automação GUI — mesma restrição aplicada em
# pfcg_role_create_service.py e pfcg_transport_service.py (apenas DEV nesta fase).
ALLOWED_GUI_TRANSPORT_ENVIRONMENTS = ("DEV",)


def _assert_environment_allowed(environment: str) -> str:
    env = str(environment or "").strip().upper()
    if env not in ALLOWED_GUI_TRANSPORT_ENVIRONMENTS:
        raise ValueError(
            f"Atribuição via GUI só é permitida em DEV nesta fase. "
            f"Ambiente '{env}' está bloqueado (QAD e PRD não são permitidos para escrita)."
        )
    return env


def _pump_com_messages(duration: float = 0.3) -> None:
    """Esvazia repetidamente a fila de mensagens COM da thread Python durante `duration` segundos.
    Necessário porque ações de automação (`.text =`, `.selected =`, `sendVKey`, `.press()`) que
    despoletam processamento no SAP GUI ficam "pendentes" até a thread chamadora bombear a fila —
    confirmado ao vivo, de forma repetida: um único `PumpWaitingMessages()` isolado não é
    suficiente quando há várias ações por bombear acumuladas (o backlog de ações anteriores não
    bombeadas impede o processamento da ação seguinte, mesmo com sleeps longos entre elas); é
    preciso bombear continuamente durante uma janela de tempo, não apenas uma vez por ação."""
    try:
        import pythoncom  # type: ignore
    except Exception:
        return
    deadline = time.time() + duration
    while time.time() < deadline:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            return
        time.sleep(0.05)


def _find_matching_session(environment: str) -> Any:
    """Procura, entre as sessões SAP GUI já abertas no SAP Logon, uma ligada ao `environment`
    pedido. Devolve `None` se não encontrar nenhuma — não abre nem faz logon numa sessão nova.

    Critério de correspondência (dois filtros obrigatórios, aplicados em conjunto):
      1. Client (mandante) da sessão == SAP_{ENV}_CLIENT do `.env` — a validação essencial
         pedida explicitamente: nunca operar numa janela que não seja do mandante correto.
      2. O nome do `environment` (ex.: "DEV") aparece na descrição da conexão no SAP Logon
         (ex.: "SAP S4F DEV") — usado para identificar o sistema, porque o ApplicationServer
         que a sessão reporta é o hostname real do servidor (ex.: "sjsaps4hapd01"), que não
         corresponde ao ASHOST configurado no `.env` (lá é o IP, ex.: "172.19.66.4") — os dois
         nunca coincidem, então comparar por ASHOST/SYSNR não é fiável e não é usado aqui.
    """
    import win32com.client  # type: ignore

    client = os.environ.get(f"SAP_{environment}_CLIENT", "").strip()
    env_label = environment.strip().lower()

    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine

    for i in range(application.Children.Count):
        connection = application.Children(i)
        description = str(getattr(connection, "Description", "") or "").lower()
        if env_label not in description:
            continue
        for j in range(connection.Children.Count):
            session = connection.Children(j)
            info = session.Info
            if client and str(info.Client) != client:
                continue
            return session
    return None


def get_gui_session_for_environment(environment: str) -> dict[str, Any]:
    """Diagnóstico read-only: verifica se SAP GUI Scripting está disponível e se existe uma
    sessão já aberta e ligada ao `environment` pedido, sem tocar em nada."""
    try:
        env = _assert_environment_allowed(environment)
    except ValueError as exc:
        return {"ok": False, "status": "ENVIRONMENT_BLOCKED", "message": str(exc)}

    try:
        project_root = find_project_root()
        load_project_env(project_root)
    except Exception as exc:
        return {"ok": False, "status": "CONFIG_ERROR", "message": str(exc)}

    try:
        import win32com.client  # noqa: F401
    except Exception as exc:
        return {
            "ok": False,
            "status": "GUI_SCRIPTING_UNAVAILABLE",
            "message": f"pywin32 não disponível neste interpretador Python: {exc}",
        }

    try:
        session = _find_matching_session(env)
    except Exception as exc:
        return {
            "ok": False,
            "status": "GUI_SCRIPTING_UNAVAILABLE",
            "message": f"SAP GUI Scripting não disponível/ativo (SAP Logon aberto? scripting "
                       f"habilitado?): {exc}",
        }

    if session is None:
        return {
            "ok": False,
            "status": "NO_ACTIVE_GUI_SESSION",
            "environment": env,
            "message": f"Nenhuma sessão SAP GUI ativa ligada ao ambiente {env} foi encontrada. "
                       f"Abra e faça logon numa sessão SAP GUI nesse sistema/client antes de tentar.",
        }

    info = session.Info
    return {
        "ok": True,
        "status": "SESSION_FOUND",
        "environment": env,
        "system_name": str(info.SystemName),
        "client": str(info.Client),
        "user": str(info.User),
        "transaction": str(info.Transaction),
    }


def assign_role_to_transport_via_gui(environment: str, role: str, request: str) -> dict[str, Any]:
    """Tenta atribuir `role` à `request` via SAP GUI Scripting, executando PFCG_MASS_TRANSPORT
    em primeiro plano (SE38) com TESTMODE desligado (execução real, não simulação).

    Devolve status GUI_SCRIPTING_UNAVAILABLE / NO_ACTIVE_GUI_SESSION quando a automação não pode
    ser tentada — nesses casos o chamador deve usar o fallback manual (instruir o utilizador)."""
    try:
        env = _assert_environment_allowed(environment)
    except ValueError as exc:
        return {"ok": False, "status": "ENVIRONMENT_BLOCKED", "message": str(exc)}

    role_clean = str(role or "").strip().upper()
    request_clean = str(request or "").strip().upper()
    if not role_clean or not request_clean:
        return {"ok": False, "status": "INVALID_INPUT", "message": "Informe role e request."}

    diag = get_gui_session_for_environment(env)
    if not diag.get("ok"):
        return diag

    try:
        session = _find_matching_session(env)
    except Exception as exc:
        return {"ok": False, "status": "GUI_SCRIPTING_UNAVAILABLE", "message": str(exc)}

    if session is None:
        return {
            "ok": False,
            "status": "NO_ACTIVE_GUI_SESSION",
            "environment": env,
            "message": "Sessão deixou de estar disponível entre a verificação e a execução.",
        }

    try:
        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE38"
        session.findById("wnd[0]").sendVKey(0)
        _pump_com_messages()

        session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = REPORT_NAME
        session.findById("wnd[0]").sendVKey(8)  # Executar (F8)
        _pump_com_messages()

        session.findById("wnd[0]/usr/ctxtAGR_NAME-LOW").text = role_clean
        _pump_com_messages()
        _set_checkbox(session, "wnd[0]/usr/chkCOMP_ROL", True)
        _set_checkbox(session, "wnd[0]/usr/chkPROFILES", True)
        _set_checkbox(session, "wnd[0]/usr/chkPERSON", False)
        _set_checkbox(session, "wnd[0]/usr/chkUSERS", False)
        # TESTMODE/CSOL_EVA só existem no ecrã quando GF_SCC4_ACTV <> SPACE (recording
        # automático/manual "com exceção"); confirmado ao vivo: neste client (GF_SCC4_ACTV =
        # SPACE, PRGN_CHK_CLIENT_CUST_SETTING termina com SY-SUBRC=0) o próprio report esconde e
        # limpa os dois campos sozinho (força execução real, não fica preso em modo simulação) —
        # por isso são best-effort aqui, não obrigatórios.
        _set_checkbox(session, "wnd[0]/usr/chkTESTMODE", False)
        _set_checkbox(session, "wnd[0]/usr/chkCSOL_EVA", False)
        _pump_com_messages()

        session.findById("wnd[0]").sendVKey(8)  # Executar (F8)
        _pump_com_messages(0.5)

        popup_result = _handle_request_popup(session, request_clean)
        if popup_result is not None:
            return popup_result

        grid_result = _read_result_grid(session)
        if grid_result is None:
            return {
                "ok": False,
                "status": "NEEDS_LIVE_VALIDATION",
                "environment": env,
                "role": role_clean,
                "request": request_clean,
                "message": (
                    "PFCG_MASS_TRANSPORT foi executado, mas não encontrei o grid de resultado "
                    "esperado (ecrã 'Transporte de funções') para confirmar o desfecho. Valide "
                    "manualmente e confirme com list_transport_request_objects()."
                ),
            }

        matching_rows = [row for row in grid_result if row["role"] == role_clean]
        if not matching_rows:
            return {
                "ok": False,
                "status": "NEEDS_LIVE_VALIDATION",
                "environment": env,
                "role": role_clean,
                "request": request_clean,
                "grid": grid_result,
                "message": "Grid de resultado não contém uma linha para a role pedida.",
            }

        row = matching_rows[0]
        if row["err_text"]:
            return {
                "ok": False,
                "status": "GUI_AUTOMATION_ERROR",
                "environment": env,
                "role": role_clean,
                "request": request_clean,
                "grid_row": row,
                "message": f"PFCG_MASS_TRANSPORT reportou erro para a role: {row['err_text']}",
            }

        return {
            "ok": True,
            "status": "ASSIGNED_VIA_GUI",
            "environment": env,
            "role": role_clean,
            "request": request_clean,
            "status_text": row["status_text"],
            "message": "PFCG_MASS_TRANSPORT executado em primeiro plano via SAP GUI Scripting. "
                       "Confirme sempre com list_transport_request_objects().",
        }
    except Exception as exc:
        return {
            "ok": False,
            "status": "GUI_AUTOMATION_ERROR",
            "environment": env,
            "role": role_clean,
            "request": request_clean,
            "message": f"Falha ao automatizar PFCG_MASS_TRANSPORT via GUI: {exc}",
        }


FALLBACK_TRIGGER_STATUSES = ("GUI_SCRIPTING_UNAVAILABLE", "NO_ACTIVE_GUI_SESSION")


def assign_role_to_transport(environment: str, role: str, request: str) -> dict[str, Any]:
    """Passo 1: tenta atribuir via SAP GUI Scripting. Se o script não estiver ativo (pywin32
    indisponível ou nenhuma sessão SAP GUI aberta e ligada ao ambiente), cai automaticamente
    para o Passo 2 (fallback): devolve as instruções para o utilizador fazer a atribuição
    manualmente na PFCG, com a role e a Request já identificadas."""
    result = assign_role_to_transport_via_gui(environment, role, request)
    if result.get("ok") or result.get("status") not in FALLBACK_TRIGGER_STATUSES:
        return result

    env = str(environment or "").strip().upper()
    role_clean = str(role or "").strip().upper()
    request_clean = str(request or "").strip().upper()
    return {
        "ok": False,
        "status": "MANUAL_FALLBACK_REQUIRED",
        "environment": env,
        "role": role_clean,
        "request": request_clean,
        "gui_attempt_status": result.get("status"),
        "message": (
            f"Automação via SAP GUI Scripting não está disponível agora ({result.get('status')}). "
            f"Faça a atribuição manualmente: transação PFCG -> função {role_clean} -> Modificar -> "
            f"menu Função/Utilidades -> Transportar -> selecione a Request {request_clean} já criada."
        ),
    }


def _set_checkbox(session: Any, element_id: str, checked: bool) -> None:
    try:
        session.findById(element_id).selected = checked
    except Exception:
        pass  # campo pode não existir neste ecrã (ex.: TESTMODE/CSOL_EVA escondidos)


# Popup padrão do framework de Customizing Transport ("Consulta ordem de customizing"),
# validado ao vivo: campo de texto com o número da Request e botão "Avançar (ENTER)" para
# confirmar. Aparece sempre que a Request "própria" corrente (default do utilizador) é
# diferente da Request pedida — por isso o campo já vem preenchido com outra Request e
# precisa de ser substituído antes de confirmar.
_REQUEST_POPUP_FIELD_ID = "wnd[1]/usr/ctxtKO008-TRKORR"
_REQUEST_POPUP_CONFIRM_BTN_ID = "wnd[1]/tbar[0]/btn[0]"

# Grid de resultado do ecrã "Transporte de funções", validado ao vivo (ALV Grid Control).
_RESULT_GRID_ID = "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell"
_STATUS_TEXT_RE = re.compile(r"\\Q(.*?)@?$")


def _handle_request_popup(session: Any, request_number: str) -> dict[str, Any] | None:
    """Preenche e confirma o popup "Consulta ordem de customizing" com a Request pedida, se
    aparecer. Devolve `None` quando não há popup a tratar (fluxo segue normalmente). Devolve um
    resultado de erro apenas se o popup aparecer num formato inesperado (não arrisca cliques às
    cegas) ou se continuar aberto depois de tentar confirmar (ex.: Request inválida/bloqueada)."""
    try:
        session.findById("wnd[1]")
    except Exception:
        return None  # sem popup — ex.: Request "própria" corrente já é a pedida

    try:
        field = session.findById(_REQUEST_POPUP_FIELD_ID)
        field.text = request_number
        _pump_com_messages()
        session.findById(_REQUEST_POPUP_CONFIRM_BTN_ID).press()
    except Exception as exc:
        return {
            "ok": False,
            "status": "NEEDS_LIVE_VALIDATION",
            "message": f"Popup de Request apareceu num formato inesperado: {exc}",
        }

    for _ in range(20):  # até ~10s (0.5s de bombagem por iteração)
        _pump_com_messages(0.5)
        try:
            session.findById("wnd[1]")
        except Exception:
            return None  # popup fechou -> confirmado com sucesso

    status_text = ""
    try:
        status_text = session.findById("wnd[0]/sbar").Text
    except Exception:
        pass
    return {
        "ok": False,
        "status": "GUI_AUTOMATION_ERROR",
        "message": f"Popup de Request continuou aberto após confirmar {request_number}. "
                   f"Status bar: {status_text!r}",
    }


def _read_result_grid(session: Any) -> list[dict[str, str]] | None:
    """Lê o grid ALV do ecrã de resultado ("Transporte de funções"): colunas STATUS, ROLE,
    ROLE_TYPE, TEXT, ERR_TEXT. Devolve `None` se o grid não existir neste ecrã (ex.: layout
    diferente do validado)."""
    try:
        grid = session.findById(_RESULT_GRID_ID)
        row_count = grid.RowCount
    except Exception:
        return None

    rows: list[dict[str, str]] = []
    for r in range(row_count):
        try:
            raw_status = str(grid.GetCellValue(r, "STATUS"))
            match = _STATUS_TEXT_RE.search(raw_status)
            status_text = match.group(1) if match else raw_status
            rows.append(
                {
                    "status_text": status_text,
                    "role": str(grid.GetCellValue(r, "ROLE")).strip().upper(),
                    "err_text": str(grid.GetCellValue(r, "ERR_TEXT")).strip(),
                }
            )
        except Exception:
            continue
    return rows
