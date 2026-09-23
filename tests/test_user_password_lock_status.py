import sys
import types

from sap_rfc.user_create_service import _decode_user_lock_status

sys.modules.setdefault(
    "fastapi",
    types.SimpleNamespace(HTTPException=type("HTTPException", (Exception,), {})),
)
from sap_script_web_cockpit_v2.web_api.user_create_common import _safe_user_create_result


def test_decode_user_lock_status_detects_failed_logon_lock():
    result = _decode_user_lock_status("128", "3")

    assert result["locked"] is True
    assert result["login_failed_locked"] is True
    assert result["failed_logon_count"] == 3
    assert result["lock_status"] == "Bloqueado por tentativas excessivas de logon incorreto"


def test_decode_user_lock_status_combines_admin_and_failed_logon_locks():
    result = _decode_user_lock_status("160", "5")

    assert result["locked"] is True
    assert result["login_failed_locked"] is True
    assert result["failed_logon_count"] == 5
    assert result["lock_reasons"] == [
        "Bloqueado globalmente por administrador",
        "Bloqueado por tentativas excessivas de logon incorreto",
    ]


def test_decode_user_lock_status_active_user():
    result = _decode_user_lock_status("0", "0")

    assert result["locked"] is False
    assert result["login_failed_locked"] is False
    assert result["failed_logon_count"] == 0
    assert result["lock_status"] == "Ativo"


def test_safe_user_create_result_keeps_password_change_lock_fields():
    safe = _safe_user_create_result(
        {
            "ok": True,
            "status": "PASSWORD_CHANGED",
            "environment": "PRD",
            "username": "S123",
            "password_source": "SAP_PASSE_PASSWD (.env)",
            "uflag": "128",
            "failed_logon_count": 3,
            "locked": True,
            "login_failed_locked": True,
            "lock_reasons": ["Bloqueado por tentativas excessivas de logon incorreto"],
            "lock_status": "Bloqueado por tentativas excessivas de logon incorreto",
            "message": "Password alterada, mas o utilizador continua bloqueado.",
            "password": "never-return-this",
        }
    )

    assert safe["login_failed_locked"] is True
    assert safe["lock_status"] == "Bloqueado por tentativas excessivas de logon incorreto"
    assert safe["failed_logon_count"] == 3
    assert "password" not in safe


def test_safe_user_create_result_keeps_unlock_lock_fields():
    safe = _safe_user_create_result(
        {
            "ok": True,
            "status": "USER_UNLOCKED",
            "environment": "PRD",
            "username": "S123",
            "uflag": "0",
            "failed_logon_count": 0,
            "locked": False,
            "login_failed_locked": False,
            "lock_reasons": [],
            "lock_status": "Ativo",
            "message": "Utilizador desbloqueado e sem bloqueio ativo em USR02.",
        }
    )

    assert safe["status"] == "USER_UNLOCKED"
    assert safe["locked"] is False
    assert safe["login_failed_locked"] is False
    assert safe["lock_status"] == "Ativo"


def test_safe_user_create_result_keeps_password_changed_and_unlocked():
    safe = _safe_user_create_result(
        {
            "ok": True,
            "status": "PASSWORD_CHANGED_AND_UNLOCKED",
            "environment": "PRD",
            "username": "S123",
            "password_source": "SAP_PASSE_PASSWD (.env)",
            "uflag": "0",
            "failed_logon_count": 0,
            "locked": False,
            "login_failed_locked": False,
            "unlocked": True,
            "lock_reasons": [],
            "lock_status": "Ativo",
            "message": "Password alterada e utilizador desbloqueado com sucesso (USR02 sem bloqueio ativo).",
            "password": "never-return-this",
        }
    )

    assert safe["status"] == "PASSWORD_CHANGED_AND_UNLOCKED"
    assert safe["unlocked"] is True
    assert safe["locked"] is False
    assert safe["lock_status"] == "Ativo"
    assert "password" not in safe


def _setup_mock_rfc(monkeypatch, *, usr02_rows=None, unlock_return=None, change_return=None):
    import sap_rfc.user_create_service as ucs

    monkeypatch.setattr(ucs, "find_project_root", lambda: None)
    monkeypatch.setattr(ucs, "load_project_env", lambda root: None)
    monkeypatch.setattr(
        ucs,
        "build_connection_params_for_env",
        lambda env: {"user": "U", "passwd": "P", "ashost": "H", "sysnr": "00", "client": "100"},
    )

    rows_queue = list(usr02_rows or [["TESTUSER", "0", "0"]])
    call_log = []

    class MockConnection:
        def __init__(self, **kwargs):
            self.kwargs = kwargs
            self.closed = False

        def call(self, func_name, **kwargs):
            call_log.append((func_name, kwargs))
            if func_name == "RFC_PING":
                return {}
            if func_name == "BAPI_USER_EXISTENCE_CHECK":
                return {"RETURN": {"NUMBER": "088"}}
            if func_name == "BAPI_USER_CHANGE":
                return change_return if change_return is not None else {"RETURN": []}
            if func_name == "BAPI_USER_UNLOCK":
                return unlock_return if unlock_return is not None else {"RETURN": []}
            if func_name == "BAPI_TRANSACTION_COMMIT":
                return {}
            if func_name == "BAPI_TRANSACTION_ROLLBACK":
                return {}
            if func_name == "RFC_READ_TABLE":
                row = rows_queue.pop(0) if rows_queue else ["TESTUSER", "0", "0"]
                return {"DATA": [{"WA": "|".join(row)}]}
            return {}

        def close(self):
            self.closed = True

    fake_pyrfc = types.ModuleType("pyrfc")
    fake_pyrfc.Connection = MockConnection
    monkeypatch.setitem(sys.modules, "pyrfc", fake_pyrfc)

    return call_log


def test_change_password_user_already_unlocked(monkeypatch):
    from sap_rfc.user_create_service import change_password_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[["TESTUSER", "0", "0"]],
    )

    result = change_password_rfc("PRD", "TESTUSER", "TempPass123!")

    assert result["ok"] is True
    assert result["status"] == "PASSWORD_CHANGED"
    assert result["locked"] is False
    assert result["login_failed_locked"] is False
    assert result.get("unlocked") is False
    assert result["message"] == "Password alterada e utilizador sem bloqueio ativo em USR02."

    called_functions = [name for name, _ in call_log]
    assert "BAPI_USER_CHANGE" in called_functions
    assert "BAPI_USER_UNLOCK" not in called_functions


def test_change_password_unlocks_locked_user_and_confirms(monkeypatch):
    from sap_rfc.user_create_service import change_password_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[
            ["TESTUSER", "128", "3"],
            ["TESTUSER", "0", "0"],
        ],
    )

    result = change_password_rfc("PRD", "TESTUSER", "TempPass123!")

    assert result["ok"] is True
    assert result["status"] == "PASSWORD_CHANGED_AND_UNLOCKED"
    assert result["locked"] is False
    assert result["login_failed_locked"] is False
    assert result["unlocked"] is True
    assert result["uflag"] == "0"
    assert result["failed_logon_count"] == 0
    assert result["lock_status"] == "Ativo"
    assert "desbloqueado com sucesso" in result["message"]

    called_functions = [name for name, _ in call_log]
    assert "BAPI_USER_CHANGE" in called_functions
    assert "BAPI_USER_UNLOCK" in called_functions


def test_independent_unlock_user_rfc_success(monkeypatch):
    from sap_rfc.user_create_service import unlock_user_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[["TESTUSER", "0", "0"]],
    )

    result = unlock_user_rfc("PRD", "TESTUSER")

    assert result["ok"] is True
    assert result["status"] == "USER_UNLOCKED"
    assert result["locked"] is False
    assert result["message"] == "Utilizador desbloqueado e sem bloqueio ativo em USR02."

    called_functions = [name for name, _ in call_log]
    assert "BAPI_USER_UNLOCK" in called_functions
    assert "BAPI_USER_CHANGE" not in called_functions


def test_independent_unlock_user_rfc_bapi_failure(monkeypatch):
    from sap_rfc.user_create_service import unlock_user_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        unlock_return={"RETURN": [{"TYPE": "E", "MESSAGE": "Sem autorização para desbloquear"}]},
    )

    result = unlock_user_rfc("PRD", "TESTUSER")

    assert result["ok"] is False
    assert result["status"] == "UNLOCK_FAILED"
    assert result["sap_return_messages"] == ["Sem autorização para desbloquear"]
    assert "Não foi possível desbloquear o utilizador" in result["message"]

    called_functions = [name for name, _ in call_log]
    assert "BAPI_USER_UNLOCK" in called_functions
    assert "BAPI_TRANSACTION_ROLLBACK" in called_functions


def test_change_password_unlock_bapi_failure(monkeypatch):
    from sap_rfc.user_create_service import change_password_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[["TESTUSER", "128", "3"]],
        unlock_return={"RETURN": [{"TYPE": "E", "MESSAGE": "Sem autorização para desbloquear"}]},
    )

    result = change_password_rfc("PRD", "TESTUSER", "TempPass123!")

    assert result["ok"] is True
    assert result["status"] == "PASSWORD_CHANGED"
    assert result["unlocked"] is False
    assert result["locked"] is True
    assert result["login_failed_locked"] is True
    assert "Sem autorização para desbloquear" in result.get("unlock_sap_return_messages", [])
    assert "não foi possível desbloquear o utilizador" in result["message"]


def test_change_password_unlock_not_confirmed(monkeypatch):
    from sap_rfc.user_create_service import change_password_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[
            ["TESTUSER", "128", "3"],
            ["TESTUSER", "128", "3"],
        ],
        unlock_return={"RETURN": []},
    )

    result = change_password_rfc("PRD", "TESTUSER", "TempPass123!")

    assert result["ok"] is True
    assert result["status"] == "PASSWORD_CHANGED"
    assert result["unlocked"] is False
    assert result["locked"] is True
    assert result["login_failed_locked"] is True
    assert "continua bloqueado" in result["message"]


def test_independent_unlock_not_confirmed_when_usr02_still_locked(monkeypatch):
    from sap_rfc.user_create_service import unlock_user_rfc

    call_log = _setup_mock_rfc(
        monkeypatch,
        usr02_rows=[["TESTUSER", "32", "0"]],
    )

    result = unlock_user_rfc("PRD", "TESTUSER")

    assert result["ok"] is False
    assert result["status"] == "UNLOCK_NOT_CONFIRMED"
    assert result["locked"] is True
    assert "continua bloqueado" in result["message"]

