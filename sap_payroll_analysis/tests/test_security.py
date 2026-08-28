"""Testes das whitelists de segurança e do wrapper safe_rfc_call."""

import pytest

from sap_payroll_analysis.security import (
    SecurityError,
    assert_function_allowed,
    assert_table_allowed,
    safe_rfc_call,
)
from sap_payroll_analysis.tests.fakes import FakeConnection


def test_allowed_functions_pass():
    for name in ("RFC_PING", "RFC_READ_TABLE", "RFC_GET_FUNCTION_INTERFACE"):
        assert_function_allowed(name)


@pytest.mark.parametrize(
    "name",
    [
        "BAPI_TRANSACTION_COMMIT",
        "RFC_DELETE_TABLE",
        "BAPI_ACC_DOCUMENT_POST",
        "PRGN_ACTIVITY_GROUP_DELETE",
        "SOME_UPDATE_FUNC",
        "JOB_SUBMIT",
        "",
    ],
)
def test_blocked_functions_raise(name):
    with pytest.raises(SecurityError):
        assert_function_allowed(name)


def test_table_whitelist():
    assert_table_allowed("PPDIT")
    assert_table_allowed("acdoca")
    with pytest.raises(SecurityError):
        assert_table_allowed("USR02")
    with pytest.raises(SecurityError):
        assert_table_allowed("")


def test_safe_rfc_call_blocks_non_whitelisted_table():
    conn = FakeConnection({"PPDIT": {"fields": [("RUNID", "C", 10, "")], "rows": []}})
    with pytest.raises(SecurityError):
        safe_rfc_call(conn, "RFC_READ_TABLE", QUERY_TABLE="USR02")
    # a chamada nunca chega ao "SAP"
    assert conn.calls == []


def test_safe_rfc_call_blocks_write_function():
    conn = FakeConnection({})
    with pytest.raises(SecurityError):
        safe_rfc_call(conn, "BAPI_ACC_DOCUMENT_POST")
    assert conn.calls == []


def test_safe_rfc_call_allows_read():
    conn = FakeConnection({"PPDIT": {"fields": [("RUNID", "C", 10, "")], "rows": []}})
    out = safe_rfc_call(conn, "RFC_READ_TABLE", QUERY_TABLE="PPDIT", FIELDS=[], OPTIONS=[], ROWCOUNT=1)
    assert "DATA" in out
    assert conn.calls and conn.calls[0][0] == "RFC_READ_TABLE"
