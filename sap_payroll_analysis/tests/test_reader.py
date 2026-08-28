"""Testes de paginação, parsing e reclassificação de erros do leitor."""

import pytest

from sap_payroll_analysis.sap_reader import (
    BufferExceeded,
    NotAuthorized,
    RfcReadError,
    TableNotAvailable,
    opt_and,
    opt_eq,
    opt_in,
    read_table,
)
from sap_payroll_analysis.tests.fakes import FakeConnection

FIELDS = [("BUKRS", "C", 4, "Empresa"), ("HKONT", "C", 10, "Conta"), ("WRBTR", "P", 13, "Valor")]


def _conn(nrows: int) -> FakeConnection:
    rows = [{"BUKRS": "2010", "HKONT": "0023120000", "WRBTR": f"{i}.00"} for i in range(nrows)]
    return FakeConnection({"BSEG": {"fields": FIELDS, "rows": rows}})


def test_pagination_multiple_pages():
    conn = _conn(12)
    res = read_table(conn, "BSEG", fields=["BUKRS", "HKONT", "WRBTR"], page_size=5)
    assert len(res.rows) == 12
    assert res.pages == 3  # 5 + 5 + 2
    assert res.rows[0]["BUKRS"] == "2010"
    assert res.rows[-1]["WRBTR"] == "11.00"


def test_pagination_exact_multiple():
    conn = _conn(10)
    res = read_table(conn, "BSEG", fields=["BUKRS"], page_size=5)
    # 5 + 5 + 1 (página vazia final para confirmar fim)
    assert len(res.rows) == 10
    assert res.pages == 3


def test_max_rows_truncates():
    conn = _conn(100)
    res = read_table(conn, "BSEG", fields=["BUKRS"], page_size=10, max_rows=25)
    assert len(res.rows) == 25
    assert res.truncated is True


def test_field_metadata_parsed():
    conn = _conn(1)
    res = read_table(conn, "BSEG")
    assert [m["FIELDNAME"] for m in res.field_meta] == ["BUKRS", "HKONT", "WRBTR"]


@pytest.mark.parametrize(
    "key, exc",
    [
        ("TABLE_NOT_AVAILABLE", TableNotAvailable),
        ("NOT_AUTHORIZED", NotAuthorized),
        ("DATA_BUFFER_EXCEEDED", BufferExceeded),
        ("FIELD_NOT_VALID", RfcReadError),
    ],
)
def test_error_classification(key, exc):
    conn = FakeConnection({"BSEG": {"fields": FIELDS, "rows": []}}, raise_for={"BSEG": key})
    with pytest.raises(exc):
        read_table(conn, "BSEG", fields=["BUKRS"])


def test_opt_in_empty_is_never_true():
    opts = opt_in("RUNID", [])
    assert opts == [{"TEXT": "1 = 2"}]


def test_opt_and_structure():
    opts = opt_and(opt_eq("BUKRS", "2010"), opt_in("BELNR", ["1", "2"]))
    texts = [o["TEXT"] for o in opts]
    assert "BUKRS = '2010'" in texts
    assert "AND" in texts
    assert texts.count("(") == 1 and texts.count(")") == 1


def test_opt_eq_escapes_quote():
    assert opt_eq("X", "O'Brien") == [{"TEXT": "X = 'O''Brien'"}]
