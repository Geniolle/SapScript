"""Testes da descoberta DDIC e da heurística de campos."""

from sap_payroll_analysis.ddic import describe_table, guess_fields
from sap_payroll_analysis.tests.fakes import FakeConnection

ACDOCA_FIELDS = [
    ("RCLNT", "C", 3, "Mandante"),
    ("RLDNR", "C", 2, "Ledger"),
    ("RBUKRS", "C", 4, "Empresa"),
    ("GJAHR", "N", 4, "Exercício"),
    ("RYEAR", "N", 4, "Ano do exercício"),
    ("POPER", "N", 3, "Período"),
    ("BELNR", "C", 10, "Nº documento"),
    ("DOCLN", "C", 6, "Nº linha"),
    ("RACCT", "C", 10, "Conta do Razão"),
    ("DRCRK", "C", 1, "Débito/Crédito"),
    ("WSL", "P", 23, "Montante em moeda transação"),
    ("RWCUR", "C", 5, "Moeda transação"),
    ("BUDAT", "D", 8, "Data de lançamento"),
]


def test_describe_table_via_read_table_meta():
    conn = FakeConnection({"ACDOCA": {"fields": ACDOCA_FIELDS, "rows": []}})
    diag = describe_table(conn, "ACDOCA")
    assert diag.exists and diag.authorized
    assert diag.field_count == len(ACDOCA_FIELDS)
    assert "RACCT" in diag.field_names()


def test_describe_table_not_available():
    # FAGLFLEXA está na whitelist mas o "sistema" não a tem.
    conn = FakeConnection({}, raise_for={"FAGLFLEXA": "TABLE_NOT_AVAILABLE"})
    diag = describe_table(conn, "FAGLFLEXA")
    assert diag.exists is False


def test_guess_fields_acdoca():
    conn = FakeConnection({"ACDOCA": {"fields": ACDOCA_FIELDS, "rows": []}})
    diag = describe_table(conn, "ACDOCA")
    g = guess_fields(diag)
    assert g["empresa"].chosen == "RBUKRS"
    assert g["conta"].chosen == "RACCT"
    assert g["moeda"].chosen == "RWCUR"
    assert g["debito_credito"].chosen == "DRCRK"
    assert g["periodo"].chosen == "POPER"
    assert g["valor"].chosen == "WSL"
    assert g["exercicio"].chosen in {"RYEAR", "GJAHR"}


def test_guess_fields_bseg_like():
    fields = [
        ("BUKRS", "C", 4, "Empresa"),
        ("BELNR", "C", 10, "Documento"),
        ("GJAHR", "N", 4, "Exercício"),
        ("BUZEI", "N", 3, "Linha"),
        ("HKONT", "C", 10, "Conta do Razão"),
        ("SHKZG", "C", 1, "Débito/Crédito"),
        ("WRBTR", "P", 13, "Montante"),
        ("PSWSL", "C", 5, "Moeda"),
    ]
    conn = FakeConnection({"BSEG": {"fields": fields, "rows": []}})
    diag = describe_table(conn, "BSEG")
    g = guess_fields(diag)
    assert g["empresa"].chosen == "BUKRS"
    assert g["conta"].chosen == "HKONT"
    assert g["valor"].chosen == "WRBTR"
    assert g["debito_credito"].chosen == "SHKZG"
    assert g["item"].chosen == "BUZEI"
