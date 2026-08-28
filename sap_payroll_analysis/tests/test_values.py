"""Testes de conversão de valor SAP e normalização de sinal (sem SAP)."""

from decimal import Decimal

import pytest

from sap_payroll_analysis.config import pad_account, pad_run
from sap_payroll_analysis.sap_reader import normalize_sign, sap_str_to_decimal


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("", Decimal("0")),
        ("   ", Decimal("0")),
        ("-", Decimal("0")),
        ("727258.35", Decimal("727258.35")),
        ("  727258.35  ", Decimal("727258.35")),
        ("727258.35-", Decimal("-727258.35")),
        ("-727258.35", Decimal("-727258.35")),
        ("+3211.71", Decimal("3211.71")),
        ("(3211.71)", Decimal("-3211.71")),
        ("1.234.567,89", Decimal("1234567.89")),
        ("1,234,567.89", Decimal("1234567.89")),
        ("724046,64", Decimal("724046.64")),
        ("0", Decimal("0")),
        ("1000", Decimal("1000")),
    ],
)
def test_sap_str_to_decimal(raw, expected):
    assert sap_str_to_decimal(raw) == expected


def test_sap_str_to_decimal_invalid():
    with pytest.raises(ValueError):
        sap_str_to_decimal("abc")


@pytest.mark.parametrize(
    "amount, flag, expected",
    [
        (Decimal("100"), "S", Decimal("100")),
        (Decimal("100"), "H", Decimal("-100")),
        (Decimal("-100"), "S", Decimal("100")),
        (Decimal("-100"), "H", Decimal("-100")),
        (Decimal("100"), "D", Decimal("100")),
        (Decimal("100"), "C", Decimal("-100")),
        (Decimal("100"), "", Decimal("100")),
        (Decimal("-42"), "", Decimal("-42")),
    ],
)
def test_normalize_sign(amount, flag, expected):
    assert normalize_sign(amount, flag) == expected


def test_pad_account():
    assert pad_account("23120000") == "0023120000"
    assert pad_account("0023120000") == "0023120000"
    assert pad_account(" 23120000 ") == "0023120000"
    assert pad_account("SECO-01") == "SECO-01"


def test_pad_run():
    assert pad_run("1296") == "0000001296"
    assert pad_run("0000001296") == "0000001296"
