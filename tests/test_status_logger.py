# -*- coding: utf-8 -*-
"""
tests/test_status_logger.py
Testes unitários para o módulo centralizador sap_rfc.status_logger.
"""

from sap_rfc.status_logger import (
    format_status_success,
    format_status_warning,
    format_status_error,
    format_status_info,
    print_status_success,
    print_status_warning,
    print_status_error,
    print_status_info,
    CLR_VERDE,
    CLR_AMARELO,
    CLR_VERMELHO,
    CLR_RESET,
)


def test_format_status_success():
    res = format_status_success("✓ [2/8] Sucesso")
    assert res.startswith(CLR_VERDE)
    assert res.endswith(CLR_RESET)
    assert "✓ [2/8] Sucesso" in res


def test_format_status_warning():
    res = format_status_warning("⚠️ [1/8] Aviso")
    assert res.startswith(CLR_AMARELO)
    assert res.endswith(CLR_RESET)
    assert "⚠️ [1/8] Aviso" in res


def test_format_status_error():
    res = format_status_error("✗ [2/8] Erro")
    assert res.startswith(CLR_VERMELHO)
    assert res.endswith(CLR_RESET)
    assert "✗ [2/8] Erro" in res


def test_print_status_wrappers(capsys):
    print_status_success("Linha Verde")
    out = capsys.readouterr().out
    assert "Linha Verde" in out

    print_status_warning("Linha Amarela")
    out = capsys.readouterr().out
    assert "Linha Amarela" in out

    print_status_error("Linha Vermelha")
    out = capsys.readouterr().out
    assert "Linha Vermelha" in out

    print_status_info("Linha Info")
    out = capsys.readouterr().out
    assert "Linha Info" in out
