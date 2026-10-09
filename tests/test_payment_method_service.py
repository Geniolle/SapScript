"""
test_payment_method_service.py - Testes unitários para a sugestão e consulta de Meios de Pagamento (T042Z)
"""
from sap_rfc.payment_method_service import suggest_next_code


def test_suggest_next_code_basic():
    used = {"A", "B", "C", "1", "2"}
    next_code, available = suggest_next_code(used)
    assert next_code == "D"
    assert "D" in available
    assert "A" not in available


def test_suggest_next_code_when_z_used():
    used = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "S", "T", "U", "V", "W", "Z"}
    next_code, available = suggest_next_code(used)
    assert next_code == "R"
    assert "R" in available
