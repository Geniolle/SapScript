"""
Testes para o validador VIES
Valida a lógica de normalização, separação e validação de VAT.
"""

import sys
from pathlib import Path

# Adicionar diretório ao path para importações
sys.path.insert(0, str(Path(__file__).parent))

from vies_service import normalizar_vat, separar_vat


def teste_normalizar_vat():
    """Testa normalização de VAT."""
    print("=" * 60)
    print("TESTE: Normalizar VAT")
    print("=" * 60)

    casos = [
        ("PT504772694", "PT504772694"),
        ("PT 504 772 694", "PT504772694"),
        ("pt504772694", "PT504772694"),
        ("ES A28017895", "ESA28017895"),
        ("ESA 280 178 95", "ESA28017895"),
    ]

    for entrada, esperado in casos:
        try:
            resultado = normalizar_vat(entrada)
            status = "✓" if resultado == esperado else "✗"
            print(
                f"{status} {entrada:<20} => {resultado:<20} "
                f"(esperado: {esperado})"
            )
            assert resultado == esperado
        except Exception as exc:
            print(f"✗ {entrada:<20} => ERRO: {exc}")

    print()


def teste_separar_vat():
    """Testa separação de VAT em país e número."""
    print("=" * 60)
    print("TESTE: Separar VAT")
    print("=" * 60)

    casos = [
        ("PT504772694", ("PT", "504772694")),
        ("ESA28017895", ("ES", "A28017895")),
        ("FRXY123456789", ("FR", "XY123456789")),
    ]

    for entrada, (pais_esperado, num_esperado) in casos:
        try:
            pais, numero = separar_vat(entrada)
            status = "✓" if (pais == pais_esperado and numero == num_esperado) else "✗"
            print(
                f"{status} {entrada:<20} => "
                f"País: {pais:<3} Número: {numero:<15} "
                f"(esperado: {pais_esperado}, {num_esperado})"
            )
            assert pais == pais_esperado
            assert numero == num_esperado
        except Exception as exc:
            print(f"✗ {entrada:<20} => ERRO: {exc}")

    print()


def teste_separar_vat_invalidos():
    """Testa rejeição de VATs inválidos."""
    print("=" * 60)
    print("TESTE: Rejeitar VATs Inválidos")
    print("=" * 60)

    casos_invalidos = [
        "",
        "P",
        "123456",
        "PTPT123456",  # Sem número depois do país
    ]

    for entrada in casos_invalidos:
        try:
            resultado = separar_vat(entrada)
            print(f"✗ {entrada:<20} => Deveria ter falhado mas retornou: {resultado}")
        except ValueError as exc:
            print(f"✓ {entrada:<20} => Rejeitado corretamente: {exc}")

    print()


def main():
    """Executa todos os testes."""
    print()
    print("╔" + "=" * 58 + "╗")
    print("║" + "TESTES DO VALIDADOR VIES".center(58) + "║")
    print("╚" + "=" * 58 + "╝")
    print()

    try:
        teste_normalizar_vat()
        teste_separar_vat()
        teste_separar_vat_invalidos()

        print("=" * 60)
        print("TODOS OS TESTES PASSARAM ✓")
        print("=" * 60)
        print()
        return 0

    except AssertionError as exc:
        print()
        print("=" * 60)
        print("TESTE FALHOU ✗")
        print("=" * 60)
        print(str(exc))
        return 1

    except Exception as exc:
        print()
        print("=" * 60)
        print("ERRO NOS TESTES")
        print("=" * 60)
        print(str(exc))
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())
