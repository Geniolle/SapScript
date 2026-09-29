"""
Testes para o módulo VIES com retry automático.
Usa mocks para simular diferentes cenários sem depender do serviço real.
"""

import sys
from pathlib import Path
from unittest.mock import patch, MagicMock
from vies_service import (
    validar_vies_com_retry,
    eh_erro_temporario,
    ViesResult,
    ViesResultStatus,
)

# Adicionar diretório ao path
sys.path.insert(0, str(Path(__file__).parent))


def teste_classificacao_erros_temporarios():
    """Testa a classificação de erros temporários vs permanentes."""
    print("=" * 60)
    print("TESTE: Classificação de Erros")
    print("=" * 60)

    # Erros temporários
    temporarios = [
        "MS_MAX_CONCURRENT_REQ",
        "timeout",
        "SERVER_BUSY",
        "SERVICE_UNAVAILABLE",
        "temporarily unavailable",
        "Connection refused",
    ]

    for erro in temporarios:
        resultado = eh_erro_temporario(erro)
        status = "✅" if resultado else "❌"
        print(f"{status} '{erro}' → temporário: {resultado}")
        assert resultado, f"Deveria ser temporário: {erro}"

    # Erros permanentes (exemplos)
    permanentes = [
        "O VAT não pode estar vazio.",
        "VAT inválido.",
        "Resposta inesperada do serviço VIES.",
    ]

    for erro in permanentes:
        resultado = eh_erro_temporario(erro)
        status = "✅" if not resultado else "❌"
        print(f"{status} '{erro}' → temporário: {resultado}")
        assert not resultado, f"Deveria ser permanente: {erro}"

    print()


def teste_retry_sucesso_na_primeira():
    """Testa retry quando a primeira tentativa tem sucesso."""
    print("=" * 60)
    print("TESTE: Retry - Sucesso na Primeira Tentativa")
    print("=" * 60)

    resultado_mock = ViesResult(
        vat="PT504772694",
        country_code="PT",
        vat_number="504772694",
        valid=True,
        name="Empresa Test",
        address="Rua Test",
        request_date="2026-09-29",
    )

    with patch("vies_service.validar_vies") as mock_vies:
        mock_vies.return_value = resultado_mock

        status = validar_vies_com_retry("PT504772694")

        assert status.is_success
        assert status.result.valid
        assert status.attempt == 1
        assert mock_vies.call_count == 1  # Sem retry

        print("✅ Sucesso na primeira tentativa - sem retry")
        print(f"   Tentativa: {status.attempt}/{status.total_attempts}")
        print(f"   Resultado: VÁLIDO")
        print()


def teste_retry_sucesso_na_segunda():
    """Testa retry com sucesso na segunda tentativa."""
    print("=" * 60)
    print("TESTE: Retry - Sucesso na Segunda Tentativa")
    print("=" * 60)

    resultado_ok = ViesResult(
        vat="FR70434317293",
        country_code="FR",
        vat_number="70434317293",
        valid=True,
        name="Test France",
        address=None,
        request_date="2026-09-29",
    )

    with patch("vies_service.validar_vies") as mock_vies:
        # Primeira chamada falha temporariamente, segunda sucede
        mock_vies.side_effect = [
            RuntimeError("MS_MAX_CONCURRENT_REQ"),
            resultado_ok,
        ]

        status = validar_vies_com_retry("FR70434317293")

        assert status.is_success
        assert status.result.valid
        assert status.attempt == 2
        assert mock_vies.call_count == 2

        print("✅ Sucesso na segunda tentativa após erro temporário")
        print(f"   Tentativas usadas: {status.attempt}")
        print(f"   Resultado: VÁLIDO")
        print()


def teste_retry_falha_permanente():
    """Testa que não faz retry em erros permanentes."""
    print("=" * 60)
    print("TESTE: Retry - Sem Retry em Erro Permanente")
    print("=" * 60)

    with patch("vies_service.validar_vies") as mock_vies:
        mock_vies.side_effect = ValueError("VAT inválido.")

        try:
            status = validar_vies_com_retry("INVALID")
            # Não deve chegar aqui
            assert False, "Deveria ter lançado exceção"
        except ValueError:
            print("✅ Erro permanente detectado - sem retry")
            print("   Tentativas: 1 (sem retry)")
            print()


def teste_retry_esgota_tentativas():
    """Testa comportamento quando se esgotam as tentativas."""
    print("=" * 60)
    print("TESTE: Retry - Esgota 4 Tentativas com Erro Temporário")
    print("=" * 60)

    with patch("vies_service.validar_vies") as mock_vies:
        # Todas as tentativas falham com erro temporário
        mock_vies.side_effect = RuntimeError("SERVER_BUSY")

        status = validar_vies_com_retry("FR70434317293")

        assert not status.is_success
        assert status.is_temporary_error
        assert status.attempt == 4
        assert mock_vies.call_count == 4

        print("✅ Esgotadas as 4 tentativas")
        print(f"   Tentativas: {status.attempt}/{status.total_attempts}")
        print(f"   Resultado: ERRO TÉCNICO")
        print()


def teste_retry_resultado_invalido():
    """Testa que VAT inválido não sofre retry."""
    print("=" * 60)
    print("TESTE: Retry - VAT Inválido (valid=false) Sem Retry")
    print("=" * 60)

    resultado_invalido = ViesResult(
        vat="FR00000000000",
        country_code="FR",
        vat_number="00000000000",
        valid=False,
        name=None,
        address=None,
        request_date="2026-09-29",
    )

    with patch("vies_service.validar_vies") as mock_vies:
        mock_vies.return_value = resultado_invalido

        status = validar_vies_com_retry("FR00000000000")

        assert status.is_success
        assert not status.result.valid
        assert status.attempt == 1
        assert mock_vies.call_count == 1  # Sem retry

        print("✅ VAT inválido retornou imediatamente")
        print(f"   Tentativa: {status.attempt}/{status.total_attempts}")
        print(f"   Resultado: NÃO VÁLIDO (sem retry)")
        print()


def teste_callback_status():
    """Testa callback de status durante retry."""
    print("=" * 60)
    print("TESTE: Callback de Status During Retry")
    print("=" * 60)

    resultado_ok = ViesResult(
        vat="PT504772694",
        country_code="PT",
        vat_number="504772694",
        valid=True,
        name="Test",
        address=None,
        request_date="2026-09-29",
    )

    callbacks_recebidos = []

    def on_status(status: ViesResultStatus):
        callbacks_recebidos.append(status)

    with patch("vies_service.validar_vies") as mock_vies:
        mock_vies.side_effect = [
            RuntimeError("timeout"),
            resultado_ok,
        ]

        status = validar_vies_com_retry(
            "PT504772694",
            on_retry_status=on_status
        )

        # Verificar que temos callbacks - pelo menos 2 (erro + sucesso)
        assert len(callbacks_recebidos) >= 2, \
            f"Esperava 2+ callbacks, obtive {len(callbacks_recebidos)}"
        print(f"✅ {len(callbacks_recebidos)} callbacks recebidos")
        for i, cb in enumerate(callbacks_recebidos, 1):
            print(f"   {i}. Tentativa {cb.attempt} - sucesso: {cb.is_success}")
        print()


def main():
    """Executa todos os testes."""
    print()
    print("╔" + "=" * 58 + "╗")
    print("║" + "TESTES DE RETRY DO VIES".center(58) + "║")
    print("╚" + "=" * 58 + "╝")
    print()

    try:
        teste_classificacao_erros_temporarios()
        teste_retry_sucesso_na_primeira()
        teste_retry_sucesso_na_segunda()
        teste_retry_falha_permanente()
        teste_retry_esgota_tentativas()
        teste_retry_resultado_invalido()
        teste_callback_status()

        print("=" * 60)
        print("TODOS OS TESTES DE RETRY PASSARAM ✅")
        print("=" * 60)
        print()
        return 0

    except AssertionError as exc:
        print()
        print("=" * 60)
        print("TESTE FALHOU ❌")
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
