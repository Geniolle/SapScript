# -*- coding: utf-8 -*-
"""
sap_rfc/status_logger.py
======================================================================
Módulo Centralizador de Formatação e Cores de Status no Terminal.
Garante uniformidade visual em todas as etapas da Tarefa 1:
- SUCESSO: Linha completa em VERDE brilhante
- AVISO/PENDÊNCIA: Linha completa em AMARELO brilhante
- ERRO: Linha completa em VERMELHO brilhante
- INFO: Texto padrão (sem ANSI adicional)
======================================================================
"""

from __future__ import annotations

import sys

# Inicialização e suporte a cores multiplataforma (Windows / Linux)
try:
    import colorama
    if hasattr(colorama, "just_fix_windows_console"):
        colorama.just_fix_windows_console()
    colorama.init(strip=False)
    CLR_VERDE = colorama.Fore.GREEN + colorama.Style.BRIGHT
    CLR_AMARELO = colorama.Fore.YELLOW + colorama.Style.BRIGHT
    CLR_VERMELHO = colorama.Fore.RED + colorama.Style.BRIGHT
    CLR_AZUL = colorama.Fore.CYAN + colorama.Style.BRIGHT
    CLR_RESET = colorama.Style.RESET_ALL
except Exception:
    CLR_VERDE = "\033[92;1m"
    CLR_AMARELO = "\033[93;1m"
    CLR_VERMELHO = "\033[91;1m"
    CLR_AZUL = "\033[96;1m"
    CLR_RESET = "\033[0m"


def format_status_success(text: str) -> str:
    """Formata uma linha inteira de status de sucesso em verde."""
    return f"{CLR_VERDE}{text}{CLR_RESET}"


def format_status_warning(text: str) -> str:
    """Formata uma linha inteira de status de aviso/pendência em amarelo."""
    return f"{CLR_AMARELO}{text}{CLR_RESET}"


def format_status_error(text: str) -> str:
    """Formata uma linha inteira de status de erro em vermelho."""
    return f"{CLR_VERMELHO}{text}{CLR_RESET}"


def format_status_info(text: str) -> str:
    """Formata texto informativo (sem cor especial ou padrão)."""
    return text


def print_status_success(text: str) -> None:
    """Imprime uma linha de sucesso (em verde completo)."""
    print(format_status_success(text))


def print_status_warning(text: str) -> None:
    """Imprime uma linha de aviso/pendência (em amarelo completo)."""
    print(format_status_warning(text))


def print_status_error(text: str) -> None:
    """Imprime uma linha de erro (em vermelho completo)."""
    print(format_status_error(text))


def print_status_info(text: str) -> None:
    """Imprime uma linha de informação neutra."""
    print(format_status_info(text))
