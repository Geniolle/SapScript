# -*- coding: utf-8 -*-
"""Processo: análise dos métodos de pagamento configurados para Portugal.

ALTERAR SOMENTE A SECÇÃO DE PARÂMETROS quando esta análise evoluir.
A lógica SAP fica no engine genérico e não deve ser copiada para este ficheiro.
"""

# =============================================================================
# PARÂMETROS DO PROCESSO — ALTERAR AQUI NO VS CODE
# =============================================================================

SAP_KEY = "S4DCLNT100"
TRANSACTION = "SE16H"
ABRIR_NOVO_MODO = True
FECHAR_MODO_NO_FIM = False
MAX_ROWS = 200
GERAR_JSON = True
GERAR_CSV = False

PAIS = "PT"

# Cada consulta é independente. Para acrescentar outra tabela à mesma análise,
# basta adicionar outro bloco na lista CONSULTAS.
CONSULTAS = [
    {
        "nome": "Métodos de pagamento por país - Portugal",
        "tabela": "T042Z",
        "filtros": [
            {"campo": "LAND1", "valor": PAIS, "opcao": "EQ"},
        ],
        # [] = devolver todas as colunas disponíveis no ALV.
        "campos_saida": [],
    },
]

# =============================================================================
# NÃO ALTERAR ABAIXO — contrato consumido pelo runner/engine
# =============================================================================

PROCESSO = {
    "id": "metodos_pagamento_pt",
    "titulo": "Análise de configuração - Métodos de pagamento PT",
    "sap_key": SAP_KEY,
    "transaction": TRANSACTION,
    "abrir_novo_modo": ABRIR_NOVO_MODO,
    "fechar_modo_no_fim": FECHAR_MODO_NO_FIM,
    "max_rows": MAX_ROWS,
    "gerar_json": GERAR_JSON,
    "gerar_csv": GERAR_CSV,
    "consultas": CONSULTAS,
}

if __name__ == "__main__":
    import sys
    from pathlib import Path

    _base_dir = Path(__file__).resolve().parent.parent
    if str(_base_dir) not in sys.path:
        sys.path.insert(0, str(_base_dir))

    from engine import executar_processo

    sys.exit(executar_processo(PROCESSO))

