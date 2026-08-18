# -*- coding: utf-8 -*-
"""Processo: análise dos métodos de pagamento configurados para Portugal.

Objetivo atual:
- listar os métodos de pagamento configurados para Portugal;
- isolar o método SEPA de fornecedores utilizado na F110;
- identificar tecnicamente o campo T042Z-FORMI no DDIC, para seguir a
  configuração do formato de pagamento sem depender de suposições.

ALTERAR SOMENTE A SECÇÃO DE PARÂMETROS quando esta análise evoluir.
A lógica SAP fica no engine genérico e não deve ser copiada para este ficheiro.
"""

# =============================================================================
# PARÂMETROS DO PROCESSO — ALTERAR AQUI NO VS CODE
# =============================================================================

METODO = "RFC"  # "RFC" ou "GUI"
SAP_KEY = "S4DCLNT100"
TRANSACTION = "SE16H"
ABRIR_NOVO_MODO = True
FECHAR_MODO_NO_FIM = False
MAX_ROWS = 200
GERAR_JSON = True
GERAR_CSV = False

# -----------------------------------------------------------------------------
# Parâmetros funcionais da análise
# -----------------------------------------------------------------------------

PAIS = "PT"

# Método de pagamento identificado na análise anterior:
# S = SEPA-Fornecedor
METODO_PAGAMENTO_ALVO = "S"

# Campo da T042Z que atualmente contém Z_SEPA_AP para o método S.
TABELA_METODOS_PAGAMENTO = "T042Z"
CAMPO_FORMATO_PAGAMENTO = "FORMI"

# -----------------------------------------------------------------------------
# Consultas
# -----------------------------------------------------------------------------
# Cada consulta é independente.
#
# Para acrescentar outra tabela à mesma análise, adicione um novo bloco.
# Não altere a lógica do engine.
# -----------------------------------------------------------------------------

CONSULTAS = [
    {
        "nome": "Métodos de pagamento por país - Portugal",
        "tabela": TABELA_METODOS_PAGAMENTO,
        "filtros": [
            {"campo": "LAND1", "valor": PAIS, "opcao": "EQ"},
        ],
        # Mantemos todas as colunas para preservar a visão completa.
        "campos_saida": [],
    },
    {
        "nome": "Método SEPA de fornecedores - Portugal",
        "tabela": TABELA_METODOS_PAGAMENTO,
        "filtros": [
            {"campo": "LAND1", "valor": PAIS, "opcao": "EQ"},
            {"campo": "ZLSCH", "valor": METODO_PAGAMENTO_ALVO, "opcao": "EQ"},
        ],
        # Campos principais para a análise do método S.
        "campos_saida": [
            "MANDT",
            "LAND1",
            "ZLSCH",
            "TEXT1",
            "FORMI",
            "PROGN",
            "BLART",
            "BLARV",
            "XIBAN",
            "XSEPA",
        ],
    },
    {
        "nome": "DDIC - definição técnica do campo T042Z-FORMI",
        "tabela": "DD03L",
        "filtros": [
            {"campo": "TABNAME", "valor": TABELA_METODOS_PAGAMENTO, "opcao": "EQ"},
            {"campo": "FIELDNAME", "valor": CAMPO_FORMATO_PAGAMENTO, "opcao": "EQ"},
        ],
        "campos_saida": [
            "TABNAME",
            "FIELDNAME",
            "POSITION",
            "ROLLNAME",
            "DOMNAME",
            "CHECKTABLE",
            "DATATYPE",
            "LENG",
            "DECIMALS",
        ],
    },
]

# =============================================================================
# NÃO ALTERAR ABAIXO — contrato consumido pelo runner/engine
# =============================================================================

PROCESSO = {
    "id": "metodos_pagamento_pt",
    "titulo": "Análise de configuração - Métodos de pagamento PT",
    "metodo": METODO,
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
