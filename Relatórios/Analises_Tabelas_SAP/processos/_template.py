# -*- coding: utf-8 -*-
"""TEMPLATE — copiar este ficheiro e renomear para criar uma nova análise SAP."""

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

# Parâmetros de negócio próprios deste processo.
# Exemplos:
# EMPRESA = "2100"
# PAIS = "PT"
# METODO_PAGAMENTO = "T"

CONSULTAS = [
    {
        "nome": "Descrição da consulta",
        "tabela": "T001",
        "filtros": [
            {"campo": "BUKRS", "valor": "2100", "opcao": "EQ"},
        ],
        # [] = todas as colunas disponíveis no ALV.
        "campos_saida": ["BUKRS", "BUTXT", "LAND1", "WAERS"],
    },
    # Pode adicionar quantas tabelas forem necessárias ao mesmo processo:
    # {
    #     "nome": "Outra tabela relacionada",
    #     "tabela": "T042Z",
    #     "filtros": [
    #         {"campo": "LAND1", "valor": "PT", "opcao": "EQ"},
    #     ],
    #     "campos_saida": [],
    # },
]

# =============================================================================
# NÃO ALTERAR ABAIXO — contrato consumido pelo runner/engine
# =============================================================================

PROCESSO = {
    "id": "NOME_UNICO_DO_PROCESSO",
    "titulo": "Título legível da análise",
    "sap_key": SAP_KEY,
    "transaction": TRANSACTION,
    "abrir_novo_modo": ABRIR_NOVO_MODO,
    "fechar_modo_no_fim": FECHAR_MODO_NO_FIM,
    "max_rows": MAX_ROWS,
    "gerar_json": GERAR_JSON,
    "gerar_csv": GERAR_CSV,
    "consultas": CONSULTAS,
}
