# -*- coding: utf-8 -*-
"""Processo: análise dos métodos de pagamento configurados para Portugal.

Objetivo atual:
- identificar o método SEPA de fornecedores utilizado na F110;
- identificar os formatos gerais e alternativos configurados;
- seguir tecnicamente o formato efetivamente utilizado pela empresa/banco;
- analisar o formato PMW Z_CGI_CT;
- confirmar o formato atual pain.001.001.03;
- investigar as estruturas internas da árvore DMEEX;
- preparar a extração da árvore Z_CGI_CT, seus nós e mappings.

IMPORTANTE:
- processo somente de leitura;
- não altera configuração SAP;
- parâmetros funcionais ficam concentrados neste ficheiro;
- lógica SAP/RFC permanece no engine genérico.
"""

# =============================================================================
# PARÂMETROS DO PROCESSO
# =============================================================================

METODO = "RFC"  # "RFC" ou "GUI"
SAP_KEY = "S4DCLNT100"
TRANSACTION = "SE16H"

ABRIR_NOVO_MODO = True
FECHAR_MODO_NO_FIM = False

MAX_ROWS = 200

GERAR_JSON = True
GERAR_CSV = False


# =============================================================================
# PARÂMETROS FUNCIONAIS
# =============================================================================

PAIS = "PT"

METODO_PAGAMENTO_ALVO = "S"

EMPRESA_ALVO = "2100"

BANCO_EMPRESA_ALVO = "BPI01"


# -----------------------------------------------------------------------------
# Formatos encontrados
# -----------------------------------------------------------------------------

# Formato geral definido no método S.
FORMATO_GERAL = "Z_SEPA_AP"

# Formato alternativo efetivamente utilizado por:
# 2100 + S + BPI01
FORMATO_ATUAL = "Z_CGI_CT"


# -----------------------------------------------------------------------------
# Objetos SAP já identificados
# -----------------------------------------------------------------------------

TABELA_METODOS_PAGAMENTO = "T042Z"

CAMPO_FORMATO_PAGAMENTO = "FORMI"

DATA_ELEMENT_FORMATO = "FORMI_COMBINED"

CHECKTABLE_FORMATO = "VPAYFRMTTEXTCOMB"

TABELA_ASSOCIACAO_FORMATO = "T042ZA_FORMAT"

TABELA_FORMATOS_PMW = "TFPM042F"


# -----------------------------------------------------------------------------
# Objetos DMEEX que vamos validar no DDIC
# -----------------------------------------------------------------------------

TABELA_DMEEX_NODES = "DMEE_TREE_NODE"

TABELA_DMEEX_HEADER = "DMEE_TREE_HEAD"


# =============================================================================
# CONSULTAS
# =============================================================================

CONSULTAS = [

    # =========================================================================
    # 01 - MÉTODOS DE PAGAMENTO DE PORTUGAL
    # =========================================================================
    {
        "nome": "Métodos de pagamento por país - Portugal",
        "tabela": TABELA_METODOS_PAGAMENTO,
        "filtros": [
            {
                "campo": "LAND1",
                "valor": PAIS,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 02 - MÉTODO S = SEPA-FORNECEDOR
    # =========================================================================
    {
        "nome": "Método SEPA de fornecedores - Portugal",
        "tabela": TABELA_METODOS_PAGAMENTO,
        "filtros": [
            {
                "campo": "LAND1",
                "valor": PAIS,
                "opcao": "EQ",
            },
            {
                "campo": "ZLSCH",
                "valor": METODO_PAGAMENTO_ALVO,
                "opcao": "EQ",
            },
        ],
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


    # =========================================================================
    # 03 - DDIC DE T042Z-FORMI
    # =========================================================================
    {
        "nome": "DDIC - definição técnica do campo T042Z-FORMI",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_METODOS_PAGAMENTO,
                "opcao": "EQ",
            },
            {
                "campo": "FIELDNAME",
                "valor": CAMPO_FORMATO_PAGAMENTO,
                "opcao": "EQ",
            },
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


    # =========================================================================
    # 04 - UTILIZAÇÕES DE FORMI_COMBINED
    # =========================================================================
    {
        "nome": "DDIC - onde o elemento FORMI_COMBINED é utilizado",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "ROLLNAME",
                "valor": DATA_ELEMENT_FORMATO,
                "opcao": "EQ",
            },
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
        ],
    },


    # =========================================================================
    # 05 - ESTRUTURA DA VPAYFRMTTEXTCOMB
    # =========================================================================
    {
        "nome": "DDIC - estrutura da check table/view VPAYFRMTTEXTCOMB",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": CHECKTABLE_FORMATO,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TABNAME",
            "FIELDNAME",
            "POSITION",
            "KEYFLAG",
            "ROLLNAME",
            "DOMNAME",
            "CHECKTABLE",
            "DATATYPE",
            "LENG",
        ],
    },


    # =========================================================================
    # 06 - TIPO TÉCNICO DE VPAYFRMTTEXTCOMB
    # =========================================================================
    {
        "nome": "DDIC - tipo técnico de VPAYFRMTTEXTCOMB",
        "tabela": "DD02L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": CHECKTABLE_FORMATO,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TABNAME",
            "TABCLASS",
            "SQLTAB",
            "CONTFLAG",
            "AS4LOCAL",
            "AS4VERS",
        ],
    },


    # =========================================================================
    # 07 - DESCRIÇÃO DO FORMATO GERAL
    # =========================================================================
    {
        "nome": "Formato geral Z_SEPA_AP - descrição no SAP",
        "tabela": CHECKTABLE_FORMATO,
        "filtros": [
            {
                "campo": "FORMI",
                "valor": FORMATO_GERAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "MANDT",
            "FORMI",
            "FORMX",
        ],
    },


    # =========================================================================
    # 08 - TABELAS QUE COMPÕEM VPAYFRMTTEXTCOMB
    # =========================================================================
    {
        "nome": "DDIC - tabelas que compõem VPAYFRMTTEXTCOMB",
        "tabela": "DD26S",
        "filtros": [
            {
                "campo": "VIEWNAME",
                "valor": CHECKTABLE_FORMATO,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 09 - ESTRUTURA DE T042ZA_FORMAT
    # =========================================================================
    {
        "nome": "DDIC - estrutura da tabela T042ZA_FORMAT",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_ASSOCIACAO_FORMATO,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 10 - FORMATO ALTERNATIVO EXATO
    # =========================================================================
    {
        "nome": "Formato alternativo - empresa 2100 método S banco BPI01",
        "tabela": TABELA_ASSOCIACAO_FORMATO,
        "filtros": [
            {
                "campo": "ZBUKR",
                "valor": EMPRESA_ALVO,
                "opcao": "EQ",
            },
            {
                "campo": "ZLSCH",
                "valor": METODO_PAGAMENTO_ALVO,
                "opcao": "EQ",
            },
            {
                "campo": "HBKID",
                "valor": BANCO_EMPRESA_ALVO,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "MANDT",
            "ZBUKR",
            "ZLSCH",
            "HBKID",
            "FORMI",
            "FORMZ",
            "DTTYP_ALTV",
            "HALGO",
        ],
    },


    # =========================================================================
    # 11 - ESTRUTURA DA TFPM042F
    # =========================================================================
    {
        "nome": "DDIC - estrutura da tabela PMW TFPM042F",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_FORMATOS_PMW,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 12 - DEFINIÇÃO PMW DE Z_CGI_CT
    # =========================================================================
    {
        "nome": "PMW - definição completa do formato Z_CGI_CT",
        "tabela": TABELA_FORMATOS_PMW,
        "filtros": [
            {
                "campo": "FORMI",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 13 - DEFINIÇÃO PMW DE Z_SEPA_AP
    # =========================================================================
    {
        "nome": "PMW - definição completa do formato geral Z_SEPA_AP",
        "tabela": TABELA_FORMATOS_PMW,
        "filtros": [
            {
                "campo": "FORMI",
                "valor": FORMATO_GERAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 14 - DESCRIÇÃO DE Z_CGI_CT
    # =========================================================================
    {
        "nome": "Formato utilizado Z_CGI_CT - descrição no SAP",
        "tabela": CHECKTABLE_FORMATO,
        "filtros": [
            {
                "campo": "FORMI",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "MANDT",
            "FORMI",
            "FORMX",
        ],
    },


    # =========================================================================
    # 15 - DDIC DA ESTRUTURA DE NÓS DMEEX
    # =========================================================================
    #
    # Não consultamos ainda os dados da tabela.
    #
    # Primeiro precisamos saber exatamente:
    # - campos de identificação da árvore;
    # - versão;
    # - ID do nó;
    # - nó pai;
    # - nome do nó;
    # - tipo do nó;
    # - posição/hierarquia.
    #
    # =========================================================================
    {
        "nome": "DMEEX - DDIC da estrutura DMEE_TREE_NODE",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_DMEEX_NODES,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 16 - TIPO TÉCNICO DE DMEE_TREE_NODE
    # =========================================================================
    {
        "nome": "DMEEX - tipo técnico de DMEE_TREE_NODE",
        "tabela": "DD02L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_DMEEX_NODES,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TABNAME",
            "TABCLASS",
            "SQLTAB",
            "CONTFLAG",
            "AS4LOCAL",
            "AS4VERS",
        ],
    },


    # =========================================================================
    # 17 - DDIC DO CABEÇALHO DA ÁRVORE DMEEX
    # =========================================================================
    #
    # Queremos descobrir no próprio sistema os campos utilizados para:
    # - tipo da árvore;
    # - ID da árvore;
    # - versão;
    # - status;
    # - descrição;
    # - informações de ativação.
    #
    # =========================================================================
    {
        "nome": "DMEEX - DDIC da estrutura DMEE_TREE_HEAD",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_DMEEX_HEADER,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 18 - TIPO TÉCNICO DE DMEE_TREE_HEAD
    # =========================================================================
    {
        "nome": "DMEEX - tipo técnico de DMEE_TREE_HEAD",
        "tabela": "DD02L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": TABELA_DMEEX_HEADER,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TABNAME",
            "TABCLASS",
            "SQLTAB",
            "CONTFLAG",
            "AS4LOCAL",
            "AS4VERS",
        ],
    },

    # =========================================================================
    # 19 - CABEÇALHO REAL DA ÁRVORE PAYM / Z_CGI_CT
    # =========================================================================
    #
    # Primeiro identificamos todas as versões existentes da árvore.
    #
    # Ainda não vamos fixar VERSION, porque queremos que o próprio SAP
    # nos diga quais versões existem e qual delas está atualmente guardada.
    #
    # =========================================================================
    {
        "nome": "DMEEX - cabeçalho da árvore PAYM Z_CGI_CT",
        "tabela": "DMEE_TREE_HEAD",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "PARAM_STRUC",
            "FIRSTNODE_ID",
            "VERSION_TYPE",
            "VERS_USER",
            "VERS_DATE",
            "VERS_TIME",
            "MSGTYPE",
            "VERSION_DESCRIPTION",
        ],
    },


    # =========================================================================
    # 20 - ESTRUTURA REAL DOS NÓS DA ÁRVORE PAYM / Z_CGI_CT
    # =========================================================================
    #
    # Nesta execução trazemos todas as versões existentes.
    #
    # Depois de identificar no cabeçalho a versão relevante,
    # podemos restringir a consulta.
    #
    # Os campos abaixo permitem reconstruir:
    #
    # - hierarquia pai/filho;
    # - ordem dos irmãos;
    # - nível;
    # - nome técnico XML;
    # - tipo de nó;
    # - mapping direto de estrutura/campo SAP;
    # - constantes;
    # - exits;
    # - regras de conversão.
    #
    # =========================================================================
    {
        "nome": "DMEEX - nós da árvore PAYM Z_CGI_CT",
        "tabela": "DMEE_TREE_NODE",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "NODE_ID",
            "TECH_NAME",
            "REF_NAME",
            "PARENT_ID",
            "BROTHER_ID",
            "FIRSTCHILD_ID",
            "NODE_TYPE",
            "DATA_TYPE",
            "EX_STATUS",
            "LEV",
            "LENGTH",
            "MP_SC_TAB",
            "MP_SC_FLD",
            "MP_SC_NODE",
            "MP_SC_REF_NAME",
            "MP_CONST",
            "MP_EXIT_FUNC",
            "MP_SELECTION",
            "CV_RULE",
        ],
    },


    # =========================================================================
    # 21 - NÓS QUE POSSUEM MAPPING POR EXIT
    # =========================================================================
    #
    # Consulta auxiliar.
    #
    # Aqui não conseguimos usar "diferente de vazio" de forma portátil
    # sem primeiro validar como o engine trata NE/NOT INITIAL.
    #
    # Portanto mantemos a árvore completa na consulta 20 e depois
    # analisamos no JSON quais nós possuem MP_EXIT_FUNC preenchido.
    #
    # Esta consulta deixa apenas os principais campos de mapping,
    # facilitando a análise visual.
    #
    # =========================================================================
    {
        "nome": "DMEEX - visão de mappings da árvore PAYM Z_CGI_CT",
        "tabela": "DMEE_TREE_NODE",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "VERSION",
            "NODE_ID",
            "TECH_NAME",
            "NODE_TYPE",
            "LEV",
            "MP_SC_TAB",
            "MP_SC_FLD",
            "MP_CONST",
            "MP_EXIT_FUNC",
            "CV_RULE",
        ],
    },

    # =========================================================================
    # 22 - DDIC DA TABELA PRINCIPAL DMEE_TREE
    # =========================================================================
    #
    # Objetivo:
    # descobrir como o SAP guarda:
    # - árvore;
    # - versão ativa;
    # - status;
    # - outras referências da árvore.
    #
    # =========================================================================
    {
        "nome": "DMEEX - DDIC da tabela principal DMEE_TREE",
        "tabela": "DD03L",
        "filtros": [
            {
                "campo": "TABNAME",
                "valor": "DMEE_TREE",
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 23 - REGISTO PRINCIPAL DE PAYM / Z_CGI_CT
    # =========================================================================
    #
    # Não restringimos campos ainda.
    # Queremos ver tudo o que DMEE_TREE guarda para esta árvore.
    #
    # =========================================================================
    {
        "nome": "DMEEX - definição principal da árvore PAYM Z_CGI_CT",
        "tabela": "DMEE_TREE",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
        ],
        "campos_saida": [],
    },


    # =========================================================================
    # 24 - NÓS COMPLETOS DA VERSÃO 000
    # =========================================================================
    {
        "nome": "DMEEX - nós completos Z_CGI_CT versão 000",
        "tabela": "DMEE_TREE_NODE",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
            {
                "campo": "VERSION",
                "valor": "000",
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "NODE_ID",
            "TECH_NAME",
            "REF_NAME",
            "PARENT_ID",
            "BROTHER_ID",
            "FIRSTCHILD_ID",
            "NODE_TYPE",
            "DATA_TYPE",
            "EX_STATUS",
            "LEV",
            "LENGTH",
            "MP_SC_TAB",
            "MP_SC_FLD",
            "MP_SC_NODE",
            "MP_SC_REF_NAME",
            "MP_CONST",
            "MP_EXIT_FUNC",
            "MP_SELECTION",
            "CV_RULE",
        ],
    },


    # =========================================================================
    # 25 - NÓS COMPLETOS DA VERSÃO 001
    # =========================================================================
    {
        "nome": "DMEEX - nós completos Z_CGI_CT versão 001",
        "tabela": "DMEE_TREE_NODE",
        "filtros": [
            {
                "campo": "TREE_TYPE",
                "valor": "PAYM",
                "opcao": "EQ",
            },
            {
                "campo": "TREE_ID",
                "valor": FORMATO_ATUAL,
                "opcao": "EQ",
            },
            {
                "campo": "VERSION",
                "valor": "001",
                "opcao": "EQ",
            },
        ],
        "campos_saida": [
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "NODE_ID",
            "TECH_NAME",
            "REF_NAME",
            "PARENT_ID",
            "BROTHER_ID",
            "FIRSTCHILD_ID",
            "NODE_TYPE",
            "DATA_TYPE",
            "EX_STATUS",
            "LEV",
            "LENGTH",
            "MP_SC_TAB",
            "MP_SC_FLD",
            "MP_SC_NODE",
            "MP_SC_REF_NAME",
            "MP_CONST",
            "MP_EXIT_FUNC",
            "MP_SELECTION",
            "CV_RULE",
        ],
    },
]

# =============================================================================
# CONTRATO CONSUMIDO PELO RUNNER / ENGINE
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


# =============================================================================
# EXECUÇÃO DIRETA
# =============================================================================

if __name__ == "__main__":

    import sys
    from pathlib import Path

    _base_dir = Path(__file__).resolve().parent.parent

    if str(_base_dir) not in sys.path:
        sys.path.insert(
            0,
            str(_base_dir),
        )

    from engine import executar_processo

    sys.exit(
        executar_processo(PROCESSO)
    )
