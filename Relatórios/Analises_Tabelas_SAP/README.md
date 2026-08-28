# Análises de Tabelas SAP

Consultas read-only de tabelas SAP via SAP GUI Scripting e SE16H/SE16N.

## Visão Geral

Não alterar ou apagar uma análise existente quando surgir uma necessidade nova. Cada assunto fica isolado num ficheiro próprio dentro de `processos/`, enquanto toda a lógica de login, navegação, filtros, leitura do ALV e exportação fica centralizada em `engine.py`.

## Estrutura

```text
Relatórios/
└── Analises_Tabelas_SAP/
    ├── engine.py
    ├── runner.py
    ├── README.md
    └── processos/
        ├── _template.py
        └── metodos_pagamento_pt.py
```

### `engine.py`

Motor genérico. Não deve receber parâmetros de negócio específicos.

Responsabilidades:

- reutilizar/fazer login no SAP via `sap_session.py`;
- abrir SE16H/SE16N;
- selecionar tabela;
- aplicar filtros pelos nomes técnicos dos campos;
- executar a consulta;
- ler o ALV;
- imprimir os resultados no terminal;
- gravar JSON e, opcionalmente, CSV em `cache/analises_tabelas_sap/<processo>/`.

### `runner.py`

Executa um processo pelo nome do ficheiro.

Exemplo:

```text
.venv\Scripts\python.exe "Relatórios\Analises_Tabelas_SAP\runner.py" metodos_pagamento_pt
```

Listar processos:

```text
.venv\Scripts\python.exe "Relatórios\Analises_Tabelas_SAP\runner.py" --listar
```

### `processos/`

Cada ficheiro representa uma análise independente. É aqui que ficam os valores variáveis:

- sistema SAP;
- transação de leitura;
- tabelas;
- filtros;
- campos a devolver;
- quantidade máxima de linhas;
- geração de JSON/CSV.

## Criar uma Nova Análise

1. Copiar `processos/_template.py`.
2. Renomear, por exemplo, para `okb9.py`, `sepa_ct_v9.py` ou `centros_custo.py`.
3. Alterar somente a secção `PARÂMETROS DO PROCESSO`.
4. Definir um `PROCESSO["id"]` único.
5. Executar pelo `runner.py`.

Exemplo de uma consulta:

```python
CONSULTAS = [
    {
        "nome": "Empresa 2100",
        "tabela": "T001",
        "filtros": [
            {"campo": "BUKRS", "valor": "2100", "opcao": "EQ"},
        ],
        "campos_saida": ["BUKRS", "BUTXT", "LAND1", "WAERS"],
    },
]
```

Um mesmo processo pode consultar várias tabelas. Basta acrescentar novos blocos na lista `CONSULTAS`.

## Regras de Uso

- O motor é somente de leitura.
- Não guardar utilizadores, passwords ou dados sensíveis nos ficheiros de processo.
- As credenciais e o destino SAP continuam centralizados no `.env` e em `sap_session.py`.
- Não duplicar lógica de SAP GUI dentro dos processos.
- Quando uma análise nova tiver tabelas/filtros diferentes, criar um novo ficheiro em `processos/` em vez de substituir um processo anterior.
