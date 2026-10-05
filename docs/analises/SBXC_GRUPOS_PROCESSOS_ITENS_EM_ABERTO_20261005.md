# Análise Técnica: Agrupamento por Processos dos 122 Itens em Aberto (/SBXC/)

**Data:** 05/10/2026  
**Ambiente:** SAP PRD  
**Origem dos Dados:** `C:\workspace\SapScript\output\SBXC_tables_PRD_20260929_194959\manifest.json`

---

## 1. Contexto & Resumo Executivo

Durante a varredura do namespace `/SBXC/` para extração de dados e tabelas no SAP PRD, identificaram-se **122 objetos com erro de leitura/extração via `RFC_READ_TABLE`**.

O diagnóstico técnico revela a divisão exata:
* **5 Tabelas Transparentes (`TRANSP`):** Tabelas de banco de dados físicas contendo registros, que falharam por necessidade de parâmetro de ordenação (`GET_SORTED`).
* **117 Estruturas Internas (`INTTAB`):** Tipos de dados, estruturas de memória e definições de BAPIs/Interfaces ABAP que **não possuem dados armazenados** em banco de dados (`DA 131 - Table/Structure is not transparent`).

---

## 2. Agrupamento Funcional e Técnico dos 122 Itens

### 2.1. Tabelas Transparentes (Requerem Parâmetro `GET_SORTED`) — 5 Itens
| Tabela | Descrição Funcional / Processo | Causa Técnica | Ação Recomendada |
| :--- | :--- | :--- | :--- |
| `/SBXC/ENVIA_GM` | Envio de Movimentação de Mercadorias (MIGO) | Exige `GET_SORTED` para `ROWSKIPS` | Ajustar script RFC com `GET_SORTED=X` |
| `/SBXC/ENVIA_PED` | Envio de Pedidos de Compra (MM-PO) | Exige `GET_SORTED` para `ROWSKIPS` | Ajustar script RFC com `GET_SORTED=X` |
| `/SBXC/ZCKP_CTRL` | Tabela de Controle do Cockpit SAP | Exige `GET_SORTED` para `ROWSKIPS` | Ajustar script RFC com `GET_SORTED=X` |
| `/SBXC/ZCKP_EKKN` | Imputação de Pedidos de Compra Cockpit | Exige `GET_SORTED` para `ROWSKIPS` | Ajustar script RFC com `GET_SORTED=X` |
| `/SBXC/ZCKP_HINVH` | Cabeçalho de Faturas / Inventário Cockpit | Exige `GET_SORTED` para `ROWSKIPS` | Ajustar script RFC com `GET_SORTED=X` |

---

### 2.2. Estruturas ABAP Internas (`INTTAB`) — 117 Itens

#### Group 1: Serviços de Documentos (`CNDDOCUMENT`) — 48 Itens
* **Escopo:** Estruturas de comunicação para integração de documentos eletrónicos e serviços de faturação.
* **Objetos:**
  * `/SBXC/CNDDOCUMENT_DATA`, `/SBXC/CNDDOCUMENT_INFO`, `/SBXC/CNDDOCUMENT_SEARCH`, `/SBXC/CNDDOCUMENT_STATS`
  * `/SBXC/ICNDDOCUMENT_SERVICE_G10` até `/SBXC/ICNDDOCUMENT_SERVICE_G37` (28 itens)
  * `/SBXC/ICNDDOCUMENT_SERVICE_GE1` até `/SBXC/ICNDDOCUMENT_SERVICE_GE9`, `/SBXC/ICNDDOCUMENT_SERVICE_GET` (10 itens)
  * `/SBXC/ICNDDOCUMENT_SERVICE_MA1`, `/SBXC/ICNDDOCUMENT_SERVICE_MAR` (2 itens)
  * `/SBXC/ICNDDOCUMENT_SERVICE_SE1` até `/SBXC/ICNDDOCUMENT_SERVICE_SE3`, `/SBXC/ICNDDOCUMENT_SERVICE_SEA` (4 itens)

#### Group 2: Tipos de Dados Internos Cockpit (`CKP_ST_*`) — 39 Itens
* **Escopo:** Estruturas de suporte a relatórios ALV e processamento em lote do Cockpit.
* **Objetos:**
  * `/SBXC/CKP_ST_001` a `/SBXC/CKP_ST_036` (36 estruturas)
  * `/SBXC/CKP_ST_DOCINFO_STRING`, `/SBXC/CKP_ST_DOCINFO_STRING2`
  * `/SBXC/CKP_ST_ITEM_STRING`, `/SBXC/CKP_ST_ITEM_STRING2`

#### Group 3: Processamento Cockpit & Faturamento (`ZCKP_*`) — 7 Itens
* **Escopo:** Estruturas funcionais de adiantamentos, bloqueio e lançamentos de despesas.
* **Objetos:**
  * `/SBXC/ZCKP_ADIANTAMENTO` (Adiantamentos a Fornecedores)
  * `/SBXC/ZCKP_ALV` (Formatação de Grelha ALV)
  * `/SBXC/ZCKP_BASE64` (Codificação de Anexos/Arquivos)
  * `/SBXC/ZCKP_BLOQ_FATURA` (Regras de Bloqueio de Faturas)
  * `/SBXC/ZCKP_COLOR` (Estilo e Status Visuais)
  * `/SBXC/ZCKP_DESPESASHEADER` & `/SBXC/ZCKP_DESPESASITEM` (Lançamento e Rateio de Despesas)

#### Group 4: Integradores & Utilitários — 7 Itens
* **Escopo:** Módulos de interface com portais externos (ex.: Saphety) e simulações financeiras.
* **Objetos:**
  * `/SBXC/BASE_ENTITY`
  * `/SBXC/CKP_MAIL_ATTACH` (Envio de Email com Anexos)
  * `/SBXC/DOCUMENT_CONTENT` & `/SBXC/DOCUMENT_FORMAT_TYPE`
  * `/SBXC/SEND_DOCINFO_2_SAPHETY` (Interface Saphety)
  * `/SBXC/SIMULA_ACCIT` (Simulação de Documento Contábil FI/CO)
  * `/SBXC/ST_IMG`

#### Group 5: BAPIs de Pedidos de Compra (MM-PO) — 5 Itens
* **Escopo:** Estruturas de entrada/saída das BAPIs de Pedidos de Compras.
* **Objetos:**
  * `/SBXC/ZBAPIMEPOACCOUNT`
  * `/SBXC/ZBAPIMEPOHEADER`
  * `/SBXC/ZBAPIMEPOITEM`
  * `/SBXC/ZBAPIMEPOSCHEDUL`
  * `/SBXC/ZBAPIPOCHANGE`

#### Group 6: Arrays & Coleções ABAP — 5 Itens
* **Escopo:** Tipos de tabela ABAP para passagem de dados complexos em memória.
* **Objetos:**
  * `/SBXC/ARRAY_OFINT`
  * `/SBXC/ARRAY_OF_CNDDOCUMENT_INF`
  * `/SBXC/ARRAY_OF_DOCUMENT_FORMAT`
  * `/SBXC/LAZY_OFBASE64BINARY`
  * `/SBXC/TUPLE_OF_ARRAY_OF_CNDDOC`

#### Group 7: BAPIs de Fornecedores (MM-Vendor) — 4 Itens
* **Escopo:** Estruturas de manutenção de cadastro de parceiros de negócios/fornecedores.
* **Objetos:**
  * `/SBXC/ZBAPIVENDOR_BANK`
  * `/SBXC/ZBAPIVENDOR_COMP`
  * `/SBXC/ZBAPIVENDOR_HEAD`
  * `/SBXC/ZBAPIVENDOR_PURCH_ORG`

#### Group 8: BAPIs de Movimentação de Mercadorias (MM-IM / MIGO) — 2 Itens
* **Escopo:** Estruturas de entrada para lançamento de movimentos de estoque.
* **Objetos:**
  * `/SBXC/ZBAPI_GM_HEAD`
  * `/SBXC/ZBAPI_GM_ITEM`

---

## 3. Plano de Ação & Próximos Passos

1. **Extração das 5 Tabelas Transparentes:**
   * Executar o script de leitura RFC ativando a ordenação primária (`GET_SORTED`) para contornar a limitação de paginação `ROWSKIPS`.
2. **Ignorar Tentativa de Leitura de Dados das 117 Estruturas:**
   * Como são estruturas `INTTAB` sem registros persistidos, a documentação de metadados (campos e tipos) deve ser obtida via DDIC (`DD03L` / `DD02L`), evitando requisições de tabela de dados.
