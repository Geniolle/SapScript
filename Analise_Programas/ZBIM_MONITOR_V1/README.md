# Análise Técnica: Monitor de Revisão de Faturas (`ZBIM_MONITOR_V1`)

Esta documentação detalha a arquitetura, modelo de dados, fluxos de negócio e código-fonte do programa **`ZBIM_MONITOR_V1`** (Transação **`ZBIM_MONITOR`**), responsável pela gestão, monitorização e resolução de faturas bloqueadas por divergência (preço, quantidade e bloqueio contábil) com integração ao SAP Business Workflow.

---

## 1. Visão Geral do Programa

* **Transação Principal**: `ZBIM_MONITOR` (Ecrã `1000`)
* **Programa ABAP**: `ZBIM_MONITOR_V1`
* **Título**: *Monitor de faturas bloqueadas* / *Monitor de revisão de faturas*
* **Pacote de Desenvolvimento**: `ZWF` (Workflows Customizados)
* **Autor Original**: `CFERNANDES` (Criado em 09.08.2021)
* **Tipo de Programa**: Programa Executável (Report com Dynpro/ALV Grid)
* **Objetivo de Negócio**:
  Centralizar num único cockpit a gestão de faturas retidas em conferência de entrada (MM - revisão de faturas `MIRO`/`MR8M`) e pagamentos bloqueados em financeiro (FI - `FB60`/`BSEG`), integrando diretamente o acompanhamento do workflow de divergências, despacho de tarefas de aprovação, conferência contra pedidos/receções e geração de documentos de regularização (notas de débito/crédito, estornos ou ajustes de inventário/receção).

---

## 2. Estrutura dos Fontes ABAP

O programa está estruturado de forma modular em includes padronizados, localizados na pasta [`abap/`](./abap/):

```text
Analise_Programas/ZBIM_MONITOR_V1/
├── README.md                      # Esta documentação técnica
├── abap/
│   ├── ZBIM_MONITOR_V1.abap       # Report principal e ponto de entrada
│   ├── ZBIM_MONITOR_V1_TOP.abap   # Declarações globais, tipos e instâncias ALV
│   ├── ZBIM_MONITOR_V1_SCR.abap   # Ecrã de seleção e filtros (Dynpro 1000)
│   ├── ZBIM_MONITOR_V1_PBO.abap   # Módulo PBO e configuração do ALV Grid
│   ├── ZBIM_MONITOR_V1_PAI.abap   # Módulo PAI e tratamento de comandos
│   ├── ZBIM_MONITOR_V1_LCL.abap   # Classe local de eventos (hotspots, navegação, dispatch)
│   └── ZBIM_MONITOR_V1_FRMS.abap  # Sub-rotinas (regras de negócio, queries, reconciliação)
└── scripts/
    └── consultar_zbim_monitor.py  # Script de diagnóstico e consulta read-only via RFC
```

### Relação dos Includes

| Include | Linhas | Responsabilidade Técnica |
| :--- | :--- | :--- |
| [`ZBIM_MONITOR_V1.abap`](./abap/ZBIM_MONITOR_V1.abap) | 24 | Controla o ciclo de vida inicial: validação de bloqueio administrativo (`f_check_adm_block`), reconciliação prévia de status (`f_update_before_run`), extração (`f_select_data`) e chamada da tela ALV (`f_display_alv`). |
| [`ZBIM_MONITOR_V1_TOP.abap`](./abap/ZBIM_MONITOR_V1_TOP.abap) | 36 | Define estruturas de saída com tabela de estilo (`ty_outtab_log`, `ty_outtab_fi` contendo `celltab TYPE lvc_t_styl`), tabelas internas globais e instâncias de controle GUI (`CL_GUI_CUSTOM_CONTAINER`, `CL_GUI_ALV_GRID`). |
| [`ZBIM_MONITOR_V1_SCR.abap`](./abap/ZBIM_MONITOR_V1_SCR.abap) | 46 | Define os 6 blocos do ecrã de seleção para seleção de modo (MM vs. FI), filtros de fatura/pedido/fornecedor, motivos de bloqueio, workflow e documentos derivados. |
| [`ZBIM_MONITOR_V1_PBO.abap`](./abap/ZBIM_MONITOR_V1_PBO.abap) | 73 | Inicializa o ALV Grid no container `C_GRID`, define layout zebra/ajuste de colunas, associa a classe de eventos e liga o modo de edição para células configuradas via `CELLTAB`. |
| [`ZBIM_MONITOR_V1_PAI.abap`](./abap/ZBIM_MONITOR_V1_PAI.abap) | 35 | Processa saídas (`SAIR`), navegação de retorno com verificação de alterações pendentes (`VOLTAR`), gravação manual (`SAVE`) e refresh estável do ALV (`g_grid->refresh_table_display`). |
| [`ZBIM_MONITOR_V1_LCL.abap`](./abap/ZBIM_MONITOR_V1_LCL.abap) | 205 | Implementa `lcl_alv->handle_hotspot_click`: navegação para transações standard (`MIR4`, `FB03`, `ME23N`, `MIGO_DIALOG`), despacho do workflow via `SWL_WI_DISPATCH` e consulta a logs de aplicação. |
| [`ZBIM_MONITOR_V1_FRMS.abap`](./abap/ZBIM_MONITOR_V1_FRMS.abap) | 1036 | Concentra toda a lógica de reconciliação de dados, conferência de histórico de pedidos (`ME_READ_HISTORY`), verificação automática de encerramento de divergências e persistência. |

---

## 3. Modos de Operação e Ecrã de Seleção

O ecrã de seleção organiza a pesquisa em dois mundos distintos selecionados via botões de opção:

### 1. Documentos Logísticos (`P_LOG` - Padrão)
* Atua sobre faturas bloqueadas oriundas do módulo de Compras / Gestão de Materiais (MM-IV / Verificação de Faturas).
* Baseia-se na tabela customizada **`ZBIM_BLK_INVOICE`** e tabelas standard **`RBKP`** (cabeçalho) e **`RSEG`** (itens).
* Avalia divergências detalhadas por posição:
  * Quantidade faturada maior que receção (`SPGRM`).
  * Preço faturado divergente do pedido (`SPGRP`).
  * Bloqueio por lançamento em conta de razão (`ZBLCK_RAZAO`).

### 2. Documentos Financeiros (`P_FI`)
* Atua sobre lançamentos de fornecedores em FI (`FB60` / `FB01`).
* Baseia-se na tabela customizada **`ZBIM_BLK_INV_FI`** e tabelas standard **`BKPF`** e **`BSEG`**.
* Controla lançamentos bloqueados para pagamento através da chave de bloqueio **`BSEG-ZLSPR`**.

### Blocos de Filtro do Ecrã de Seleção:
* **Bloco 1 (`b01`) - Dados empresa e modo**: Empresa (`S_BUKRS`), seleção `P_LOG` ou `P_FI`.
* **Bloco 2 (`b02`) - Dados do documento**: Fatura (`S_BELNR`), Exercício (`S_GJAHR`), Pedido de Compra (`S_EBELN`), Fornecedor (`S_LIFNR`).
* **Bloco 3 (`b03`) - Motivo de bloqueio**: Bloqueio de Preço (`P_SPGRP`), Bloqueio de Quantidade (`P_SPGRM`), Bloqueio Razão (`P_RAZAO`).
* **Bloco 4 (`b04`) - Dados de criação e workflow**: Data do Workitem (`S_WI_CD`), Status do Workitem (`S_STAT`), ID do Workitem (`S_WIID`), Utilizador responsável (`S_WIUSER`).
* **Bloco 5 (`b05`) - Documentos gerados**:
  * Ajuste de Inventário (`S_DOCINV` / `S_INVDAT`)
  * Ajuste de Entrada de Mercadorias (`S_DOCEM` / `S_EMDAT`)
  * Nota de Débito / Crédito (`S_DOC_NC` / `S_NCDATE`)
  * Status global (`S_STATUS`), Data fim (`S_ZEOP`), Doc. Estorno (`S_DOCEST`), Flag Nota Débito (`S_NC_ND`).
* **Bloco 6 (`b06`) - Filtros especiais de caixa de entrada do utilizador**:
  * `P_MYWIDT`: Workitems por iniciar do utilizador logado (`WI_STAT = 'PRONTO'`).
  * `P_MYWID`: Workitems em curso sob responsabilidade do utilizador logado.
  * `P_ALL`: Todos os workitems do utilizador logado (em qualquer status).
  * `P_ALLWI`: Visão irrestrita (todos os workitems de todos os utilizadores).

---

## 4. Fluxo Funcional e Regras de Negócio

### 4.1. Bloqueio Administrativo do Monitor (`f_check_adm_block`)
Antes de qualquer execução, o monitor consulta a tabela de parâmetros **`ZBIM_PARAM_T`**:
```abap
SELECT SINGLE * FROM zbim_param_t INTO @DATA(ls_param_t)
  WHERE zprocess = 'BIM' AND fname = 'ADM_BLOCK'.
```
Se `LOW = 'X'`, é emitido um pop-up `POPUP_TO_INFORM` informando *"Currently the monitor is blocked for maintenance purposes"* e a execução é interrompida.

### 4.2. Reconciliação Automática Pré-Execução (`f_update_before_run`)
Esta é uma das rotinas mais críticas do programa:
1. Lê todos os registos em aberto (`ZSTATUS <> 'CLOSED'`).
2. Cruza com as tabelas de itens de fatura `RSEG` e partidas de fornecedor `BSEG`:
   * **Divergência de Quantidade (`SPGRM`)**: Se o motivo do bloqueio já foi desfeito na `RSEG` (ex.: receção de mercadoria concluída no almoxarifado), o monitor define `ZSTATUS = 'CLOSED'`, `ZDIV_CLOSED = 'X'`, atualiza a data `ZEOP_DATE = sy-datum` e invoca a função **`ZBIM_CLOSE_WORKITEM`** para fechar automaticamente o workitem de workflow pendente associado.
   * **Divergência de Preço (`SPGRP`)**: Se a discrepância de preço foi removida na `RSEG` (ex.: nota de crédito de acerto de preço processada ou autorização sem bloqueio), invoca **`ZBIM_CLOSE_WORKITEM`** com o evento de encerramento.
   * **Documentos FI (`P_FI`)**: Se o bloqueio de pagamento `BSEG-ZLSPR` foi removido manualmente ou por outro processo, invoca **`ZBIM_RAISE_EVENT`** (`NO_DIV`) e finaliza a posição.
3. Atualiza os dados de aprovação/agente do workitem chamando a função **`ZBIM_GET_WIID`**. Caso haja múltiplos agentes possíveis, lê a tabela standard de workflow **`SWWUSERWI`** e concatena todos os utilizadores separados por barra vertical (ex.: `USER1|USER2`).

### 4.3. Cálculo e Comparação no Ecrã ALV (`f_select_data`)
Para cada posição logística exibida:
1. Executa a função standard **`ME_READ_HISTORY`** com `WEBRE = 'X'` para obter o histórico completo de entradas de mercadorias (`EKBE`/`EKBES`).
2. Calcula a divergência real:
   * **Diferença de Quantidade**:
     $$\text{Diferença EM vs FAT} = \text{EKBES-BPMNG} - \text{ZTOT\_REMNG}$$
   * **Diferença de Valor**:
     $$\text{Diferença Valor} = \text{EKBES-REEWR} - \text{EKBES-WEWRT}$$
     (Ou valor unitário da fatura multiplicado pela quantidade nos casos de faturas de ativos).
3. Converte a decisão do utilizador (`ZUSER_DECISION`) a partir do domínio standard **`Y_USER_DECISION_V`**:
   * `1` $\rightarrow$ **ACCEPT** / **ACEITA DIV**
   * `2` $\rightarrow$ **REJECT** / **RECUSA DIV**
4. Se a fatura tiver sido desbloqueada em `BSEG-ZLSPR = ' '` e não for um caso com nota de débito pendente (`ZDEBIT_NOTE <> 'SNC'`), marca o status global como `CLOSED`.

### 4.4. Interatividade do Utilizador (ALV Hotspots e Ações)
A classe local `lcl_alv` responde aos cliques nas colunas do relatório:

| Coluna | Ação Realizada |
| :--- | :--- |
| **`BELNR`** | Salta para visualização da fatura: `MIR4` (no modo MM) ou `FB03` (no modo FI), pulando o primeiro ecrã. |
| **`EBELN`** | Salta para visualização do Pedido de Compras na transação `ME23N`. |
| **`ZMBLNR_INV`** | Abre o visualizador de documento de material (`MIGO_DIALOG`, ação `A04`) para o documento de ajuste de inventário. |
| **`ZMBLNR_EM`** | Abre `MIGO_DIALOG` para o documento de material da entrada de mercadorias. |
| **`ZBELNR_NC`** | Salta para visualização da Nota de Débito/Crédito gerada via transação `MIR4`. |
| **`ZEXECUTE`** | **Despacho direto de Workflow**: grava `USER_TRATA = sy-uname`, invoca a função **`SWL_WI_DISPATCH`** com função `'APRO'` para que o utilizador execute a sua etapa de decisão, espera a conclusão e atualiza os status no ALV. |
| **`ICON_LOG_NC`** | Abre o Application Log (`BAL_DSP_LOG_DISPLAY`) referente à geração da Nota de Débito/Crédito. |
| **`ICON_LOG_MM`** | Abre o Application Log referente a movimentos de materiais de ajuste. |
| **`ICON_LOG_DESB`** | Abre o Application Log referente ao desbloqueio da fatura no sistema. |

### 4.5. Campo Editável de Documento Interno (`ZINTERN_DOC`)
O relatório permite edição direta em linha da coluna **`ZINTERN_DOC`** (Documento Interno de suporte).
Ao clicar no botão de salvar (`SAVE`) ou ao retornar (`VOLTAR`), a sub-rotina `f_save_data` / `f_update_table` valida alterações via `g_grid->check_changed_data` e grava os novos valores diretamente em `ZBIM_BLK_INVOICE` ou `ZBIM_BLK_INV_FI`.

---

## 5. Dicionário de Dados e Tabelas Relacionadas

### 5.1. Tabela `ZBIM_BLK_INVOICE` (Faturas Logísticas / MM)

| Campo | Tipo | Descrição |
| :--- | :--- | :--- |
| **`MANDT`** | CLNT(3) | Mandante (Chave) |
| **`BUKRS`** | CHAR(4) | Empresa (Chave) |
| **`BELNR`** | CHAR(10) | Número do documento de fatura contábil / MIRO (Chave) |
| **`GJAHR`** | NUMC(4) | Exercício (Chave) |
| **`BUZEI`** | NUMC(6) | Item de documento na fatura (Chave) |
| **`ZLSPR`** | CHAR(1) | Chave de bloqueio de pagamento (`BSEG-ZLSPR`) |
| **`ZUSER_DECISION`** | CHAR(20) | Decisão do utilizador (1=Aceita, 2=Recusa) |
| **`ZREASON`** | CHAR(48) | Justificação da decisão de divergência |
| **`EBELN`** / **`EBELP`** | CHAR(10) / NUMC(5) | Pedido de compras e posição do pedido |
| **`LIFNR`** / **`NAME`** | CHAR(10) / CHAR(35) | Código e razão social do fornecedor |
| **`XBLNR`** | CHAR(16) | Número de referência da fatura externa do fornecedor |
| **`SPGRP`** | CHAR(1) | Flag de motivo de bloqueio: Preço |
| **`SPGRM`** | CHAR(1) | Flag de motivo de bloqueio: Quantidade |
| **`ZBLCK_RAZAO`** | CHAR(1) | Flag de bloqueio por lançamento em conta de razão |
| **`MENGE`** / **`MEINS`** | QUAN(13) / UNIT(3) | Quantidade do Pedido de Compras e Unidade de Medida |
| **`WEMNG`** | QUAN(13) | Quantidade recebida (Entrada de Mercadorias) |
| **`REMNG`** | QUAN(13) | Quantidade faturada |
| **`ZTOT_REMNG`** | QUAN(13) | Quantidade total faturada |
| **`ZDIF_EM_FAT`** | QUAN(13) | Diferença entre quantidade recebida e faturada |
| **`NETPR`** / **`ZNETPR_UNIT`** | CURR(11) | Valor líquido total e unitário do Pedido |
| **`WRBTR`** / **`ZWRBTR_UNIT`** | CURR(13) | Valor bruto total e unitário da fatura |
| **`ZWRBTR_DIF`** | CURR(13) | Diferença de valor entre receção e fatura |
| **`ZDEBIT_NOTE`** | CHAR(3) | Indicador de criação de Nota de Débito (`ND` / `SNC`) |
| **`ZINTERN_DOC`** | CHAR(15) | Referência de Documento Interno (editável no monitor) |
| **`WI_ID`** | NUMC(12) | ID do Work Item SAP Business Workflow |
| **`WI_CRUSER`** | CHAR(150) | Responsável(eis) pelo Work Item atual |
| **`WI_STAT`** | CHAR(12) | Status do processamento do Work Item |
| **`ZMBLNR_INV`** / **`ZMJAHR_INV`** | CHAR(10) / NUMC(4) | Documento e ano do ajuste de inventário gerado |
| **`ZMBLNR_EM`** / **`ZMJAHR_EM`** | CHAR(10) / NUMC(4) | Documento e ano da receção de mercadorias de ajuste |
| **`ZBELNR_NC`** / **`ZGJAHR_NC`** | CHAR(10) / NUMC(4) | Documento e ano da Nota de Débito/Crédito gerada |
| **`ZSTATUS`** | CHAR(10) | Status global do processo (`CLOSED` quando concluído) |
| **`ZEOP_DATE`** | DATS(8) | Data de encerramento do processo (End Of Process Date) |
| **`ZDOC_ESTORNO`** | CHAR(10) | Documento contábil de estorno |
| **`BALOGNR`** / **`BALOGNR_MM`** / **`BALOGNR_DESB`** | CHAR(20) | Identificadores de Application Log (NC, MM e Desbloqueio) |
| **`USER_TRATA`** | CHAR(12) | Utilizador que iniciou o tratamento do processo |

### 5.2. Tabela `ZBIM_BLK_INV_FI` (Faturas Financeiras / FI)
Armazena faturas do contas a pagar criadas diretamente em FI que possuem bloqueio de pagamento (`BSEG-ZLSPR` ativo). Possui estrutura simplificada com foco em cabeçalho, montante bruto (`GROSS_AMOUNT`), data do documento (`BLDAT`), data de lançamento (`PSTNG_DATE`), workitem e documento interno.

### 5.3. Parâmetros de Sistema (`ZBIM_PARAM_T`)
Controla o comportamento dos tipos de movimento e regras do processo:

| Parâmetro | Valor Configurado | Significado no Processo |
| :--- | :--- | :--- |
| **`BWART_RECECAO`** | `101` | Tipo de movimento para entrada regular de mercadorias |
| **`BWART_INVENT`** | `985` | Tipo de movimento customizado para regularização de inventário |
| **`BWART_CREDITNOTE`** | `800` | Tipo de movimento para geração de nota de crédito/débito |
| **`BWART_CREDITMAIL`** / **`BWART_CREDITNOTEMAIL`** | `801` | Tipo de movimento para notas enviadas por e-mail |
| **`BWART_RELEASE`** | `999` | Tipo de movimento para liberação |
| **`DOC_TYPE`** | `RD` | Tipo de documento para lançamento de notas |
| **`LGORT`** | `0001` | Depósito padrão para lançamentos de ajuste |
| **`WERKS_TRADE`** | `1003`, `1004` | Centros de mercadoria/trading |
| **`ADM_BLOCK`** | *(Vazio / Desativado)* | Flag de bloqueio geral do monitor para manutenção |

---

## 6. Arquitetura de Workflow e Integração

O ecossistema BIM é suportado por:
* **Classe de Workflow**: `ZCL_WF_BIM_REQ` (implementa os métodos de negócio: `ANALYSE_DECISION`, `CLOSE_WORKITEM`, `PROCESS_DIV_ACCEPTED`, `PROCESS_DIV_REJECTED`, `PROCESS_IM_ACTION`, etc.).
* **Objeto de Negócio**: Subtipo `ZRBUS2081` do objeto standard `BUS2081` (*Incoming Invoice*).
* **Eventos**:
  * `BLOCKED`: Disparado quando uma fatura é bloqueada na MIRO.
  * `BLOCKED_FI`: Disparado no bloqueio de fatura em FI.
  * `NO_DIV`: Disparado quando a divergência é sanada (fechamento automático).
  * `CANCELLED`: Disparado no estorno ou cancelamento do documento.
* **Módulos de Função Principais**:
  * `ZBIM_CLOSE_WORKITEM`: Encerramento de workitems via evento.
  * `ZBIM_GET_WIID`: Consulta o ID, status e agente atribuído.
  * `ZBIM_DIV_ACCEPTED` / `ZBIM_DIV_REJECTED`: Execução das decisões de aceitação/recusa.
  * `ZBIM_GOODSMVT_CREATE`: Geração de documentos de material para regularização.
  * `ZBIM_INCOMINGINVOICE_CREATE`: Criação de faturas subsequentes / notas.
  * `SWL_WI_DISPATCH`: Despacho interativo da tela de aprovação do workitem.

---

## 7. Script de Consulta e Diagnóstico

Para consultar faturas no monitor em modo estritamente de leitura via RFC sem abrir o SAP GUI:

```bash
# Consultar até 10 faturas de MM pendentes:
python Analise_Programas/ZBIM_MONITOR_V1/scripts/consultar_zbim_monitor.py --modo log --limite 10

# Consultar faturas financeiras (FI) filtradas por empresa:
python Analise_Programas/ZBIM_MONITOR_V1/scripts/consultar_zbim_monitor.py --modo fi --empresa 2010

# Consultar uma fatura específica:
python Analise_Programas/ZBIM_MONITOR_V1/scripts/consultar_zbim_monitor.py --modo log --fatura 5105927489 --exercicio 2026
```
