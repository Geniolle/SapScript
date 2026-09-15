# Análise Técnica: Split Contábil de Faturas com Mais de 999 Linhas (`FI_BILL_ISSUE_SPLIT`)

Esta análise documenta o mecanismo do SAP standard e as customizações ativas para contornar o limite de **999 linhas contábeis em FI** (erro **`F5 727`**), com foco na BAdI **`FI_BILL_ISSUE_SPLIT`** e na classe **`ZCLFI_BILL_ISSUE_SPLIT`**.

---

## 1. O Problema: Limite de 999 Linhas no SAP FI (`F5 727`)

* **Causa Raiz**: O campo `BUZEI` (número da linha de lançamento contábil na tabela `BSEG`) possui definição de dicionário como `NUMC 3` (apenas 3 dígitos numéricos: `001` a `999`).
* **Ponto de Disparo**: Localizado no include standard **[`abap/LFACIFSI.abap`](./abap/LFACIFSI.abap)** / **`LFACIFPP`**, rotina `PROJECT_TO_V_TABLES`, linhas 1057 a 1066:
  ```abap
  COUNT = COUNT + 1.
  ...
  XVBSEG-BUZEI = COUNT.
  IF XVBSEG-BUZEI = 0.
    MESSAGE E727 WITH 999.   " <--- Erro F5 727: Número máximo de posições em FI atingido (999)
  ENDIF.
  ```
* **Cenário**: Quando uma fatura de SD (ex.: operações de alta volumetria, faturas de transferência intercompany `ZIG2`, vendas com centenas de itens de material) gera mais de 999 linhas de receitas, impostos e clientes, o contador estoura o campo de 3 dígitos ao atingir o item 1.000, voltando para `000` e bloqueando a contabilização.

---

## 2. A Solução Standard: Rotina de Split (`SPLIT_INVOICE`)

No fluxo standard de integração SD $\rightarrow$ FI:
1. O faturamento de SD invoca a função `FI_DOCUMENT_PROJECT` (include `LFACIU04`).
2. **Antes** de tentar numerar e gravar as linhas em `PROJECT_TO_V_TABLES`, o SAP executa o particionamento condicional:
   ```abap
   PERFORM SPLIT_INVOICE.        " Divide o documento contábil em fatias de até 990 itens
   PERFORM PROJECT_TO_V_TABLES.  " Numera as linhas por fatia ($$1, $$2, etc.)
   ```
3. O split cria documentos intermediários (`$$1`, `$$2`, etc.), equilibrando o balanço de débito e crédito de cada documento filho através de uma conta técnica de compensação configurada na transação `OBYC` / tabela `T030` (operação **`SPL`** / conta ex.: `0027800800`).

---

## 3. BAdI `FI_BILL_ISSUE_SPLIT` e Implementação `ZCLFI_BILL_ISSUE_SPLIT`

O SAP delega a ativação e os parâmetros do split de faturamento SD à BAdI **`FI_BILL_ISSUE_SPLIT`** (interface `IF_EX_FI_BILL_ISSUE_SPLIT`).

No sistema, foi implementada a classe customizada **`ZCLFI_BILL_ISSUE_SPLIT`** ([código completo aqui](./abap/ZCLFI_BILL_ISSUE_SPLIT.abap)):

### Métodos Implementados:

1. **`ACTIVATE_AUTOMATIC_SPLIT`**:
   * Define `E_AUTOMATIC_SPLIT = 'X'`.
   * Ativa a rotina automática de checagem de volume no include `LFACIFSI`.

2. **`SET_NUMBER_OF_INVOICE_ITEMS`** *(Ponto de Atenção Identificado)*:
   ```abap
   METHOD IF_EX_FI_BILL_ISSUE_SPLIT~SET_NUMBER_OF_INVOICE_ITEMS.
   *    e_number_of_invoice_items = 900.
   ENDMETHOD.
   ```
   * **Observação**: A atribuição explícita de `900` itens por fatia está **comentada** no código ativo do SAP PRD. Sem esse valor explícito, a decisão de particionar fica restrita à contagem calculada internamente em `DET_SPLIT_DOC_STRUCTURE`.

3. **`SET_DOCUMENT_TYPE_SUBSEQ`**:
   * Define `E_DOCUMENT_TYPE_SUBSEQ = 'ZV'`.
   * Indica que o primeiro documento contábil é criado com o tipo original (ex.: `RH`) e os documentos subsequentes derivados da quebra são criados com o tipo de documento contábil **`ZV`**.

---

## 4. Tabelas SAP Envolvidas

| Tabela | Função no Processo de Split |
| :--- | :--- |
| **`T003`** | Cadastro de tipos de documento (`RH` e `ZV`). Ambos precisam de parametrização harmônica para permitir o split (ex.: intervalo de numeração interno, flag líquido, etc.). |
| **`T8G12`** | Classificação de tipos de documento no New G/L Document Split (ambos configurados com processo `0000` / variante `0001`). |
| **`T030`** | Determinação de contas de razão. Operação `SPL` define a conta transitória de compensação do split. |
| **`NRIV`** | Intervalos de numeração contábil de cada tipo de documento. |
| **`VBFA`** | Tabela de fluxo de documentos SD, onde todos os documentos FI gerados pelo split ficam amarrados à mesma fatura `VBRK`. |

---

## 5. Caso Real Investigado: Fatura `2541036794`

* **Tipo de Documento**: `ZIG2` (Faturamento Intercompany).
* **Volume**: **1.944 itens de receita (`ZPI0`)** na tabela `VBRP`.
* **Comportamento Observado**:
  * Ao tentar liberar para contabilidade sem a correta quebra, o SAP emitiu o erro `F5 727`.
  * A reconstituição das estruturas `ACCIT` no include `LV60BF00` confirmou que todas as posições possuíam `POSNR_SD` atribuído corretamente (`XACCIT-POSNR_SD = XVBRP-POSNR`).
  * A investigação revelou que quando `e_number_of_invoice_items` está comentado e não há segmentação prévia, a rotina `SPLIT_INVOICE` pode dar `EXIT` prematuro se os critérios de agrupamento de imposto ou cliente não atenderem aos filtros do include `LFACIFSI`.

---

## 6. Scripts de Diagnóstico

Dentro da pasta [`scripts/`](./scripts/) estão disponíveis ferramentas de diagnóstico direto via RFC:

1. **Diagnóstico da BAdI e Parâmetros**:
   ```bash
   python scripts/diagnostico_split_badi.py
   ```
   Valida o código ativo da classe `ZCLFI_BILL_ISSUE_SPLIT`, as propriedades de `RH`/`ZV` na `T003` e a conta de compensação `SPL` na `T030`.

2. **Diagnóstico de Volume e Fluxo de Fatura**:
   ```bash
   python scripts/diagnostico_itens_fatura.py --fatura 2541036794
   ```
   Lê os dados da fatura, conta os itens na `VBRP` e valida no fluxo `VBFA` se foram gerados múltiplos documentos contábeis ou se o fluxo falhou.

---

## 7. Arquivos ABAP Preservados

* [`abap/ZCLFI_BILL_ISSUE_SPLIT.abap`](./abap/ZCLFI_BILL_ISSUE_SPLIT.abap): Código consolidado da classe de implementação da BAdI.
* [`abap/LFACIFSI.abap`](./abap/LFACIFSI.abap): Código standard completo da rotina `SPLIT_INVOICE` (Função `SAPLFACI`).
* [`abap/LV60BF00.abap`](./abap/LV60BF00.abap): Interface de faturamento SD $\rightarrow$ Contabilidade (`ACCOUNTING_ITEM_LINE`).
* [`abap/RV60AFZZ.abap`](./abap/RV60AFZZ.abap): Includes de User Exit da faturação SD.
* [`abap/RV60C901.abap`](./abap/RV60C901.abap): Rotina SD de faturação.
