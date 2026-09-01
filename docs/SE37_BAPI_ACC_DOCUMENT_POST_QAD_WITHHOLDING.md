# SE37 case - `BAPI_ACC_DOCUMENT_POST` - QAD - Fornecedor - retenção

Este ficheiro resume o caso de teste usado localmente para reproduzir a chamada da `BAPI_ACC_DOCUMENT_POST` com retenção de fornecedor em QAD.

## Objetivo

Gerar um documento FI de fornecedor com retenção ativa, de forma a validar:

- `ACCOUNTPAYABLE`
- `ACCOUNTGL`
- `ACCOUNTWT`
- `CURRENCYAMOUNT`
- posterior leitura em `WITH_ITEM`

## Contexto observado no SAP QAD

- `BUKRS`: `2010`
- `LIFNR`: `0010000040`
- `WITHT`: `P5`
- `WT_WITHCD`: `63`
- `T059Z` para `PT / P5 / 63` retorna taxa `25.0000`

## Dados de entrada

Use os seguintes valores no teste:

- `environment`: `QAD`
- `branch`: `fornecedor`
- `company_code`: `2010`
- `posting_date`: `2026-09-01`
- `document_date`: `2026-09-01`
- `vendor_account`: `0010000040`
- `expense_gl_account`: valor ativo de `SAP_FI_EXPENSE_GL_ACCOUNT` no ambiente local
- `amount`: `100.00`
- `currency`: `EUR`
- `withholding_tax_type`: `P5`
- `withholding_tax_code`: `63`
- `withholding_tax_base_amount`: `100.00`
- `withholding_tax_amount`: `25.00`

## Estrutura do payload

### DOCUMENTHEADER

```text
COMP_CODE   = 2010
DOC_TYPE    = KR
DOC_DATE    = 2026-09-01
PSTNG_DATE  = 2026-09-01
USERNAME    = utilizador SAP atual
HEADER_TXT  = teste fornecedor com retenção
```

### ACCOUNTPAYABLE

```text
ITEMNO_ACC  = 1
VENDOR_NO   = 0010000040
ITEM_TEXT   = teste fornecedor com retenção
W_TAX_CODE  = 63
```

### ACCOUNTGL

```text
ITEMNO_ACC  = 2
GL_ACCOUNT  = valor ativo de SAP_FI_EXPENSE_GL_ACCOUNT
ITEM_TEXT   = teste fornecedor com retenção
```

### ACCOUNTWT

```text
ITEMNO_ACC   = 1
WT_TYPE      = P5
WT_CODE      = 63
BAS_AMT_LC   = 100.00
BAS_AMT_TC   = 100.00
BAS_AMT_IND  = X
AWH_AMT_LC   = 25.00
AWH_AMT_TC   = 25.00
```

### CURRENCYAMOUNT

```text
ITEM 1
  ITEMNO_ACC = 1
  CURR_TYPE  = 00
  CURRENCY   = EUR
  AMT_DOCCUR = -100.00

ITEM 2
  ITEMNO_ACC = 2
  CURR_TYPE  = 00
  CURRENCY   = EUR
  AMT_DOCCUR = +100.00
```

## Variante sugerida no SE37

Se quiser gravar um conjunto de dados no teste da função:

- nome sugerido: `Z_QAD_FI_WITHHOLDING_P5_63`
- função: `BAPI_ACC_DOCUMENT_POST`

## Observações importantes

- Este caso não inclui IVA `ACCOUNTTAX`.
- A retenção é tratada por `ACCOUNTWT`.
- O objetivo é confirmar se o SAP grava linhas em `WITH_ITEM`.
- Se `WITH_ITEM` continuar vazio, o problema já não está no payload básico do BAPI, mas no customizing ou na lógica standard aplicada no ambiente.

## Execução local usada como referência

O mesmo cenário foi testado localmente pelo script:

- `tests/manual/run_qad_withholding_tax_document.py`

