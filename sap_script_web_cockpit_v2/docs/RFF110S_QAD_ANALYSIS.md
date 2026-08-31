# Análise do RFF110S em QAD

Este documento regista o que foi observado ao ler o programa standard
`RFF110S` diretamente no QAD via RFC.

## Como foi lido

- Conexão RFC ao QAD com `pyrfc`
- Leitura via `RPY_PROGRAM_READ`
- O source principal veio sem linhas úteis
- A lógica funcional ficou concentrada sobretudo em:
  - `RFF110S_FORMS`
  - `RFF110S_SELSCR_BOE`

## Campos relevantes observados no standard SAP

### Campos de proposta

- `PAR_XVL`
  - controla a execução da proposta
  - quando está em `X`, o programa trata o fluxo como proposal-only
- `PAR_BUDA`
  - posting date do documento
- `PAR_GRDA`
  - date limit para os open items
- `PAR_NEDA`
  - posting date do próximo payment run
- `SEL_BUKR`
  - company code(s)
- `PAR_ZWE`
  - métodos de pagamento aceites
- `SEL_KRED`
  - contas de fornecedor
- `SEL_DEBI`
  - contas de cliente

### Validação importante

O standard rejeita a execução quando:

- `SEL_KRED` está vazio
- e `SEL_DEBI` está vazio

## O que isto significa para o cockpit

O nosso fluxo automático não pode assumir que o `document_number` por si só
seleciona a proposta no SAP.

O `RFF110S` seleciona por:

- empresa
- conta
- método de pagamento
- datas de seleção

Depois disso, no cockpit, usamos o `document_number` criado como referência
para confirmar se o item apareceu na proposta.

## Correções aplicadas no código do cockpit

- Garantimos que `environment` segue no payload da proposta.
- Garantimos que `payment_method` está presente para cliente e fornecedor.
- Garantimos que `company_code`, `account_number` e `posting_date` são
  enviados para a proposta.
- Adicionámos logging na camada RFC para registar:
  - request preparado
  - open-item check
  - seleção enviada ao `RFF110S`
  - resultado final da proposta

## Nota sobre o documento criado

No nosso caso, o `document_number` do FI é o identificador operacional do
workflow.

Ele é usado para:

- mostrar o resultado do FI ao utilizador
- verificar se o item entrou na proposta
- diagnosticar quando a proposta vem vazia

## Probe local com documento de fornecedor

Teste executado no QAD com:

- `environment`: `QAD`
- `operation_type`: `pagamento`
- `company_code`: `2010`
- `payment_method`: `S`
- `account_number`: conta de fornecedor configurada no ambiente
- `posting_date`: `2026-08-27`
- `next_due_date`: assumido como `posting_date` quando vazio
- `document_number`: `6050000047`

Resultado observado:

- o documento foi encontrado em aberto
- o item não entrou na proposta
- a proposta terminou com zero itens

Isto confirma que a camada Python já consegue montar e submeter a proposta,
mas a seleção standard do SAP ainda não está a incluir esse documento nos
critérios nativos do `RFF110S`.

Mas não substitui os critérios nativos de seleção do `RFF110S`.
