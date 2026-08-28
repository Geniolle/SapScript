# F110 vs ZFI_EXTERNAL_PAYMENTS

Este documento compara o fluxo atual do cockpit para F110 com o padrão que um programa
customizado de pagamentos externos, como `ZFI_EXTERNAL_PAYMENTS`, deveria seguir.

## Contexto atual do repositório

- Não encontrei nenhuma implementação chamada `ZFI_EXTERNAL_PAYMENTS` neste workspace.
- A base atual do F110 está em `sap_rfc/f110_service.py`.
- O cockpit web encaminha a execução via worker Windows, com bridge RFC para SAP.

## Fluxo atual do F110 neste projeto

1. A API recebe o payload da proposta.
2. O worker normaliza `run_date` e `next_due_date`.
3. O executor RFC valida a presença do documento.
4. O executor abre sessão XMI.
5. O executor cria o job XBP.
6. O executor adiciona a etapa `RFF110S`.
7. O executor fecha e tenta iniciar o job.
8. O executor lê o status e a lista da proposta.

## Ponto fraco do fluxo atual

- O job pode ficar em `P` e não avançar para execução.
- Se a etapa não entrar corretamente, o SAP acusa `BT182` ou equivalentes de job incompleto.
- O fluxo depende de uma sequência rígida de chamadas RFC e de permissões XBP/XMI.

## O que um `ZFI_EXTERNAL_PAYMENTS` bem estruturado faria

1. Encapsularia toda a orquestração num único programa ABAP.
2. Receberia parâmetros de negócio já fechados.
3. Criaria e liberaria o job internamente, sem expor a montagem a múltiplas camadas do cockpit.
4. Registraria mensagens de erro e sucesso em log próprio.
5. Retornaria um identificador estável de execução e resultado.

## Comparação objetiva

| Tema | F110 atual no cockpit | ZFI_EXTERNAL_PAYMENTS ideal |
| --- | --- | --- |
| Responsabilidade | Orquestração distribuída entre API, worker e RFC | Orquestração centralizada no ABAP |
| Superfície de erro | Alta, por múltiplas camadas | Menor, porque o SAP controla o ciclo inteiro |
| Dependência de runtime | Windows + PyRFC + bridge | Principalmente SAP/ABAP |
| Rastreamento | Job, run id e logs do worker | Log ABAP e job único |
| Evolução | Mais sensível a mudanças em bridge/RFC | Mais estável para operação SAP pura |

## Checklist de implementação para aproximar o cockpit do padrão ABAP

- Garantir que o job nunca seja deixado incompleto.
- Garantir que a etapa do job exista antes de liberar.
- Garantir que o runtime do worker seja sempre Windows quando usar PyRFC.
- Garantir validação de documento antes de tentar propor.
- Garantir log de erro explícito quando SAP não aceitar a montagem do job.
- Isolar a sequência de execução para não depender de lógica espalhada na UI.

## Recomendação prática

Se existir um `ZFI_EXTERNAL_PAYMENTS` no SAP, o cockpit deve preferir chamá-lo como
um programa ABAP único e não reproduzir internamente toda a lógica de montagem de job.
Isso reduz risco de job incompleto, problema de release e divergência entre ambientes.

Se não existir esse programa no SAP, o fluxo atual precisa continuar centralizado e
monitorado no cockpit, mas sem deixar a definição do job em estado parcial.

## Checklist de validação operacional

- Confirmar qual ambiente está disponível antes de abrir a proposta.
- Se `QAD` estiver indisponível e a validação for apenas de leitura, usar a conexão produtiva.
- Não criar job novo em produtivo quando o objetivo for só validar parâmetros.
- Validar apenas a estrutura do payload, o mapeamento de conta e o `NEXT = hoje + 1`.
- Confirmar se o documento alvo aparece como encontrado antes de qualquer tentativa de execução.
- Registrar o resultado da validação em `.md` para evitar mudança de critério entre sessões.

## Regra para validação sem criação

Quando o pedido for apenas validar o formato ou as informações do payload:

- usar a ligação produtiva;
- executar apenas consultas ou verificações não destrutivas;
- não chamar fluxo de criação/liberação de job;
- não chamar fluxo que gere proposta ou movimentação em SAP.
