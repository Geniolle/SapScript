# Fluxo FI Default Document

Este fluxo foi isolado para evitar regressões entre worker, API e RFC.

## Regra de arquitetura

- O worker executa o documento FI.
- A API guarda e expõe o job.
- O worker nunca atualiza o SQLite local diretamente para gravar o resultado FI.
- A gravação de `fi_document_result` deve sempre passar pela API.

## Arquivos principais

- `sap_script_web_cockpit_v2/worker/fi_default_document_job.py`
- `sap_script_web_cockpit_v2/worker/sap_tasks.py`
- `sap_script_web_cockpit_v2/web_api/main.py`
- `sap_rfc/fi_document_service.py`

## Contrato estável

- Entrada: `environment`, `branch` e `payload`.
- Saída: `fi_document_result` no job e log final do worker.
- Runtime compatível:
  - RFC/PyRFC em Windows.
  - API pode rodar fora do Windows.

## Quando mexer

Só altere este fluxo se houver:

- mudança de contrato da API;
- mudança de runtime da RFC;
- mudança explícita do formato do resultado FI.

Se a alteração for apenas visual ou de organização, não toque neste caminho crítico.
