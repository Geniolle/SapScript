# SAP Script Web Cockpit

Repositório GitHub: https://github.com/Geniolle/SapScript.git

Este pacote liga uma pagina web em Docker ao teu SAP Cockpit atual, sem executar SAP GUI dentro do container.

## Arquitetura

```text
Navegador
  -> FastAPI em Docker, pacote web_api
  -> SQLite com fila de jobs
  -> Worker Python nativo no Windows
  -> Modulo sap_cockpit_web_ready no teu projeto SAP Script
  -> SAP GUI Scripting
  -> STATUS vindo de wnd[0]/sbar
```

## Por que `web_api` e nao `app`?

O teu projeto SAP Script atual ja usa imports como:

```python
from app.config import ...
from app.ui import ...
```

Por isso a aplicacao web foi colocada no pacote `web_api`, para nao criar conflito com o pacote `app` do teu projeto SAP original.

## Procedimento

Os passos operacionais, formatos de execução, regras da interface web e comandos
do proxy local ficam centralizados em:

- [docs/PROCEDIMENTO.md](docs/PROCEDIMENTO.md)
- [docs/RUNTIME_POLICY.md](docs/RUNTIME_POLICY.md)
- [docs/FI_DEFAULT_DOCUMENT_FLOW.md](docs/FI_DEFAULT_DOCUMENT_FLOW.md)

Use esse documento como referência única para a operação diária. Aqui fica apenas a visão geral do pacote e a ponte para o procedimento.
