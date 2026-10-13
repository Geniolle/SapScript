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

## Acesso e Servidor (Oracle Cloud)

O Cockpit Web foi migrado para a nuvem. Os dados de acesso e operação encontram-se em:
- Documentação detalhada: [docs/MIGRACAO_ORACLE_CLOUD.md](../docs/MIGRACAO_ORACLE_CLOUD.md)
- **URL Web:** [http://140.238.123.30](http://140.238.123.30)
- **Utilizador:** `Geniolle` | **Senha:** `P@1geniolle.`
- **Acesso SSH:** `ssh -i "C:\workspace\SapScript\Servidor\ssh-key-2026-10-13.key" ubuntu@140.238.123.30`

