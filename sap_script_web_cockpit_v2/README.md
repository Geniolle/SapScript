# SAP Script Web Cockpit

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

## 1. Copiar o ficheiro do Cockpit para o projeto SAP Script

Copia este ficheiro:

```text
sap_cockpit_web_ready.py
```

para a raiz do teu projeto SAP Script, no mesmo nivel onde consegues importar `app.config` e `app.ui`.

Podes testar no terminal, mantendo o comportamento antigo:

```powershell
python sap_cockpit_web_ready.py
```

A diferenca e que agora o ficheiro tambem expoe:

```python
run_sap_cockpit(payload)
```

que sera chamada pelo worker Windows.

## 2. Subir a interface web com Docker

Na pasta deste pacote:

```bash
docker compose up --build
```

Abrir:

```text
http://localhost:8000
```

## 3. Preparar o worker Windows

No Windows:

```powershell
cd worker
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-windows.txt
```

## 4. Executar o worker Windows

Ajusta `SAP_SCRIPT_PROJECT_DIR` para a pasta raiz do teu projeto SAP Script original.

Exemplo:

```powershell
$env:API_BASE_URL = "http://localhost:8000"
$env:WORKER_TOKEN = "change-me"
$env:SAP_SCRIPT_PROJECT_DIR = "C:\\Users\\teu_user\\Documents\\SAP_SCRIPT"
$env:SAP_COCKPIT_MODULE = "sap_cockpit_web_ready"
python worker.py
```

## 5. Como executar pela web

Na pagina, escolher:

```text
Rotina: Executar SAP Cockpit
Ambiente: S4Q
Processo / pasta: nome exato da pasta dentro de PROCESSOS_DIR
Subprocesso / ficheiro .py: nome exato do script .py
```

Se o subprocesso pedir Excel, preencher o caminho completo do ficheiro no Windows:

```text
C:\SAP\ficheiro.xlsx
```

## 6. Requests

A web suporta:

```text
4 - Nao transportar
1 - Usar request existente
2 - Criar nova request
```

A opcao `3 - Pesquisar suas request criadas` continua no modo terminal, mas nao foi ativada na web porque exige uma escolha manual numa lista. Podemos transformar isso depois numa pagina propria de pesquisa/selecao.

## 7. Regra obrigatoria de STATUS

Toda execucao termina devolvendo:

```python
session.findById("wnd[0]/sbar").Text
```

No ficheiro novo, isto ficou centralizado em:

```python
read_sbar_status(session)
```

Mesmo quando ocorre erro, o retorno tenta preencher `STATUS` com o texto real de `wnd[0]/sbar`.
