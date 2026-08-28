# Procedimento adotado

Repositório GitHub: https://github.com/Geniolle/SapScript.git

## Fluxo Anterior

```text
VSCode -> Run no ficheiro SAP cockpit -> menus no terminal -> SAP GUI
```

## Fluxo Atual

```text
Pagina web -> cria job -> worker Windows -> run_sap_cockpit(payload) -> SAP GUI
```

## Regra de runtime

O caminho correto para este projeto está fixado em:

- [docs/RUNTIME_POLICY.md](RUNTIME_POLICY.md)
- [docs/FI_DEFAULT_DOCUMENT_FLOW.md](FI_DEFAULT_DOCUMENT_FLOW.md)

Resumo operacional:

- O `web_api` não deve executar o FI diretamente no container.
- `WORKFLOW_PYTHON_EXEC` deve apontar para um Python Windows.
- `SAP_FI_BRIDGE_PYTHON` deve apontar para o Python Windows com `pyrfc`.
- O resultado FI do job é gravado via API no fluxo isolado documentado em `FI_DEFAULT_DOCUMENT_FLOW.md`.
- Não substituir este meio por WSL/Linux nem por `sys.executable` fora de Windows.
- A comparação entre o F110 atual e um `ZFI_EXTERNAL_PAYMENTS` está em `F110_VS_ZFI_EXTERNAL_PAYMENTS.md`.
- Se `QAD` estiver indisponível e a intenção for apenas validar informação, usar a ligação produtiva em modo somente leitura, sem criar job ou proposta.

## Ficheiro do cockpit

Copia `sap_cockpit_web_ready.py` para a raiz do teu projeto SAP Script, no
mesmo nivel onde consegues importar `app.config` e `app.ui`.

Podes testar no terminal, mantendo o comportamento antigo:

```powershell
python sap_cockpit_web_ready.py
```

A diferenca e que agora o ficheiro tambem expoe `run_sap_cockpit(payload)`,
que sera chamada pelo worker Windows.

## Subir a interface web

Na pasta do pacote:

```bash
docker compose up --build
```

Abrir:

```text
http://localhost:8000
```

Se houver alterações no `web_api` ou em ficheiros carregados pelo container, o
Docker precisa ser reiniciado para aplicar as mudanças.

## Preparar o worker Windows

No Windows:

```powershell
cd worker
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-windows.txt
```

Se o `FI` for usado, confirma também que a venv RFC existe e que o `.env` ou os
scripts de arranque definem `SAP_FI_BRIDGE_PYTHON` para esse Python.

## Executar o worker Windows

Ajusta `SAP_SCRIPT_PROJECT_DIR` para a pasta raiz do teu projeto SAP Script
original.

```powershell
$env:API_BASE_URL = "http://localhost:8000"
$env:WORKER_TOKEN = "change-me"
$env:SAP_SCRIPT_PROJECT_DIR = "C:\\Users\\teu_user\\Documents\\SAP_SCRIPT"
$env:SAP_COCKPIT_MODULE = "sap_cockpit_web_ready"
$env:JIRA_SYNC_PROJECTS = "IT - Salsa Jeans, SAP - Desenvolvimento"
$env:JIRA_SYNC_JQL = ""
$env:JIRA_AUTO_TRIGGER_PROJECTS = "IT - Salsa Jeans, SAP - Desenvolvimento"
$env:JIRA_AUTO_TRIGGER_SUPPLIERS = "Evolutive"
$env:JIRA_AUTO_TRIGGER_ASSIGNEES = "Clayton Lopes"
python worker.py
```

`JIRA_SYNC_PROJECTS` aceita múltiplos projetos separados por vírgula. Se
`JIRA_SYNC_JQL` estiver preenchido, ele substitui a JQL automática.

O intervalo de sincronização em background é controlado por `POLL_SECONDS`.

A sincronização JIRA corre automaticamente em background quando a aplicação
web arranca.

## Primeiro teste recomendado

1. Subir o Docker.
2. Abrir o SAP GUI e fazer login no ambiente pretendido.
3. Iniciar o worker Windows.
4. Na web, executar primeiro `Ler STATUS atual do SAP`.
5. Depois executar `Abrir transacao` com `SE10`.
6. Por fim executar `Executar SAP Cockpit` preenchendo ambiente, processo e subprocesso.

## Campos minimos para executar o Cockpit

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "4"
}
```

## Exemplo com request existente

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "1",
  "request_number": "S4QK900396"
}
```

## Exemplo criando nova request

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "2",
  "request_type": "1",
  "request_desc": "REQUEST CRIADA VIA WEB"
}
```

## Configuração de Auto-Trigger por Categoria (JIRA)

O sistema de Auto-Trigger lê categorias do JIRA para rotinas SAP. O mapeamento é definido pela variável `AUTO_TRIGGER_CATEGORY_MAP` no `.env`.

* **Categoria:** `"FI Extracto Cadeias de Pesquisa"`
  * **Processo:** `"Cadeias de Pesquisa"`
  * **Subprocesso:** `"Criar Atribuir Cadeias.py"`
  * **Ambiente SAP (Sistema):** `DEV` (DESENVOLVIMENTO)

## Headroom

Comandos úteis para usar o proxy local do Headroom com as ferramentas do
projeto:

```powershell
$env:PATH += ";C:\Users\clayton.silva\AppData\Roaming\Python\Python312\Scripts"
headroom proxy --port 8787
```

```text
agy
agy --dangerously-skip-permissions
```

```powershell
$env:OPENAI_BASE_URL="http://127.0.0.1:8787/v1" ; codex --approve-for-me
```

```powershell
$env:ANTHROPIC_BASE_URL="http://127.0.0.1:8787" ; claude --dangerously-skip-permissions
```

