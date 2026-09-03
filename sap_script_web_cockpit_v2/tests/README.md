# Testes do cockpit (`sap_script_web_cockpit_v2`)

Rede de seguranca minima introduzida com a refatoracao do **Agente Salsa IT**.
Correr **antes e depois** de cada fase da refatoracao.

## Como correr

A partir de `sap_script_web_cockpit_v2/`, com o Python do venv do cockpit
(tem `fastapi`/`pydantic`); o `js_smoke` precisa de `node` no PATH.

```powershell
.\.venv\Scripts\python.exe tests\run_all.py
```

Ou individualmente:

```powershell
.\.venv\Scripts\python.exe tests\js_smoke.py
.\.venv\Scripts\python.exe -m unittest tests.test_salsa_agent_routes -v
```

## O que cobre

| Ficheiro | Alvo | Apanha |
|---|---|---|
| `js_smoke.py` | `<script>` inline de `web_api/templates/index.html` | erro de runtime no topo do script (TDZ, `const` antes de declarar), funcoes de topo duplicadas, `if (false ...)`/condicao constante, funcoes globais de arranque em falta |
| `test_salsa_agent_routes.py` | rotas `/api/salsa-it-agent/pfcg/*` em `web_api/main.py` | criacao de job + `job_id`, rejeicao de nome de role invalido (400), mapeamento de estado no GET, job inexistente (404), job de outra task (400), whitelist de campos no resultado |

`js_smoke.py` faz SKIP se `node` nao estiver no PATH.
Os testes de rotas redirecionam a BD de jobs para um diretorio temporario
(`DATA_DIR`), nao tocam na base real.

## Nota

`js_smoke.py` e um teste de fumo, nao um substituto de lint completo. O objetivo
e travar exatamente a classe de bug que derrubou o cockpit (ver
`docs/ERROS_RESOLVIDOS.md`, incidente do `ASI_MAIN_MENU_ACTION`).
