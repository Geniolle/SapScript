# Refatoracao do Agente Salsa IT

Objetivo: reduzir a fragilidade das camadas de orquestracao (frontend monolitico,
`main.py` com boilerplate, dispatch do worker) **sem alterar comportamento**, para
deixar de "estragar coisas que funcionam" a cada edicao.

Branch: `refactor/agente-salsa-it`. Uma fase por commit, cada uma revertivel.

## Rede de seguranca

Correr **antes e depois** de cada fase, a partir de `sap_script_web_cockpit_v2/`:

```powershell
.\.venv\Scripts\python.exe tests\run_all.py
```

(`js_smoke` precisa de `node` no PATH; os testes usam o venv do cockpit.)

Cobre: o `<script>`/`static/js` avalia sem excecao no topo (TDZ, `const` antes de
declarar, funcoes duplicadas, condicoes constantes, globais de arranque em falta);
contrato das rotas `/api/salsa-it-agent/pfcg/*`; limpeza de jobs orfaos.

**Limite**: nao valida os fluxos de chat PFCG ponta a ponta (selecao Excel ->
analise -> criacao, transporte, composta, delete). Isso exige clicar na UI com o
worker + SAP a correr.

## Estado

| Fase | Descricao | Estado | Commit |
|---|---|---|---|
| 0 | Rede de seguranca (`tests/`) | **Feito** | `Fase 0: rede de seguranca` |
| 1 | Remover footguns confirmados (2 funcoes duplicadas mortas + `if (false ...)`) | **Feito** | `Fase 1: remover footguns confirmados` |
| 2a | Extrair o `<script>` inline (11 230 linhas) para `web_api/static/js/cockpit.js` + `window.__COCKPIT__` | **Feito** | `Fase 2a: extrair o <script> inline` |
| — | `js_smoke` avalia cada ficheiro como `<script>` separado (apanha quebras de ordem de carregamento) | **Feito** | `js_smoke: avaliar cada ficheiro como <script> separado` |
| 4 (parcial) | Limpar jobs orfaos do worker no arranque (`reap_orphan_running_jobs` + endpoint + hook) | **Feito** | `Fase 4 (parcial): limpar jobs orfaos` |
| 4 (resto) | `worker/sap_tasks.py`: cadeia `if task ==` -> `TASK_HANDLERS` dict (`sap_cockpit` fica explicito) | **Feito** | `Fase 4 (resto): dispatch do worker por dict` |
| 3 | Extrair helpers/estado partilhados de `main.py` -> `web_api/common.py` + `web_api/pfcg_common.py` (rotas ficam em main.py) | **Feito** | `Fase 3: extrair helpers/estado partilhados` |
| 2b | Dividir `cockpit.js` -> `cockpit.core.js` + `cockpit.agent.js` (corte na linha 2381) | **Feito** | `Fase 2b: dividir cockpit.js em 2 modulos` |
| 1b | Unificar os ~9 loops de polling num `asiPollJob({...})`; `asiSetState(patch)` | **Pendente** | — |
| 5 | Substituir a cascata `awaitingInput` por tabela de estados explicita | **Pendente** — depende de 1b | — |

### Pendentes: 1b e 5 exigem verificacao por fluxo

Sao os dois unicos passos que **nao** sao mecanicos: 1b e um port a mao do corpo
"sucesso/erro" de cada um dos ~9 pollers (`asiStartPfcgPolling`,
`asiStartPfcgCreateExcelSelectionPolling`, `asiStartPfcgCreateExcelAnalysisPolling`,
`asiPollPfcgSubAnalysis`, `asiPollPfcgIndividualPreview`, `asiPollPfcgIndividualConfirm`,
`asiPollPfcgTransportSearch`, `asiPollPfcgDeletePreview`, `asiPollPfcgDeleteConfirm`)
para um callback `onResolved(data)`; cada um chama renderers e estados
especificos (`asiPfcgRoleState`, `asiBuildPfcg*Html`, `ASI_PFCG_ROLE_RESULT_ACTIONS`, ...).

O `js_smoke` confirma que a pagina carrega e o contrato das rotas aguenta, **nao**
confirma que o resultado renderiza igual. Por isso 1b faz-se **um poller por
commit**, e cada commit so se fecha depois de clicar esse fluxo na UI com o
worker + SAP. 5 so depois de 1b.

## Verificacao exigida pelas fases pendentes

As fases pendentes mudam caminhos que a rede de seguranca atual **nao** cobre.
Antes de dar cada uma por concluida, com o worker Windows + ligacao SAP ativos:

1. **1b (pollers)** — para cada fluxo migrado, clicar ate ao fim e confirmar que
   o resultado renderiza e o composer volta a ficar disponivel:
   - Funcoes PFCG -> Analisar por nome / por Transacao / por Utilizador
   - Criar funcoes -> Selecionar Excel -> Analisar -> (preview) -> Confirmar
   - Criar Individualmente (nome/descricao/tcodes -> preview -> confirmar)
   - Funcao Composta (preview -> confirmar)
   - Eliminar funcao (preview -> confirmar)
   - Procurar/!criar Ordem de Transporte
   - Testes Unitarios -> Criar Documento FI Default / Executar F110 Default
2. **2b (modulos)** — `tests\run_all.py` + carregar a pagina com a consola aberta
   (F12): zero erros, e repetir o checklist do ponto 1 (uma vez chega).
3. **3 (routers)** — `tests\run_all.py` + `curl` a cada rota `/api/salsa-it-agent/*`
   (POST devolve `job_id`, GET mapeia estado). O container **nao** faz hot-reload
   de codigo Python: `docker compose restart sap-script-web` depois de mexer.
4. **4 resto (worker)** — reiniciar o worker e correr um job de cada `task` PFCG.
5. **5 (estado)** — checklist do ponto 1 completo (todos os passos que pedem input
   de texto ao utilizador).

## Notas de ambiente

- **Hot reload**: `index.html` e lido pelo Jinja a cada pedido (muda "a quente");
  `main.py` / `store.py` / `worker.py` **nao** — exigem `docker compose restart
  sap-script-web` (web) ou reiniciar o `start_worker_auto.ps1` (worker).
- **CSP**: nao ha header Content-Security-Policy; `/static` esta montado. Carregar
  JS de `/static/js/` e seguro.
- **Job orfao pendente** (incidente de 2026-09-03): limpa-se no proximo arranque
  do worker (novo hook), ou a mao:
  `curl -X POST -H "X-Worker-Token: <token>" ".../api/worker/jobs/reap-orphans?worker_name=<NOME>"`

Ver tambem `docs/ERROS_RESOLVIDOS.md`.
