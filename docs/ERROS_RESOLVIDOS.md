# Erros resolvidos — registo de aprendizagem

Registo de incidentes já ultrapassados, para não os reinvestigar do zero.
Cada entrada segue o mesmo formato: **Sintoma → Causa raiz → Correção → Como diagnosticar → Como prevenir → Ficheiros**.

Ordenar do mais recente para o mais antigo.

---

## 2026-09-03 — Cockpit web: lista de Tickets vazia, "nunca sincronizado"

### Sintoma
- Página `http://127.0.0.1:8010/` abre, mas a "Fila de Tickets" fica vazia.
- O texto "Última sincronização" nunca aparece.
- Consola do browser: `Uncaught ReferenceError: Cannot access 'ASI_MAIN_MENU_ACTION' before initialization at (index):4835`.
- Nos logs do container só se veem `GET /` e `GET /static/styles.css` — **nenhuma chamada a `/api/...`**.

### Causa raiz
Alteração (não commitada) no `sap_script_web_cockpit_v2/web_api/templates/index.html`:
dentro do array `const salsaAgentActions = [ … ]` foi inserido `...ASI_MAIN_MENU_ACTION`,
mas o `const ASI_MAIN_MENU_ACTION = {…}` só era declarado ~560 linhas mais abaixo.

`salsaAgentActions` é avaliado no momento em que o `<script>` corre. Ao chegar ao spread,
`ASI_MAIN_MENU_ACTION` ainda está na **Temporal Dead Zone** (TDZ) de um `const` →
`ReferenceError` **não apanhado, no nível de topo do script** → o browser aborta a
execução de **todo o `<script>`** nesse ponto.

Consequência: nunca chegam a existir/correr `loadJobs()`, `switchView()`,
`loadJiraTickets()`, os `DOMContentLoaded`, nem o arranque
`loadJobs().then(() => { startPolling(); switchView('jira'); })`. Zero pedidos à API.

O backend (Jira → BD SQLite → `/api/jira/tickets`) esteve sempre funcional
(`POST /api/jira/sync` devolvia `synced_count: 444`). O corte era 100% no browser.

### Correção
Mover `const ASI_MAIN_MENU_ACTION = {…}` para **antes** de `const salsaAgentActions = [`.
O objeto é estático (`id`, `label`, `icon`) e não depende de mais nada, portanto sobe sem
efeitos secundários. Remover a declaração original que ficava depois.

### Como diagnosticar (da próxima vez)
1. **Logs do container primeiro**: `docker logs --since 20m --timestamps sap-script-web`.
   Se a página carrega mas **não há pedidos `/api/...`**, o problema é o JS do browser a
   abortar no arranque — não é o backend.
2. Confirmar o backend isoladamente:
   - `docker exec sap-script-web python -c "from web_api.store import list_jira_tickets; print(len(list_jira_tickets(limit=100000, exclude_closed=False)))"`
   - chamada direta à API Jira de dentro do container (ver `sap_script_web_cockpit_v2/check_jira.py`).
3. Pedir sempre a **consola do browser (F12)** e o separador **Network**. O `ReferenceError`
   com número de linha aponta diretamente para o `index.html` renderizado.
4. Reproduzir sem browser: extrair o bloco `<script>` do HTML servido, neutralizar Jinja
   (`{{…}}` → valor, `{%…%}` → vazio) e correr em Node com mocks de `document`/`window`/`fetch`.
   Um `UNCAUGHT: ReferenceError …` confirma o problema.

### Como prevenir
- **`node --check` NÃO chega**: TDZ é erro de *runtime*, não de sintaxe. O ficheiro passava
  no `--check`.
- Declarar `const`/`let` **antes** do primeiro uso, sobretudo quando são espalhados (`...`)
  dentro de estruturas de dados avaliadas de imediato (arrays/objetos de topo).
- Ao editar o `<script>` monolítico do `index.html`, validar sempre carregando a página
  com a consola aberta (ver secção "Cuidado com scripts monolíticos" no `CLAUDE.md`).

### Ficheiros
- `sap_script_web_cockpit_v2/web_api/templates/index.html` (`ASI_MAIN_MENU_ACTION`, `salsaAgentActions`)

### Pendente relacionado (não é a causa deste incidente)
- `sap_script_web_cockpit_v2/web_api/store.py` → `save_jira_tickets_to_db()` faz
  `DELETE FROM jira_tickets WHERE key NOT IN (...)`. Se `fetch_jira_tickets_from_api()`
  devolver lista parcial (falha a meio da paginação) ou vazia (credenciais em falta),
  apaga tickets válidos / limpa a tabela toda. Endurecer: só sincronizar destrutivamente
  quando a busca ao Jira for confirmada como completa.
