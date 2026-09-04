# Erros resolvidos — registo de aprendizagem

Registo de incidentes já ultrapassados, para não os reinvestigar do zero.
Cada entrada segue o mesmo formato: **Sintoma → Causa raiz → Correção → Como diagnosticar → Como prevenir → Ficheiros**.

Ordenar do mais recente para o mais antigo.

---

## 2026-09-04 — Criar Individualmente (PFCG): ecrã "Deseja utilizar ordem de transporte?" sem opção de voltar

### Sintoma
- No fluxo "Criar Individualmente", depois de indicar nome, descrição e transações, o
  ecrã "Deseja utilizar ordem de transporte em DEV?" (3 botões: "Sem transporte (Local)",
  "Criar nova Request", "Usar Request existente") não tinha nenhum botão "← Voltar" —
  se o utilizador se enganou na descrição/transações, não havia forma de corrigir sem
  reiniciar a conversa. Os ecrãs vizinhos do mesmo fluxo (lista de Requests existentes,
  pré-visualização final) já tinham "← Voltar".

### Correção
Em `asiAskPfcgTransportMode()`, reutilizar o mesmo botão/handler já usado na
pré-visualização (`ASI_PFCG_INDIVIDUAL_BACK_ACTION` → `asiHandlePfcgIndividualBack()`,
que reinicia o fluxo desde o nome da função) — evita duplicar um handler quase igual.

### Ficheiros
- `sap_script_web_cockpit_v2/web_api/static/js/cockpit.agent.js` (`asiAskPfcgTransportMode`)

---

## 2026-09-04 — Criar Individualmente (PFCG): cockpit fica preso em "A verificar se <função> já existe..." mesmo com o job terminado com sucesso

### Sintoma
- Fluxo "Criar Individualmente": ao indicar o nome da função (ex.: `Z_TESTE_WEB_V3`), o
  cockpit mostra "A verificar se Z_TESTE_WEB_V3 já existe em DEV..." (com spinner) e **fica
  preso nesse texto indefinidamente**, mesmo que a função já esteja confirmada como
  inexistente (validado diretamente via RFC) e o utilizador tenha esperado bastante mais
  do que os poucos segundos normais de resposta.

### Causa raiz
Falso alarme na camada backend/worker: o job `pfcg_role_analysis` correspondente
(`state='succeeded'`, resultado `{"ok": true, "status": "NAO_EXISTE", ...}`) tinha
terminado em ~3 segundos, e o endpoint `GET /api/salsa-it-agent/pfcg/analyze/{job_id}`
devolvia exatamente a forma esperada pelo frontend. O problema estava só no frontend.

Em `asiRenderMessages()` (`cockpit.agent.js`), o conteúdo de cada balão é `msg.html` se
existir, caindo para `msg.text` só quando `msg.html` é vazio/falsy — `msg.html` tem
sempre prioridade. A mensagem de "a verificar..." é criada com
`html: asiBuildPfcgGenericProcessingHtml(...)` (o spinner). No ramo de sucesso
"não existe" de `asiPollPfcgCreateExistsCheck()`, o `asiUpdateMessage(messageId, {...})`
só atualizava `text` e `isProcessing`, **sem limpar `html`** — por isso o balão
continuava a renderizar o spinner antigo para sempre, apesar de o `text` interno já ter
mudado e de a pergunta seguinte ("Qual é a descrição...") já ter sido adicionada por
baixo. Todos os outros ramos (EXISTE, erros, timeout) já definiam um novo `html`
(`asiBuildPfcgErrorHtml(...)`) e por isso não sofriam do problema — só faltava neste
ramo específico.

### Correção
Em `asiPollPfcgCreateExistsCheck()`, no ramo `NAO_EXISTE`, incluir `html: ''` na
chamada a `asiUpdateMessage`, para o renderizador cair de volta no `text` atualizado.

### Como diagnosticar (da próxima vez)
1. Confirmar primeiro que o job realmente terminou (`docker exec sap-script-web
   python3` a consultar `/data/sap_script_jobs.sqlite3`, tabela `jobs`, coluna
   `state`/`status`) e que a API devolve a forma esperada (`curl` direto ao endpoint de
   polling) — isto isola se o problema é backend/worker ou frontend.
2. Com backend e API confirmados corretos, procurar no frontend por
   `asiUpdateMessage(messageId, {...})` que define `isProcessing: false` **sem** também
   definir `html` — comparar com os ramos irmãos (erro/timeout) da mesma função, que
   normalmente definem sempre os dois.
3. Lembrar que `asiRenderMessages()` dá sempre prioridade a `msg.html` sobre `msg.text`
   — atualizar só o `text` de uma mensagem que já tem `html` não muda nada visualmente.

### Como prevenir
- Sempre que uma mensagem for criada com `html` (spinner/processamento), qualquer
  atualização posterior que a "resolva" (sucesso, erro ou timeout) deve explicitamente
  definir um novo `html` (mesmo que seja `''`) — nunca assumir que só mudar o `text`
  chega.

### Ficheiros
- `sap_script_web_cockpit_v2/web_api/static/js/cockpit.agent.js`
  (`asiPollPfcgCreateExistsCheck`, `asiRenderMessages`)

---

## 2026-09-04 — Criar Individualmente (PFCG via RFC): função criada com sucesso mas nunca vinculada à Request de transporte

### Sintoma
- Fluxo "Configurações > Perfil de Autorização > DEV > Função Simples > Preparar criação >
  Criar Individualmente", escolhendo "Nova Request de transporte": o cockpit reporta sucesso
  total (`✓ Função criada em DEV`, `CREATED`, `Perfil gerado: Sim`, `Transporte: Nova Request —
  S4DK953705`), a função e a Request existem mesmo em SAP, mas a função **não fica atribuída**
  a essa Request (confirmado pelo utilizador ao verificar em SAP; SE01/SE09 não mostram a
  função sob a Request/task).

### Causa raiz
`PRGN_RFC_CREATE_ACTIVITY_GROUP` aceita um parâmetro `REQUEST=<número>`, não lança nenhuma
exceção e devolve `NEW_REQUEST` preenchido, **mas nunca regista de facto o objeto em E071**
(confirmado por leitura real, read-only, via `RFC_READ_TABLE` em E071 filtrando pela Request
criada — `TRKORR = 'S4DK953705'` — e pela sua task automática — `TRKORR = 'S4DK953706'`: ambas
devolvem `TABLE_WITHOUT_DATA`, i.e. zero linhas). A verificação pós-escrita existente em
`create_pfcg_role_rfc()` já confirmava a função/textos/tcodes/perfil de forma independente,
mas nunca verificava o vínculo ao transporte em si — por isso o "sucesso" reportado era real
para a função, mas enganador quanto ao transporte.

Este projeto já tinha documentado (`sap_rfc/pfcg_gui_transport.py`, escrito antes deste
incidente) que não existe caminho por RFC/BAPI para isto: `TR_OBJECT_INSERT` (a RFC óbvia para
"inserir objeto numa request") **não é remote-enabled** neste sistema — o facto de
`connection.get_function_description('TR_OBJECT_INSERT')` funcionar sem erro **não prova que a
função seja invocável via RFC**: essa chamada só lê metadados (via `RFC_GET_FUNCTION_INTERFACE`),
que funcionam mesmo para function modules sem a flag "Remote-Enabled" — o erro real só aparece
ao tentar `connection.call(...)` de verdade. `PFCG_MASS_TRANSPORT` (o report por trás do botão
"Transportar" da PFCG) também está confirmado como bloqueado em execução background/job (mesmo
via BAPI_XBP_*). A única via funcional confirmada é reproduzir, via SAP GUI Scripting, o mesmo
caminho manual de um utilizador (SE38 → `PFCG_MASS_TRANSPORT` em primeiro plano).

Esse mecanismo (`assign_role_to_transport()` em `sap_rfc/pfcg_gui_transport.py`) já existia no
código, testado e documentado, mas **nunca estava ligado a nada** — zero chamadores em todo o
projeto. O worker chamava `create_pfcg_role_rfc()` (via bridge de subprocesso isolado
`.venv-rfc`, sem `pywin32`) e terminava ali.

### Correção
Em `sap_script_web_cockpit_v2/worker/sap_tasks.py` (`_run_pfcg_role_create_rfc`), depois do
bridge RFC devolver sucesso com uma `transport_request`, chamar
`sap_rfc.pfcg_gui_transport.assign_role_to_transport(environment, role_name, transport_request)`
— isto corre no processo principal do worker (que já importa `win32com.client`/`pythoncom` a
nível de topo), nunca no subprocesso `.venv-rfc` (sem `pywin32`, só usado para a chamada RFC de
criação). O resultado fica em `payload["transport_assignment"]`, propagado ao frontend.

Frontend (`asiBuildPfcgIndividualResultHtml` em `cockpit.agent.js`): quando existe
`result.transport_request`, mostra uma linha "Vínculo ao transporte" (✓ Confirmado / ⚠ Requer
ação manual) e, se `transport_assignment.ok !== true`, uma nota com a mensagem de erro/fallback
(`assign_role_to_transport()` já devolve `MANUAL_FALLBACK_REQUIRED` com instruções PFCG quando
não há sessão SAP GUI ativa ligada ao ambiente).

### Como diagnosticar (da próxima vez)
1. Nunca confiar apenas na ausência de exceção de `PRGN_RFC_CREATE_ACTIVITY_GROUP` para o
   vínculo ao transporte — só a leitura real de E071 (`RFC_READ_TABLE`, filtrando por
   `TRKORR` = Request e pelas suas tasks filhas via `E070-STRKORR`) confirma o vínculo.
2. `RFC_READ_TABLE` nesta ligação lança `ABAPApplicationError(key=TABLE_WITHOUT_DATA)` quando a
   seleção não devolve linhas em E070/E071 — tratar isso como "zero linhas", não como erro real.
3. `connection.get_function_description(func)` só prova que a função **existe e tem uma
   assinatura conhecida** — nunca que é remote-enabled/invocável. Só uma chamada real
   (`connection.call(...)`) confirma isso.
4. Objeto de transporte real confirmado por leitura em E071 nesta ligação: `PGMID='R3TR'`,
   `OBJECT='ACGR'` para Activity Groups/Funções PFCG (roles standard `/AIF/*` e customizadas já
   existentes como `Z_TEST`, `ZPEGGING`).

### Como prevenir
- Antes de escrever uma nova chamada RFC "óbvia" para preencher uma lacuna, procurar primeiro
  no próprio repositório (`grep` por nomes de função SAP candidatos) — já podia existir uma
  investigação anterior documentada com a resposta certa (como aqui, em
  `sap_rfc/pfcg_gui_transport.py`) em vez de repetir o mesmo teste às cegas.
- Ao adicionar um novo passo de automação a um fluxo de criação, confirmar sempre que fica
  realmente ligado a um chamador (`grep` por callers) — código correto mas nunca invocado
  produz o mesmo sintoma de bug que código inexistente.

### Ficheiros
- `sap_script_web_cockpit_v2/worker/sap_tasks.py` (`_run_pfcg_role_create_rfc`)
- `sap_rfc/pfcg_gui_transport.py` (`assign_role_to_transport`, pré-existente, agora ligado)
- `sap_rfc/pfcg_role_create_service.py` (`create_pfcg_role_rfc` — investigado, não alterado)
- `sap_script_web_cockpit_v2/web_api/static/js/cockpit.agent.js` (`asiBuildPfcgIndividualResultHtml`)

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
