# SAP Cockpit — Auditoria para Refatoração (Fase 1)

Branch: `refactor/sap-cockpit-core` (criada a partir de `main`, com as alterações locais pendentes herdadas, não commitadas).

## 1. Inventário

| Componente | Ficheiro | Tamanho / defs |
|---|---|---|
| Cockpit desktop (terminal original) | `SAP Cockpit.py` | ~36 KB, 39 funções, sem `main` (código de topo) |
| Cockpit web/terminal "pronto" | `sap_script_web_cockpit_v2/sap_cockpit_web_ready.py` | ~49 KB, 50 funções, `run_sap_cockpit`, `main_terminal` |
| Worker SAP Windows | `sap_script_web_cockpit_v2/worker/sap_tasks.py` | ~126 KB, `TASK_HANDLERS` com 56 entradas |
| Loop do worker | `worker/worker.py` | — |
| API FastAPI | `web_api/main.py` (~157 KB), `web_api/store.py` (~60 KB) | — |
| Serviços RFC | `sap_rfc/` (≈ 60 módulos) | — |

Seleção do módulo do cockpit no worker: `SAP_COCKPIT_MODULE` (default `sap_cockpit_web_ready`), em `worker/sap_tasks.py:806`.

## 2. Duplicação real

- 39 funções no `SAP Cockpit.py`; 50 em `sap_cockpit_web_ready.py`.
- **37 nomes de função coincidem** nos dois ficheiros. Não comparei ainda os corpos linha a linha: a afirmação de duplicação é ao nível do nome.
- Duplicados de maior impacto: leitura de `.env` (`_parse_env_line`, `_carregar_dotenv_manual`, `_carregar_dotenv`, `_obter_credenciais_env`), ligação SAP (`_safe_find`, `_press`, `_send_vkey`, `_set_text`, `_ensure_se10`), gestão de request SE10 (`_criar_nova_request_no_sap`, `perguntar_opcao_request`, `_extract_*`), sessão (`_is_sap_logado`, `_aguardar_login`, `_encontrar_sessao_do_sistema`, `_tem_alguma_sessao_ativa`) e execução (`executar_processo`, `_analisar_exec_signature`).
- Diferenças que o Core deve preservar: a versão web usa `interactive: bool` e lança `SapCockpitError` em vez de `sys.exit`/`input()`; a versão web tem `_build_exec_kwargs`, `_load_process_module` e `_executar_um_script`, sem equivalente no desktop.

## 3. Riscos e pontos de atenção

1. **Credenciais expostas (corrigido nesta fase, ainda por rodar):** `sap_script_web_cockpit_v2/README.md` e `docs/MIGRACAO_ORACLE_CLOUD.md` continham utilizador e senha em claro. Foram substituídos por texto neutro nesta branch. **A senha continua no histórico do Git** e deve ser considerada comprometida: rodar a senha no servidor Oracle e limpar o histórico só com autorização explícita.
2. **Ficheiros com alterações locais não commitadas** que a refatoração toca: `worker/sap_tasks.py`, `web_api/main.py`, `sap_rfc/hr_lookup_service.py`, `salsa_agent/*`, `templates/index.html`, `cockpit.agent.js`, `cockpit.core.js`, `styles.css`, e `data/sap_script_jobs.sqlite3`. Commitar ou separar estas alterações antes da fase 2, para não misturar trabalho do utilizador com a refatoração.
3. **Ficheiros não rastreados** em `sap_rfc/`: `ZRGGBS000_LCL.abap` (0 bytes) e `otpm_service.py`. Não tocar até confirmar a origem.
4. `data/sap_script_jobs.sqlite3` está modificado e rastreado. Os testes podem escrever nele; não commitar.
5. Execuções SAP GUI simultâneas sobre a mesma sessão ainda não têm controlo central. Previsto na fase 6.
6. `_e2e_web_api.log` e `_e2e_worker.log` estão vazios e na raiz: candidatos a limpeza, só após validação.

## 4. Matriz de equivalência (esboço)

| Função | Terminal (`SAP Cockpit.py`) | Web (`sap_cockpit_web_ready.py`) | Estado |
|---|---|---|---|
| Selecionar ambiente | `selecionar_ambiente()` interativo | `selecionar_ambiente(payload, interactive)` | Mesmo nome, contrato diferente. Core deve receber dados, não ler `input()`. |
| Executar processo | `executar_processo(...)` | `executar_processo(...)` com `_build_exec_kwargs` | Requer comparação de assinaturas antes de unificar. |
| Obter sessão SAP | (ausente como função própria; ver `_encontrar_sessao_do_sistema`) | `obter_sessao_sap(...)` | Só web. Candidato a Core. |
| Entrada principal | Código de topo (sem função de entrada) | `run_sap_cockpit(payload)` / `main_terminal()` | Adaptadores distintos. |

Matriz completa (requests, credenciais, erros, processos) fica para a fase 2.

## 5. Baseline de testes (antes de qualquer alteração de código)

- Comando: `sap_script_web_cockpit_v2\.venv\Scripts\python.exe tests\run_all.py`
- Resultado: **65 testes OK**, `js_smoke` OK, `RESULTADO: OK`.
- Não foram executados testes de integração SAP/RFC/GUI. Não houve acesso a SAP.

## 6. Estado das fases

| Fase | Estado |
|---|---|
| 1. Inventário e testes de referência | Concluída (esta auditoria + baseline de 65 testes) |
| 2. Segurança e contratos | Concluída (commit `b34330e`, ver [SAP_COCKPIT_MIGRATION.md](SAP_COCKPIT_MIGRATION.md)); lacuna: desktop mantém fallbacks `TypeError` |
| 3. Auditoria das funções duplicadas | Concluída (secção 7): matriz, riscos e ordem de migração. Sem alterações de código |
| 3b. Segurança do desktop e núcleo de funções puras | Concluída parcialmente (secção 7.11): TypeError corrigido no desktop; `parse_env_line`, `is_scripting_disabled_error` e `validar_argumentos_exec` centralizados. Credenciais e requests aguardam decisão |
| 3c. Migrar executor e sessões | Pendente |
| 4. Adaptar o terminal | Pendente |
| 5. Integrar o worker Web | Pendente |
| 6. Reforçar segurança e rastreabilidade | Parcial: credenciais em README/migração removidas |
| 7. Testes de integração e regressão | Pendente (requer ambiente Windows com SAP GUI) |
| 8. Remover duplicações e atualizar documentação | Pendente |

Parar nesta fase é intencional: as fases 2 a 5 exigem decidir como tratar as alterações não commitadas (ponto 3.2) e validar o comportamento do worker Windows.

## 7. Fase 3 — Auditoria das funções duplicadas

### 7.1 Método

- Comparação por AST de `SAP Cockpit.py` (desktop) e `sap_script_web_cockpit_v2/sap_cockpit_web_ready.py` (web) ao commit `b34330e`. Os ficheiros não foram importados nem executados.
- "Lógica igual" = AST igual, sem docstring. Os nomes de variáveis contam: renomeações aparecem como diferença.
- Texto: `difflib` sobre o código-fonte de cada função, para detetar diferenças de formatação e mensagens.
- Limite: a comparação é estrutural. Não prova equivalência de comportamento em runtime; isso exige testes de equivalência (secção 7.6).
- Nenhuma operação SAP foi executada. Nenhum ficheiro `.env` foi lido.

### 7.2 Resumo

- Funções com o mesmo nome: 37 (desktop 39, web 50).
- Idênticas: 19. Semelhantes: 11. Divergentes: 7. Específicas (sem par no outro ficheiro): 2 no desktop e 15 na web.
- Nenhuma das 37 é específica de um só ambiente. As específicas aparecem na secção 7.5.

### 7.3 Matriz das 37 funções partilhadas

Linhas: desktop = `SAP Cockpit.py`, web = `sap_cockpit_web_ready.py` em `b34330e`.

| # | Função | Desktop (linhas) | Web (linhas) | Classificação | Nota |
|---|---|---|---|---|---|
| 1 | `_aguardar_login` | 941-948 | 1078-1085 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 2 | `_analisar_exec_signature` | 711-749 | 760-791 | Semelhante | Web acrescenta p_modo_nao_interativo e p_pedir_confirmacao |
| 3 | `_carregar_dotenv` | 167-172 | 184-189 | Semelhante | Só texto da mensagem de aviso |
| 4 | `_carregar_dotenv_manual` | 125-164 | 151-181 | Semelhante | Nome de variável local diferente (linha/linha_txt) e docstring |
| 5 | `_criar_nova_request_no_sap` | 560-612 | 555-620 | Divergente | ALTA: web usa tipo='1' e desc default quando o payload omite campos; cria request SE10 sem escolha explícita |
| 6 | `_dismiss_popup_if_any` | 922-938 | 1059-1075 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 7 | `_encontrar_sessao_do_sistema` | 971-991 | 1108-1125 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 8 | `_ensure_se10` | 456-463 | 453-460 | Semelhante | Só texto da mensagem de erro (acentos) |
| 9 | `_erro_scripting_inativo` | 903-910 | 1041-1049 | Divergente | Desktop termina com sys.exit(1); web lança SapCockpitError quando interactive=False |
| 10 | `_extract_request_by_scanning_usr` | 504-537 | 501-534 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 11 | `_extract_request_from_known_ids` | 493-501 | 490-498 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 12 | `_extract_request_number_from_text` | 466-474 | 463-471 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 13 | `_get_created_request_number` | 540-544 | 537-539 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 14 | `_is_sap_logado` | 913-919 | 1052-1056 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 15 | `_is_scripting_disabled_error` | 873-896 | 1020-1034 | Semelhante | Mesmo conjunto de padrões, lista reformatada |
| 16 | `_log_alerta_rz11` | 899-900 | 1037-1038 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 17 | `_log_scripting_status_apenas_quando_logado` | 951-954 | 1088-1091 | Semelhante | Só texto da mensagem (acentos) |
| 18 | `_obter_credenciais_env` | 175-209 | 205-241 | Divergente | ALTA: assinatura diferente; web resolve prefixo, aceita SAP_USER, SAP_LANGUAGE, SAP_PASSWORD_<SISTEMA>CLNT<cliente> e SAP_PASSWORD. Mesmo .env pode dar credenciais diferentes |
| 19 | `_parse_env_line` | 93-122 | 124-148 | Semelhante | Nome de variável local diferente (linha/linha_txt) |
| 20 | `_press` | 430-435 | 427-432 | Semelhante | Só texto da mensagem de erro (acentos) |
| 21 | `_resetar_env_request` | 405-409 | 402-406 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 22 | `_safe_find` | 421-427 | 418-424 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 23 | `_select_radio_if_exists` | 547-557 | 542-552 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 24 | `_send_vkey` | 438-440 | 435-437 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 25 | `_set_text` | 443-453 | 440-450 | Semelhante | Só texto da mensagem de erro (acentos) |
| 26 | `_sleep` | 417-418 | 414-415 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 27 | `_tem_alguma_sessao_ativa` | 957-968 | 1094-1105 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 28 | `_try_get_obj_text` | 477-490 | 474-487 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 29 | `carregar_pesquisar_request` | 371-381 | 372-380 | Semelhante | Só texto e formatação |
| 30 | `escolher_request_por_linha` | 384-402 | 383-399 | Semelhante | Só texto da mensagem (acentos) |
| 31 | `executar_processo` | 752-865 | 957-1012 | Divergente | Desktop: loop interativo com exec inline e fallback TypeError. Web: ramo não interativo + _executar_um_script (corrigido na Fase 2) |
| 32 | `ler_texto` | 61-90 | 96-121 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 33 | `limpar_buffer_teclado` | 53-58 | 88-93 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |
| 34 | `perguntar_opcao_request` | 620-703 | 628-752 | Divergente | Assinatura diferente (payload, interactive); desktop tem fallback TypeError em listar_requests |
| 35 | `selecionar_ambiente` | 294-336 | 244-272 | Divergente | Desktop lê SAP_ENVIRONMENT/SAP_ENV e stdin; web lê payload['ambiente'] e sys.argv, sem variáveis de ambiente |
| 36 | `selecionar_pasta_processo` | 339-359 | 336-360 | Divergente | Web aceita payload['processo']; desktop trata KeyboardInterrupt/EOFError e a web não |
| 37 | `validar_request` | 362-368 | 363-369 | Idêntica | Lógica AST igual; diferenças só de formatação ou docstring |

### 7.4 Duplicações comprovadas e divergências

**Duplicação de execução (ALTA, ainda presente no desktop).** A Fase 2 removeu o retry por `TypeError` só da web. `SAP Cockpit.py` mantém o padrão em `executar_processo` (linhas 860–862: `exec_fn(ambiente, **kwargs)` seguido de `exec_fn(ambiente)` em `except TypeError`). Um erro interno de um script depois de uma escrita SAP executa o script duas vezes. Há também o fallback de `listar_requests` (linha 676). **Corrigir antes de qualquer centralização.** A Fase 2 não cobriu este ponto e deve ser assumida como lacuna.

**Credenciais (ALTA).** `_obter_credenciais_env` tem assinaturas diferentes:
- Desktop: `SAP_{AMBIENTE}_USER` e `SAP_{AMBIENTE}_PASSWD`, apenas.
- Web: prefixo resolvido por `_resolver_prefixo_ambiente`, com fallbacks para `SAP_USER`, `SAP_LANGUAGE`, `SAP_PASSWORD_<SISTEMA>CLNT<cliente>` e `SAP_PASSWORD`.

O mesmo `.env` pode resolver credenciais diferentes no terminal e na web. Não centralizar antes de decidir a regra única.

**Defaults de criação de request (ALTA, escrita SAP).** Em `perguntar_opcao_request`, a web chama `_criar_nova_request_no_sap(..., tipo=payload.get("request_type") or "1", desc=payload.get("request_desc") or "REQUEST CRIADA VIA WEB", interactive=False)`. Se o payload omitir o tipo, cria uma request de customizing sem escolha explícita. O terminal pergunta sempre. Esta é uma escrita no SAP. Recomendação: exigir `request_type` e `request_desc` no contrato web (`CockpitRunRequest`) e rejeitar a ausência. **Mudança de comportamento; precisa de aprovação.**

**Seleção de ambiente (MÉDIA).** Desktop lê `SAP_ENVIRONMENT`/`SAP_ENV` e depois pede stdin. Web lê `payload['ambiente']` e `sys.argv`, sem variáveis de ambiente. Não é um defeito, mas o adaptador de terminal deve manter as duas fontes.

**Erro de scripting inativo (MÉDIA).** Desktop termina com `sys.exit(1)`. Web lança `SapCockpitError` quando `interactive=False`. A função tem o mesmo nome e papéis diferentes; o Core deve receber uma política de erro, não decidir sozinho.

**Tratamento de interrupção (BAIXA).** Desktop trata `KeyboardInterrupt`/`EOFError` em `selecionar_pasta_processo`. Web não. Irrelevante no worker, mas o adaptador de terminal deve preservar o comportamento.

**Texto (BAIXA).** A web usa mensagens sem acentos. Não afeta a lógica, mas altera o texto visível nas duas interfaces; decidir um padrão antes da migração.

**Duplicação não detetada por nome.** `_build_exec_kwargs` (web) replica a lógica inline de `executar_processo` do desktop: aba do Excel, popup de ficheiro, `request_ctx` e passagem de parâmetros do payload. Não aparece na matriz porque o nome não coincide. É a principal candidata à centralização de regra de negócio.

### 7.5 Funções específicas (sem par direto)

- Só no desktop: `_obter_conexao_env` e `_validar_ligacao_rfc`. Validação RFC no arranque; a web não a tem. Decidir se o Core a inclui.
- Só na web: `_build_exec_kwargs`, `_executar_um_script`, `_load_process_module`, `_obter_nomes_logon_candidatos`, `_raise_or_exit`, `_read_subprocess_web_config`, `_reportar_metadata_sap`, `_resolve_processo_path`, `_resolve_subprocesso_path`, `_resolver_prefixo_ambiente`, `_validar_argumentos_exec`, `main_terminal`, `obter_sessao_sap`, `read_sbar_status`, `run_sap_cockpit`.
- `obter_sessao_sap` (web) corresponde ao código de topo do desktop (linhas ~1005–1159), que não está numa função. Por isso o desktop executa a abertura de sessão ao ser importado.

### 7.6 Candidatas à centralização

**Grupo A: funções puras ou quase puras (risco baixo, primeiro lote).**
- Parsing de ambiente: `_parse_env_line`, `_carregar_dotenv_manual`, `_carregar_dotenv`, `_resolver_prefixo_ambiente`.
- Caminhos: `_resolve_processo_path`, `_resolve_subprocesso_path`.
- Classificação de erros: `_is_scripting_disabled_error`.
- Introspeção de subprocessos: `_analisar_exec_signature`, `_validar_argumentos_exec`.

**Grupo B: helpers SAP GUI sem estado (risco médio; exigem fakes de COM).**
- `_safe_find`, `_press`, `_send_vkey`, `_set_text`, `_ensure_se10`, `_select_radio_if_exists`, `_try_get_obj_text`, `_get_created_request_number`, `_extract_request_number_from_text`, `_extract_request_from_known_ids`, `_extract_request_by_scanning_usr`, `_sleep`, `_resetar_env_request`.
- Mover como funções que recebem `session`; nenhuma lê globais além de constantes.

**Grupo C: sessão e login (risco alto).**
- `_is_sap_logado`, `_tem_alguma_sessao_ativa`, `_encontrar_sessao_do_sistema`, `_aguardar_login`, `_dismiss_popup_if_any`, `_log_alerta_rz11`, `_log_scripting_status_apenas_quando_logado`, `_erro_scripting_inativo` (com política de erro injetada).

**Não centralizar nesta fase (adaptadores):**
- `selecionar_ambiente`, `selecionar_pasta_processo`, `perguntar_opcao_request`, `_criar_nova_request_no_sap`, `executar_processo`, `main_terminal`, `run_sap_cockpit` e `_build_exec_kwargs` (após decisão sobre defaults e credenciais).
- `_obter_credenciais_env`: decisão de produto primeiro (ver 7.4).

### 7.7 Proposta de arquitetura

```
sap_cockpit/core/             funções puras: ambiente, .env, caminhos, classificação de erros,
                              introspeção de executar(), contratos (CockpitRunRequest/Result)
sap_cockpit/core/gui.py       helpers SAP GUI que recebem session (grupo B)
sap_cockpit/core/session.py   sessão e login (grupo C), com política de erro injetada
sap_cockpit/cli/              adaptador de terminal: input(), sys.exit, KeyboardInterrupt
sap_script_web_cockpit_v2/    adaptador web: payload, SapCockpitError, sem prompts
```

- Sem efeitos no import: nenhum módulo do Core executa código de topo. O desktop e a web apenas importam.
- Sem dependências novas. `inspect`, `ast`, `dataclasses` e `pywin32` já são usados.
- Compatibilidade: `sap_cockpit_web_ready` continua a expor `run_sap_cockpit` e `main_terminal` como reexportações, para não quebrar `SAP_COCKPIT_MODULE` nem o worker.

### 7.8 Ordem recomendada de migração

1. **Decisões do utilizador (bloqueantes):** regra única de credenciais; `request_type` e `request_desc` obrigatórios na web.
2. **Correção de segurança no desktop:** remover os fallbacks `TypeError` de `SAP Cockpit.py` (linhas 676 e 860–862), reutilizando `_validar_argumentos_exec` movido para o Core. Pequeno e coberto pelo teste de re-execução já existente.
3. **Testes de equivalência (golden) antes de mover qualquer função:** para cada função do grupo A, comparar o resultado do desktop e da web com as mesmas entradas.
4. **Grupo A:** mover para o Core, com desktop e web a importar; testes verdes.
5. **Grupo B:** helpers GUI com fake de COM; comparar chamadas feitas ao objeto `session`.
6. **Desktop sem código de topo:** envolver a abertura de sessão em função e proteger com `__main__`. Só depois é possível testar o import do desktop.
7. **Grupo C:** sessão e login, com política de erro injetada.
8. **Adaptadores:** terminal e web sobre o Core; remover duplicados já substituídos.
9. **Remoção final:** só depois de equivalência comprovada (requisito da Fase 8).

### 7.9 Testes necessários antes de cada alteração

| Alteração | Teste mínimo |
|---|---|
| Fallback TypeError no desktop | Mesmo teste de re-execução da Fase 2 aplicado a `executar_processo` do desktop |
| Credenciais | Tabela de combinações de `.env` (`SAP_DEV_USER`, `SAP_USER`, `SAP_PASSWORD`, `SAP_PASSWORD_<SIS>CLNT<cli>`, `SAP_DEV_LANG`) com saída esperada para desktop e web |
| Defaults de request | Payload sem `request_type` deve falhar; com valor explícito deve criar com esse tipo (fake de session, sem SAP) |
| Grupo A | Golden tests: mesma entrada, mesma saída nas duas origens, antes e depois da migração |
| Grupo B | Fake de `session` que regista chamadas; comparar sequência de `Press`, `Text` e `sendVKey` |
| Grupo C | Fake de `application` com `Children` e `Info`; comparar sessão escolhida e mensagens |
| Desktop sem código de topo | Import não chama `GetObject` nem `Dispatch` (já existe na web) |
| Qualquer alteração | `run_all.py` com venv do cockpit e `test_sap_cockpit_core` com venv do worker |

### 7.10 Limitações desta auditoria

- Análise estrutural: não executa as funções. Equivalência de runtime só se prova com os testes da secção 7.9.
- Não comparei os corpos linha a linha de cada função. A classificação "Idêntica" assenta no AST e a "Semelhante" no `difflib`, com revisão manual dos casos marcados.
- Não verifiquei o comportamento de `_obter_credenciais_env` com um `.env` real; não era necessário e evita tocar em credenciais.
- Não avaliei o comportamento das funções de SAP GUI contra um SAP real.


### 7.11 Resultado da Fase 3b

**Implementado:**

- Reexecução após `TypeError` removida nos dois caminhos do `SAP Cockpit.py`: `executar_processo` (antes linhas 860–862) e `perguntar_opcao_request` (antes linha 676, `listar_requests`).
- Validação de assinatura com `inspect.signature().bind()` antes de cada `executar()`, em `sap_cockpit_core.validar_argumentos_exec`. Desktop e web usam a mesma função.
- Funções puras centralizadas em `sap_cockpit_core.py` (raiz do projeto), após prova de equivalência:
  - `parse_env_line` (desktop `_parse_env_line`, web `_parse_env_line`).
  - `is_scripting_disabled_error` (desktop e web).
- Ambos os ficheiros passaram a importar estas funções com os mesmos nomes locais (`_parse_env_line`, `_is_scripting_disabled_error`), para não alterar os pontos de chamada.

**Prova de equivalência:** fixture `tests/fixtures/sap_cockpit_golden_puras.json`, gerada a partir de `HEAD` (`b34330e`) antes da centralização. 22 casos para `parse_env_line` e 17 para `is_scripting_disabled_error`. Desktop e web deram o mesmo resultado em todos os casos; a implementação central também. A cobertura limita-se a este corpus.

**Não implementado (dependências identificadas, decisão pendente):**

1. **Regra de credenciais (7.4).** A regra aprovada (`SAP_{AMB}_USER` e `SAP_{AMB}_PASSWD`, sem fallbacks automáticos) não foi aplicada. Verificação por nomes de variáveis (sem valores):
   - O `.env` do projeto tem `SAP_{DEV,QAD,PRD}_USER`, `_PASSWD` e `_LANG`. Não tem `SAP_USER`, `SAP_PASSWORD`, `SAP_LANGUAGE` nem `SAP_PASSWORD_<SIS>CLNT<cli>`.
   - `Processos/transportar_request.py` (login STMS remoto) depende de `SAP_USER`, `SAP_LANGUAGE` e `SAP_PASSWORD_<SISTEMA>CLNT<cliente>`.
   - `scripts/inspect_dmeex.py` depende de `SAP_USER`, `SAP_LANG` e `SAP_LANGUAGE`.
   - `.env.example` e `.env.sap_agent.example` documentam `SAP_USER`, `SAP_LANGUAGE` e `SAP_PASSWORD_S4?CLNT100`.
   - Conclusão: o fallback da web é a mesma convenção usada por outros fluxos. Removê-lo só no cockpit web criaria duas regras no mesmo projeto. Decisão necessária: (a) remover os fallbacks também em `transportar_request.py` e `inspect_dmeex.py`, ou (b) manter a convenção e documentá-la como regra única.

2. **Contrato de requests SE10 (7.4).** `request_type` e `description` não foram tornados obrigatórios:
   - `web_api/main.py` (linhas 3741–3756) define `request_option` default `"4"` e `request_type` default `"1"` quando o formulário omite o campo.
   - `web_api/templates/index.html` (linha 2062): o seletor `request-type-select` começa desativado. Na prática o formulário envia sempre `"1"`.
   - Tornar o campo obrigatório quebra este consumidor. Não foi verificado se o formulário envia `request_desc`.
   - Decisão necessária: (a) o formulário passa a enviar `request_type` e `request_desc` explícitos, e só depois a API rejeita a ausência; ou (b) manter o default no formulário e aceitar que ele é explícito na UI.

**Dependências verificadas, sem alteração:** `obter_sessao_sap`, `_obter_credenciais_env` e `selecionar_ambiente` não foram tocados.

**Testes:**

- `run_all.py` (venv do cockpit): 95 executados, 77 OK, 18 saltados (dependências do worker em falta), `RESULTADO: OK`. Antes da Fase 3b: 82 executados (65 + 17). Novos: 13 testes em `tests/test_sap_cockpit_fase3b.py`.
- `worker\.venv\...\python.exe -m unittest tests.test_sap_cockpit_core tests.test_sap_cockpit_fase3b`: 30 OK. Inclui o teste da web (`ExecucaoWebUmaVezTest`), que o venv do cockpit salta.
- Regressão contra o desktop de `HEAD`: os testes `DesktopSemRetryTest` e `DesktopListarRequestsSemRetryTest` falham (2 ERROR e 1 FAIL). Confirma que apanham o bug. Ficheiro temporário removido após a verificação.

**Riscos residuais:**

- Desktop continua a executar código de topo ao ser importado (Fase 4). Os testes do desktop leem o código por AST e não o importam.
- Um script que levanta `TypeError` interno após escrever no SAP agora reporta erro e não é repetido. A operação parcial continua a ser responsabilidade do script; não há transação.
- As credenciais e os requests continuam com as regras antigas até decisão.
- `sap_cockpit_core.py` está na raiz do projeto. A imagem Docker da web copia só `web_api` e `salsa_agent`, por isso o módulo não está lá; o worker Windows usa a raiz por `SAP_SCRIPT_PROJECT_DIR`. Se a web passar a importar o core, é preciso acrescentá-lo ao `Dockerfile`.
- Nenhum deploy foi feito. O worker Windows precisa de reinício quando for autorizado.
