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
| 2. Extrair Core partilhado | Pendente |
| 3. Migrar executor e sessões | Pendente |
| 4. Adaptar o terminal | Pendente |
| 5. Integrar o worker Web | Pendente |
| 6. Reforçar segurança e rastreabilidade | Parcial: credenciais em README/migração removidas |
| 7. Testes de integração e regressão | Pendente (requer ambiente Windows com SAP GUI) |
| 8. Remover duplicações e atualizar documentação | Pendente |

Parar nesta fase é intencional: as fases 2 a 5 exigem decidir como tratar as alterações não commitadas (ponto 3.2) e validar o comportamento do worker Windows.
