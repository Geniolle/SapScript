# Estrutura do Projeto SapScript

> Índice técnico e funcional navegável do projeto `SapScript`.
> **Regra incremental:** Os processos são documentados progressivamente à medida que são trabalhados. As secções não iniciadas permanecem como marcadores até haver trabalho ativo nelas.
>
> **Fluxo de navegação:** `Processo funcional → Frontend → API → Worker/Orquestração → Serviço → SAP/RFC → Testes`

---

## Agente Salsa IT (Web Cockpit & Orquestrador de Agentes SAP)

- **Localização Base:** `C:\workspace\SapScript\sap_script_web_cockpit_v2`
- **Componentes Centrais de Runtime:**
  - **Frontend UI:** `web_api/templates/index.html`, `web_api/static/js/cockpit.core.js`, `web_api/static/js/cockpit.agent.js`, `web_api/static/styles.css`
  - **API Backend (FastAPI):** `web_api/main.py`, `web_api/common.py`, `web_api/pfcg_common.py`, `web_api/user_create_common.py`, `web_api/store.py` (Base SQLite `db.sqlite3`)
  - **Worker (Execução SAP Windows):** `worker/worker.py`, `worker/sap_tasks.py` (despacho por dicionário `TASK_HANDLERS`), `worker/start_worker_auto.ps1`
  - **Rede de Segurança & Testes:** `tests/run_all.py`, `tests/js_smoke.py`, `tests/test_salsa_agent_routes.py`, `tests/test_worker_dispatch.py`, `tests/test_worker_reap.py`

### 1. Configurações

#### 1.1 Perfil de Autorização (PFCG)
- **Pesquisar / Analisar Função (Nome, Texto com curinga `Z*`, Transação, Objeto, Utilizador):**
  - **Frontend:** `cockpit.agent.js` (`pfcg-role-analyze`, `asiStartPfcgRoleAnalysis`, `asiStartPfcgPolling`, `asiPfcgRoleState`)
  - **API:** `POST /api/salsa-it-agent/pfcg/analyze`, `POST /api/salsa-it-agent/pfcg/search`, `POST /api/salsa-it-agent/pfcg/transactions/analyze`, `POST /api/salsa-it-agent/pfcg/users/analyze`, `POST /api/salsa-it-agent/pfcg/transaction/roles`, `POST /api/salsa-it-agent/pfcg/object/roles`, `POST /api/salsa-it-agent/pfcg/user/roles`
  - **Worker:** `worker/sap_tasks.py` (`pfcg_role_analysis`, `pfcg_role_search`, `pfcg_role_transactions_analysis`, `pfcg_role_users_analysis`, `pfcg_transaction_roles`, `pfcg_object_roles`, `pfcg_user_roles`)
  - **Serviço RFC:** `sap_rfc/pfcg_role_service.py` (`get_role_details`), `sap_rfc/pfcg_role_search_service.py` (`search_roles_by_pattern`), `sap_rfc/pfcg_role_transactions_service.py`, `sap_rfc/pfcg_role_users_service.py`, `sap_rfc/pfcg_transaction_roles_service.py`, `sap_rfc/pfcg_object_roles_service.py`, `sap_rfc/pfcg_user_roles_service.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`, `tools/test_worker_pfcg_role_analysis.py`

- **Criar Função Simples (Excel ou Individual):**
  - **Frontend:** `cockpit.agent.js` (`pfcg-create`, `pfcg-create-select-excel`, `pfcg-create-individual`, `asiPollPfcgIndividualPreview`, `asiPollPfcgIndividualConfirm`)
  - **API:** `POST /api/salsa-it-agent/pfcg/create/select-excel`, `POST /api/salsa-it-agent/pfcg/create/analyze`, `POST /api/salsa-it-agent/pfcg/create/rfc/preview`, `POST /api/salsa-it-agent/pfcg/create/rfc/confirm`
  - **Worker:** `worker/sap_tasks.py` (`select_excel_file`, `pfcg_create_excel_analysis`, `pfcg_role_create_preview`, `pfcg_role_create_rfc`)
  - **Serviço RFC:** `sap_rfc/pfcg_role_create_service.py` (`preview_role_creation`, `execute_role_creation_rfc`), `sap_rfc/pfcg_role_create_cli.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`, `tools/test_sap_pfcg_role_prd.py`

- **Função Composta (Excel ou Individual):**
  - **Frontend:** `cockpit.agent.js` (`pfcg-composta`, `pfcg-composta-analyze`, `pfcg-composta-select-excel`, `pfcg-composta-individual`)
  - **API:** `POST /api/salsa-it-agent/pfcg/composta/preview`, `POST /api/salsa-it-agent/pfcg/composta/confirm`
  - **Worker:** `worker/sap_tasks.py` (`pfcg_composta_create_preview`, `pfcg_composta_create`)
  - **Serviço RFC:** `sap_rfc/pfcg_role_create_service.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`

- **Eliminar Perfil (Individual ou Massa via Excel):**
  - **Frontend:** `cockpit.agent.js` (`pfcg-delete`, `pfcg-delete-select-excel`, `pfcg-delete-individual`, `asiPollPfcgDeletePreview`, `asiPollPfcgDeleteConfirm`)
  - **API:** `POST /api/salsa-it-agent/pfcg/delete/rfc/preview`, `POST /api/salsa-it-agent/pfcg/delete/rfc/confirm`, `POST /api/salsa-it-agent/pfcg/delete/rfc/bulk/preview`, `POST /api/salsa-it-agent/pfcg/delete/rfc/bulk/confirm`
  - **Worker:** `worker/sap_tasks.py` (`pfcg_role_delete_preview`, `pfcg_role_delete_rfc`, `pfcg_role_bulk_delete_preview`, `pfcg_role_bulk_delete_rfc`)
  - **Serviço RFC:** `sap_rfc/pfcg_role_delete_service.py`, `sap_rfc/pfcg_role_bulk_delete_cli.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`

- **Ordens de Transporte PFCG:**
  - **Frontend:** `cockpit.agent.js` (`pfcg-transport-search`, `asiPollPfcgTransportSearch`)
  - **API:** `POST /api/salsa-it-agent/pfcg/transport/search`
  - **Worker:** `worker/sap_tasks.py` (`pfcg_transport_search`)
  - **Serviço RFC:** `sap_rfc/pfcg_transport_service.py`, `sap_rfc/pfcg_gui_transport.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`

#### 1.2 Utilizador (SU01 / CUA)
- **Consulta e Pesquisa de Utilizador:**
  - **Frontend:** `cockpit.agent.js` (`user-search`, `user-data`)
  - **API:** `POST /api/salsa-it-agent/user/search`, `POST /api/salsa-it-agent/user/data`
  - **Worker:** `worker/sap_tasks.py` (`user_search`, `user_data`)
  - **Serviço RFC:** `sap_rfc/user_search_service.py`, `sap_rfc/user_data_service.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`

- **Criar Utilizador (Individual / RH / Preview / Confirm):**
  - **Frontend:** `cockpit.agent.js` (`user-individual-create`, `asiStartUserCreateWorkflow`)
  - **API:** `POST /api/salsa-it-agent/user/create/hr-lookup`, `POST /api/salsa-it-agent/user/create/rfc/preview`, `POST /api/salsa-it-agent/user/create/rfc/confirm`
  - **Worker:** `worker/sap_tasks.py` (`hr_lookup`, `user_create_preview`, `user_create_rfc`)
  - **Serviço RFC:** `sap_rfc/hr_lookup_service.py`, `sap_rfc/user_create_service.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`

- **Alteração de Senha & Desbloqueio de Utilizador:**
  - **Frontend:** `cockpit.agent.js` (`user-individual-password`, `user-individual-unlock`)
  - **API:** `POST /api/salsa-it-agent/user/password/change`, `POST /api/salsa-it-agent/user/unlock`
  - **Worker:** `worker/sap_tasks.py` (`user_change_password_rfc`, `user_unlock_rfc`)
  - **Serviço RFC:** `sap_rfc/user_create_service.py`, `sap_rfc/_rfc_common.py`, `sap_rfc/user_unlock_cli.py`, `sap_rfc/user_password_change_cli.py`
  - **Testes:** `tests/test_salsa_agent_routes.py`, `tests/test_user_password_lock_status.py` (em `SapScript/tests`)

- **CUA (Atribuição / Remoção):**
  - **Frontend:** `cockpit.agent.js` (`cua-add-excel`, `cua-add-individual`, `cua-rm-excel`, `cua-rm-individual`)
  - **API:** `POST /api/salsa-it-agent/cua/adicionar`, `POST /api/salsa-it-agent/cua/remover`
  - **Worker:** `worker/sap_tasks.py` (`sap_cockpit` com `Processos/Funções PFCG/H. CUA_ADICIONAR.py` / `I. CUA_REMOVER.py`)
  - **Testes:** `tests/test_salsa_agent_routes.py`

#### 1.3 OBYC
- **Analisar e Configurar Regras OBYC:**
  - **Frontend:** `cockpit.agent.js` (`obyc-configurar`, `obyc-analisar`, `obyc_excel_preview`, `obyc_excel_validate`)
  - **API:** `POST /api/salsa-it-agent/configuracoes/obyc`, `POST /api/salsa-it-agent/configuracoes/obyc/analisar`, `POST /api/salsa-it-agent/configuracoes/obyc/excel/preview`, `POST /api/salsa-it-agent/configuracoes/obyc/excel/validate`
  - **Worker:** `worker/sap_tasks.py` (`obyc_rfc_read_table`, `obyc_excel_preview`, `obyc_excel_validate`)
  - **Serviço RFC:** `sap_rfc/obyc_service.py`
  - **Testes:** `SapScript/tests/test_obyc_service.py`

#### 1.4 Conta Razão (GL Account)
- **Criar Conta Razão por Modelo:**
  - **Frontend:** `cockpit.agent.js` (`gl-account`, `asiBuildGlAccountEnvironmentAction`)
  - **API:** `POST /api/jobs` (task `gl_account_create_by_model`)
  - **Worker:** `worker/sap_tasks.py` (`gl_account_create_by_model` -> `_run_gl_account_create_by_model`)
  - **Serviço RFC:** `sap_rfc/gl_account_service.py`, `sap_rfc/gl_account_cli.py`

#### 1.5 Condição de Pagamento (OBB8 / ZTERM)
- **Pesquisa de Intervalo (RFC) & Criação por Cópia (SAP GUI Scripting):**
  - **Frontend:** `cockpit.agent.js` (`condicao-pagamento`, `condicao-pagamento-pesquisar-intervalo`, `condicao-pagamento-criar-copia`, `condicao-pagamento-criar-individual`)
  - **API:** `POST /api/salsa-it-agent/configuracoes/zterm/interval`, `POST /api/salsa-it-agent/configuracoes/zterm/copy`
  - **Worker:** `worker/sap_tasks.py` (`zterm_interval_analysis`, `zterm_copy_gui`)
  - **Serviço RFC & Automação GUI:** `sap_rfc/zterm_service.py` (`get_zterm_interval_analysis`), `Processos/zterm_copy_gui.py` (`copy_zterm_gui`), `Processos/criar_request.py` (`criar_nova_request_auto`)
  - **Documentação Técnica:** `docs/RELATORIO_TECNICO_OBB8_ZTERM.md`

---

### 2. Projetos

#### 2.1 Projeto Perfil de Autorização
- **Fluxo Geral:** Reconciliação Integrada de 7 Etapas (Tarefa 1): `[1/7]` Matrizes Departamentais × Proposta $\rightarrow$ `[2/7]` Proposta × TSTC $\rightarrow$ `[3/7]` Proposta Ativa $\rightarrow$ `[4/7]` Catálogo Vivo $\rightarrow$ `[5/7]` PFCG_COMPOSTA $\rightarrow$ `[6/7]` Preparar CUA_ADICIONAR (apenas Excel, idempotente, sem escrita SAP) $\rightarrow$ `[7/7]` Executar Sincronização SAP (exige confirmação explícita, BAPI + Commit + Readback obrigatório + atualização de CUA_ADICIONAR via ID por ambiente PRD/QAD independente).
- **Frontend:** `cockpit.agent.js` (`projeto-perfil-execucao`, `projeto-perfil-departamento`, `projeto-perfil-utilizador`, `projeto-perfil-pesquisa`, `projeto-perfil-corrigir`, `projeto-perfil-su53`)
- **API:** `POST /api/jobs` (tasks `projeto_perfil_execucao`, `projeto_perfil_departamento`, `projeto_perfil_utilizador`, `projeto_perfil_pesquisa`, `projeto_perfil_corrigir`, `projeto_perfil_su53`)
- **Worker:** `worker/sap_tasks.py` (`_run_projeto_perfil_task`)
- **Serviço / Scripts Operacionais / Reconciliação:**
  - `sap_rfc/projeto_perfil_service.py`, `sap_rfc/projeto_perfil_cli.py`
  - `sap_rfc/pfcg_composta_sync_service.py`, `sap_rfc/pfcg_composta_sync_cli.py` (Serviço RFC modular: execução exclusiva RFC via `.env` sem SAP GUI, pré-validação read-only em `AGR_AGRS`, processamento estritamente delta, readback físico em `AGR_AGRS` pós-escrita e retorno estruturado confiável)
  - `scripts/reconciliar_proposta_ativa_com_prd.py` (Integração encadeada Fases 1 a 4 com flag `--real` para gravação no Excel)
  - `Processos/Projeto Autorizações/Projeto Perfil.py` (`preparar_cua_adicionar`, `executar_atribuicoes_cua_sap`, `verificar_e_perguntar_pendencias` [Refatorado para execução interativa individual processo-a-processo com pré-validação read-only RFC e dupla confirmação para `CUA_REMOVE`], `executar_atualizacao_integrada_excel`, `executar_correcao_sincronizacao_posterior`)
  - `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- **Testes:** `tests/test_auditoria_matrizes_vs_proposta.py`, `tests/test_pfcg_composta_sync_service.py`, `tests/test_reconciliar_proposta_ativa_com_prd.py`, `tests/test_preparacao_e_execucao_cua.py`, `tests/test_despachante_interativo_processos.py`, `tests/test_projeto_perfil_agent.py`

---

### 3. Testes Unitários

#### 3.1 Criar Documento FI (FB01 / BAPI)
- **Fluxo:** Criação de documento financeiro de teste (Cliente D, Fornecedor K, Razão S) em DEV, QAD e PRD (modos Default ou Manual).
- **Frontend:** `cockpit.agent.js` (`testes-unitarios-criar-documento-fi-*`)
- **API:** `POST /api/fi/default-document`
- **Worker:** `worker/sap_tasks.py` (`fi_default_document` -> `worker/fi_default_document_job.py`)
- **Serviço RFC:** `sap_rfc/fi_document_service.py`, `sap_rfc/fi_payload_builder.py`, `sap_rfc/fi_config.py`
- **Testes:** `SapScript/tests/test_fi_default_document_job.py`, `SapScript/tests/test_fi_document_service.py`

#### 3.2 Executar F110 (Pagamentos Automáticos)
- **Fluxo:** Geração de proposta e execução de pagamento F110 em DEV, QAD e PRD (modos Default e Manual).
- **Frontend:** `cockpit.agent.js` (`testes-unitarios-executar-f110-*`)
- **API:** `POST /api/f110/proposal`, `POST /api/f110/payment`
- **Worker:** `worker/sap_tasks.py` (`f110_proposal` -> `worker/f110_proposal_job.py`, `f110_payment` -> `worker/f110_payment_job.py`)
- **Serviço RFC:** `sap_rfc/f110_service.py`
- **Testes:** `SapScript/tests/test_f110_proposal_job.py`, `SapScript/tests/test_f110_service.py`

---

### 4. Assistente de IA & Tickets Jira

- **Fluxo:** Leitura de tickets Jira, extração de sinais de erro SAP, diagnóstico via Gemini e chat interativo.
- **Frontend:** `cockpit.agent.js` (sidebar Jira, `asiStartChatWithGemini`)
- **API:** `POST /api/sap-agent/chat`, `GET /api/sap-agent/chat-job/{job_id}`, `POST /api/sap-agent/sap-query`
- **Worker / Integrações:** `web_api/jira_client.py`, `worker/sap_tasks.py` (`sap_agent_analysis`)

---

## Projeto Autorizações (Projeto Perfil Desktop)
*(será documentado quando trabalharmos nele)*

---

## Funções PFCG (Scripts Operacionais Desktop)
*(será documentado quando trabalharmos nele)*

---

## Conciliação OBYC / MM
*(será documentado quando trabalharmos nele)*

---

## Automação F110 / Pagamentos
*(será documentado quando trabalharmos nele)*

---

## DMEE / DMEEX Árvores de Formato
*(será documentado quando trabalharmos nele)*

---

## SAP Cockpit Desktop Original
*(será documentado quando trabalharmos nele)*
