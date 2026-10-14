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
- **Fluxo Geral:** Reconciliação Integrada de 8 Etapas da Tarefa 1 (com Pre-Check interativo de atualização a partir do OneDrive corporativo antes de qualquer leitura ou etapa): `[1/8]` Matrizes Departamentais × Proposta $\rightarrow$ `[2/8]` Proposta × SAP TSTC $\rightarrow$ `[3/8]` Proposta Ativa $\rightarrow$ `[4/8]` Catálogo Vivo $\rightarrow$ `[5/8]` PFCG_AUTHORITY (reconciliação estrita com PFCG_CREATE e PFCG_COMPOSTA em dry-run) $\rightarrow$ `[6/8]` Validar Funções Individuais (valida e classifica estritamente roles esperadas como SINGLE [89 esperadas]: SINGLE_OK, SINGLE_AUSENTE, SINGLE_TIPO_DIVERGENTE; auditoria segura de utilizadores, tcodes, autorizações e relações mantendo bloqueadas para revisão manual sem deleção cega) $\rightarrow$ `[7/8]` Validar Funções Compostas (valida estritamente roles esperadas como COMPOSITE [35 esperadas]: COMPOSITE_OK, COMPOSITE_AUSENTE, COMPOSITE_TIPO_DIVERGENTE; correção sob confirmação explícita com delete da Single incorreta via BPC_DELETE_SINGLE_ROLE com nome truncado a 30 caracteres + commit + readback ausência + recriação via PRGN_RFC_CREATE_AGR_MULTIPLE com COLLECTIVE_AGR='X' + commit + readback COLL_AGR='X') $\rightarrow$ `[8/8]` CUA_ADICIONAR (preparação apenas em Excel, idempotente, compactada no modo normal, sem escrita SAP).
- **Frontend:** `cockpit.agent.js` (`projeto-perfil-execucao`, `projeto-perfil-departamento`, `projeto-perfil-utilizador`, `projeto-perfil-pesquisa`, `projeto-perfil-corrigir`, `projeto-perfil-su53`, `projeto-perfil-atribuir-su53`)
- **API:** `POST /api/salsa-it-agent/projeto-perfil/{action}`, `GET /api/salsa-it-agent/projeto-perfil/{action}/{job_id}`, `POST /api/jobs` (tasks `projeto_perfil_execucao`, `projeto_perfil_departamento`, `projeto_perfil_utilizador`, `projeto_perfil_pesquisa`, `projeto_perfil_corrigir`, `projeto_perfil_su53`, `projeto_perfil_atribuir_su53`)
- **Worker:** `worker/sap_tasks.py` (`_run_projeto_perfil_task`)
- **Serviço / Scripts Operacionais / Reconciliação:**
  - `sap_rfc/status_logger.py` (Módulo Centralizador de Formatação e Cores de Status no Terminal: linhas completas formatadas em verde para sucesso, amarelo para aviso/pendência, vermelho para erro)
  - `sap_rfc/pfcg_authority_reconcile_service.py` (Etapa `[5/8]`: Reconciliação PFCG_AUTHORITY com PFCG_CREATE e PFCG_COMPOSTA: classificação estrita por coluna onde `AGR_NAME_COMPOSTA` $\rightarrow$ `COMPOSITE` e `AGR_NAME` $\rightarrow$ `SINGLE`; detecção de `AUTHORITY_COMPOSITE_SEM_CHILD`, deduplicação idempotente, detecção de `CONFLITO_TIPO_AUTHORITY`, execução em dry-run `SAP_WRITE=0` / `EXCEL_WRITE=0` e escrita segura com backup timestampado em `output/`)
  - `sap_rfc/role_existence_validator_service.py` (Serviço Oficial de Validação de Existência e Tipos de Roles no SAP: separação estrita onde `TIPO_ESPERADO` determina qual etapa trata a role; Etapa `[6/8]` `executar_etapa_validacao_funcoes_individuais` para Single Roles esperadas e `auditar_single_role_divergente`; Etapa `[7/8]` `executar_etapa_validacao_funcoes_compostas` para Composite Roles esperadas e `corrigir_composite_role_divergente_com_readback`; guards isolados `guard_ro` e `guard_write`; semântica formal de `UNKNOWN` quando ambiente está offline/indisponível sem gerar falso MISSING; bloqueio absoluto de `PRGN_RFC_CREATE_ACTIVITY_GROUP` para Composite Roles; aborto automático de operações dependentes se houver `PASSWORD_MUST_BE_CHANGED`)
  - `sap_rfc/projeto_perfil_service.py` (`atribuir_objetos_e_funcoes_su53`, diagnóstico SU53, reconciliação), `sap_rfc/projeto_perfil_cli.py` (`--action atribuir_su53`, `--action su53`)
  - `sap_rfc/pfcg_composta_sync_service.py`, `sap_rfc/pfcg_composta_sync_cli.py` (Serviço RFC modular: fonte única de verdade `obter_pendencias_pfcg_composta`; pré-validação READ-ONLY `prevalidar_pfcg_composta` executada apenas após validação de compostas; relações bloqueadas se parent for `PARENT_TIPO_DIVERGENTE` ou child for `CHILD_TIPO_DIVERGENTE`; reconciliação de já existentes no Excel sem escrita SAP; delta real executável e readback físico obrigatório pós-escrita)
  - `scripts/reconciliar_proposta_ativa_com_prd.py` (Integração encadeada Fases 1 a 4 com flag `--real` para gravação no Excel)
  - `Processos/Projeto Autorizações/Projeto Perfil.py` (`executar_atualizacao_integrada_excel` [Padronizado para 8 etapas oficiais `/8`: `[1/8]` a `[8/8]`, com Pre-Check de atualização segura OneDrive -> Local com backup em `output/`], `verificar_e_perguntar_pendencias` [Processo operacional com verificação prévia de logon RFC com aborto imediato sob `PASSWORD_MUST_BE_CHANGED`, despachante com confirmações e readbacks físicos obrigatórios], `preparar_cua_adicionar` [Etapa `[8/8]`], `executar_atribuicoes_cua_sap`, `sincronizar_departamento_prd_qad_rfc`, `executar_correcao_sincronizacao_posterior`)
  - `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- **Testes:** `tests/test_pfcg_authority_reconcile_service.py`, `tests/test_role_existence_validator_service.py`, `tests/test_auditoria_matrizes_vs_proposta.py`, `tests/test_pfcg_composta_sync_service.py`, `tests/test_reconciliar_proposta_ativa_com_prd.py`, `tests/test_preparacao_e_execucao_cua.py`, `tests/test_despachante_interativo_processos.py`, `tests/test_projeto_perfil_agent.py`

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

#### 4.1 Agente Salsa IT — Análise Passiva de Acessos (Fase 1)
- **Fluxo:** Leitura passiva, pré-classificação semântica e deteção de dependências operacionais de tickets Jira atribuídos a Clayton Lopes em `CA Acessos` (Zero operações SAP, Zero escrita Jira).
- **Frontend:** `cockpit.agent.js` (`openSalsaDiagnosticModal`, `triggerSalsaManualReanalyze`, `renderSalsaModalData`), `templates/index.html` (modal `salsa-agent-diagnostic-modal` com badge de status operacional, solicitante, dependência e botão contextual `Agente IA` na tabela de tickets)
- **API:**
  - `GET /api/salsa-it-agent/tickets/{ticket_key}/eligibility`
  - `POST /api/salsa-it-agent/tickets/{ticket_key}/analyze`
  - `GET /api/salsa-it-agent/tickets/{ticket_key}/analysis`
  - `GET /api/salsa-it-agent/tickets/{ticket_key}/history`
  - `GET /api/salsa-it-agent/analyses`
- **Módulos do Agente:**
  - `salsa_agent/models.py`: Modelos tipados (`ClassificationCode`, `AnalysisLevel`, `AnalysisStatus`, `OperationalStatus`, `ExtractedEntities`, `AnalysisResult`).
  - `salsa_agent/classifier.py`: Nível 1 — Classificador determinístico v1.1.0 com suporte a formas clíticas e coloquiais ("criar-me um user", "acesso sap pela primeira vez"), deteção de negações e defesa contra prompt injection.
  - `salsa_agent/extractor.py`: Extrator estrito de entidades (utilizador SAP, utilizador de referência, solicitante, ambientes, empresa).
  - `salsa_agent/ai_provider.py`: Nível 2 — Abstração de IA desacoplada com modo passivo `DisabledAIProvider` (`AI_NOT_CONFIGURED`) e `MockAIProvider`.
  - `salsa_agent/orchestrator.py`: Orquestrador com validação de elegibilidade no backend, deteção de dependências/bloqueios via comentários (`_detect_operational_status`), e idempotência via SHA-256 de conteúdo e versão de regras.
- **Persistência SQLite:** `web_api/store.py` (`salsa_it_ticket_analysis`, `salsa_it_ticket_analysis_history` com colunas `operational_status` e `identified_dependency`).
- **Testes:** `tests/test_salsa_it_phase1.py` (20 testes específicos), `tests/run_all.py` (61 testes totais aprovados).


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
