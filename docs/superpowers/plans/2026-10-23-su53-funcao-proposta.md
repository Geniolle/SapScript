# SU53 com Funcao da Proposta Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Incluir aplicacao e funcao da sheet `Proposta` em cada objeto do diagnostico SU53 no terminal e no JSON.

**Architecture:** Preservar a leitura RFC atual. Enriquecer os objetos agrupados com a aplicacao do erro mais recente e reutilizar `buscar_funcao_individual_para_transacao` para resolver a funcao e as linhas da sheet `Proposta`.

**Tech Stack:** Python 3.12, openpyxl, pyrfc, pytest.

**Spec:** `docs/superpowers/specs/2026-10-23-su53-funcao-proposta-design.md`

## Global Constraints

- Somente leitura no SAP e no Excel.
- Usar `P_TCODE`; quando vazio, usar `NAME`.
- Nao adicionar dependencias.
- Preservar contratos atuais do diagnostico SU53.

## Review Focus

- `P_TCODE` vazio com `NAME` preenchido deve resolver a funcao.
- `P_TCODE` preenchido deve ter prioridade sobre `NAME`.
- Transacao ausente na `Proposta` deve gerar estado explicito, sem excecao.
- Multiplas funcoes para a mesma transacao devem ser preservadas.
- Objetos antigos sem campos de aplicacao devem continuar validos.

---

### Task 1: Enriquecer objetos SU53 com aplicacao

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: erros normalizados de `consultar_erros_su53_rfc`.
- Produces: objetos agrupados com `tcode_app`, `nome_servico` e `aplicacao` do registo mais recente.

- [ ] Escrever testes para prioridade de `P_TCODE`, fallback para `NAME` e ausência dos dois campos.
- [ ] Executar os testes e confirmar falha pelo campo `aplicacao` inexistente.
- [ ] Alterar `agrupar_erros_para_analise` no menor ponto responsável.
- [ ] Executar os testes e confirmar aprovação.

### Task 2: Resolver função da Proposta no serviço

**Files:**
- Modify: `sap_rfc/projeto_perfil_service.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: `aplicacao` dos objetos agrupados e `buscar_funcao_individual_para_transacao`.
- Produces: cada item de `objetos` com `aplicacao`, `funcoes_proposta` e primeira `funcao_proposta`.

- [ ] Escrever teste de serviço com `FAGLB03`, `Z_ACCOUNTING_REPORT`, descrição e linhas 24/35.
- [ ] Executar e confirmar falha pela ausência dos campos.
- [ ] Implementar enriquecimento sem alterar o fluxo de atribuição.
- [ ] Executar teste focado e suite do projeto.

### Task 3: Apresentar função no terminal e manter Cockpit/API

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: objeto enriquecido e caminho do Excel.
- Produces: bloco textual solicitado e JSON já enriquecido para Cockpit/API.

- [ ] Escrever teste do bloco textual com aplicação, objeto, atividade, horário, função, descrição e linhas.
- [ ] Executar e confirmar falha pelo relatório inexistente.
- [ ] Implementar um formatador puro e chamá-lo no fluxo interativo.
- [ ] Executar teste focado, suite e `graphify update .`.

### Verificação final

- [ ] Executar `\.venv-rfc\Scripts\python.exe -m pytest tests/test_projeto_perfil_su53_report.py -q`.
- [ ] Executar `\.venv-rfc\Scripts\python.exe -m pytest -q`.
- [ ] Executar diagnóstico somente leitura para `S80002051` e confirmar `FAGLB03` / `Z_ACCOUNTING_REPORT`.
- [ ] Rever `git diff` apenas nos ficheiros deste plano.
