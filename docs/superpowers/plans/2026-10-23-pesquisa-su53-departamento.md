# Pesquisa SU53 por Departamento Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Adicionar pesquisa SU53 somente leitura para todos os utilizadores de um departamento selecionado.

**Architecture:** Helpers leem `CONTROLO` e `Proposta Ativa`. Um pesquisador departamental reutiliza consulta, agrupamento e mapeamento existentes, agrega resultados e devolve utilizadores com erros para seleção individual.

**Tech Stack:** Python, pandas, openpyxl, pytest.

**Spec:** `docs/superpowers/specs/2026-10-23-pesquisa-su53-departamento.md`

## Global Constraints

- Não adicionar dependências.
- Não modificar o workbook.
- Não efetuar alterações SAP no lote departamental.
- Preservar fluxo individual e argumentos CLI existentes.

## Review Focus

- Departamento sem utilizadores termina sem RFC.
- Comparação de departamento ignora diferenças de maiúsculas e espaços.
- Uma falha RFC não interrompe os utilizadores seguintes.
- Somente utilizadores com erros ficam disponíveis para fluxo individual.
- O lote não chama rotinas de atualização SAP.

---

### Task 1: Fontes departamentais do Excel

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Produces: `listar_departamentos_controlo(caminho_excel)` e `listar_utilizadores_departamento(caminho_excel, departamento)`.

- [x] Escrever testes falhando para ordem da `CONTROLO`, contagens e filtro departamental.
- [x] Executar testes e confirmar RED.
- [x] Implementar os dois leitores sem alterar o workbook.
- [x] Executar testes e confirmar GREEN.

### Task 2: Pesquisa e resumo somente leitura

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: lista de utilizadores da Task 1 e helpers SU53 existentes.
- Produces: `pesquisar_departamento_su53(...)` com resultados agregados.

- [x] Escrever teste falhando para sucesso, ausência de erros e falha RFC.
- [x] Executar teste e confirmar RED.
- [x] Implementar pesquisa sequencial e resumo.
- [x] Executar testes e confirmar GREEN.

### Task 3: Menu inicial e seleção individual posterior

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: helpers das Tasks 1 e 2.
- Produces: modo `[1] Utilizador` e `[2] Departamento`, com retorno opcional ao fluxo individual.

- [x] Escrever teste do fluxo interativo com dependências RFC simuladas.
- [x] Executar teste e confirmar RED.
- [x] Integrar o menu no `main()` sem duplicar o fluxo individual.
- [x] Executar testes focados, compilação e `graphify update .`.
