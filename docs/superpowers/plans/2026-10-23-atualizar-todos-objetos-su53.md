# Atualizar Todos os Objetos SU53 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Adicionar execução sequencial de todos os objetos SU53 com uma confirmação e resumo final.

**Architecture:** Reutilizar `diagnosticar_objeto_autorizacao` com confirmação opcional. Um helper pequeno confirma o lote, chama a rotina para cada objeto e agrega resultados.

**Tech Stack:** Python, pytest, unittest.mock.

**Spec:** `docs/superpowers/specs/2026-10-23-atualizar-todos-objetos-su53.md`

## Global Constraints

- Não adicionar dependências.
- Não executar alterações SAP nos testes.
- Preservar seleção e processamento individual existentes.
- Processar ocorrências repetidas separadamente.

## Review Focus

- Lote cancelado não chama o executor.
- Falha num objeto não interrompe os seguintes.
- Resumo conta sucessos e falhas corretamente.
- Confirmação interna não se repete durante o lote.
- `[T]` só aparece quando existem vários objetos.

---

### Task 1: Execução em lote de objetos

**Files:**
- Modify: `Processos/Projeto Autorizações/Pesquisa_Erros_SU53.py`
- Test: `tests/test_projeto_perfil_su53_report.py`

**Interfaces:**
- Consumes: `diagnosticar_objeto_autorizacao(usuario, obj_info, caminho_excel, target_env)`.
- Produces: confirmação opcional na rotina existente e helper de lote com resultado agregado.

- [x] **Step 1: Escrever testes falhando**

Testar confirmação única, sequência completa, continuação após falha, cancelamento e visibilidade condicional de `[T]`.

- [x] **Step 2: Confirmar falha dos testes**

Run: `.\.venv-rfc\Scripts\python.exe -m pytest tests\test_projeto_perfil_su53_report.py -q`

- [x] **Step 3: Implementar mudança mínima**

Adicionar parâmetro para ignorar confirmação individual, helper de lote e ramo `[T]` no menu de objetos.

- [x] **Step 4: Confirmar testes focados**

Run: `.\.venv-rfc\Scripts\python.exe -m pytest tests\test_projeto_perfil_su53_report.py -q`

- [x] **Step 5: Atualizar Graphify**

Run: `graphify update .`
