---
name: sap_authorization_profile_analysis
description: Regras, padrões de arquitetura e procedimentos para o processo de Análise de Perfis de Autorização SAP via CUA GUI (USLA04 em SE16 ALV Grid) e RFC. Utilizar quando o utilizador solicitar análises de autorizações, criação ou refatoração de rotinas de leitura de tabelas CUA ou ajustes na interface do Assistente de Autorizações.
---

# Skill: Análise de Perfis de Autorização SAP (CUA & RFC)

Esta skill documenta as regras de negócio, padrões de arquitetura, métodos de extração SAP GUI/ALV e diretrizes de UI para o processo de **Análise de Perfis de Autorização SAP**.

---

## 1. Arquitetura do Protocolo CUA

- **Método de Execução:** Toda a análise de autorizações CUA é realizada estritamente via **SAP GUI Scripting** (SE16), **NÃO via RFC**.
- **Sistema Central CUA:** Conexão direta ao ambiente `SPACLNT001` (Cliente Central CUA `SPA` / `001`).
- **Transação Autorizada:** Apenas a transação **`SE16`** é utilizada no ambiente CUA (a `SE16N` não existe / não deve ser utilizada).

---

## 2. Extração da Tabela CUA `USLA04` & ALV Grid

- **Tabela Principal:** `USLA04` (Atribuição de funções/roles a utilizadores nos subsistemas CUA).
- **Campos de Filtro Obrigatórios:**
  - `BNAME`: Utilizador SAP a analisar (ex: `S80001870`).
  - `SUBSYSTEM`: Subsistema/Ambiente recetor (ex: `S4DCLNT100`, `S4PCLNT100`).

### Leitura Nativa de ALV Grid (`GuiGridView`)
Quando a `SE16` exibe os resultados num controlo **ALV Grid**, a extração utiliza a função `_read_alv_grid`:
1. Obter a contagem total de linhas: `grid.RowCount`.
2. Iterar linha a linha (`0` a `RowCount - 1`) lendo os valores por coluna via `grid.GetCellValue(row, col)`.
3. Mapeamento de colunas técnicas:
   - `BNAME` $\rightarrow$ Utilizador
   - `SUBSYSTEM` $\rightarrow$ Subsistema Alvo
   - `AGR_NAME` $\rightarrow$ Nome da Role / Função PFCG
   - `FROM_DAT` $\rightarrow$ Data de Início de Validade (`YYYY-MM-DD` / `DD.MM.YYYY`)
   - `TO_DAT` $\rightarrow$ Data de Fim de Validade (`31.12.9999` = Ativa)
   - `ORG_FLAG` $\rightarrow$ Origem da Atribuição (`""` = Direta, `"C"` = Role Composta, `"X"` = Organização RH)

---

## 3. Estados de Job & Polling no Frontend

- **Estados de Sucesso Aceites:**
  - `succeeded`: Análise concluída com roles ativas encontradas.
  - `succeeded_with_warnings`: Análise concluída com sucesso (ex: 0 roles ativas ou avisos de validade).
  - Ambos os estados DEVEM ser aceites nos loops de polling (`job.state === 'succeeded' || job.state === 'succeeded_with_warnings'`).

- **Validação de Tabelas por Tipo de Fluxo:**
  - **Fluxo CUA (`isCuaFlow`):** Valida apenas a tabela principal **`USLA04`**. Não exigir `USL04`, `AGR_USERS` ou `USZBVSYS`.
  - **Fluxo DEV RFC (`isDevFlow`):** Valida `AGR_USERS` e `AGR_TCODES`.
  - **Fluxo Dados Mestre (`isMasterData`):** Valida `USR02`, `USR21`, `USR04` e `AGR_USERS`.

---

## 4. Padrões visuais e de UI do Chat (`authorization_assistant.js`)

- **Cartões Soltos (`Cards Soltos`):** As pílulas de sugestões de rotinas e opções de processos devem ser renderizadas como **cartões soltos** diretamente abaixo das bolhas de texto do assistente, nunca encapsulados na mesma bolha de texto.
- **Estilo de Seleção dos Cartões (`.selected`):**
  - **Estado Normal:** Fundo branco limpo (`#ffffff`), border `#e2e8f0`.
  - **Estado Selecionado:** Fundo **cinzento claro suave (`#f1f5f9`)**, border `#cbd5e1`, box-shadow `rgba(203, 213, 225, 0.6)`.
- **Nomes de Sistemas em Negrito:** Códigos de sistema (**`S4D`**, **`S4P`**, **`S4Q`**, **`SPA`**) usam `font-weight: 700`.
- **Tipografia Compacta:** Pílulas e cartões usam fontes proporcionais (`0.82rem` / `0.86rem`) com `padding: 8px 12px`.
- **Shell Responsiva:** O contentor do chat (`.auth-chat-shell`) preenche **100% da largura útil disponível** (`width: 100%; max-width: 100%`).
