---
name: sap_authorization_profile_analysis
description: Regras, padrões de arquitetura e procedimentos para o processo de Análise de Perfis de Autorização SAP via CUA GUI (USLA04 em SE16 ALV Grid), execução de sub-rotinas CUA/SU01 e validação RFC. Utilizar quando o utilizador solicitar análises de autorizações, criação ou refatoração de rotinas CUA (USLA04, CUA_ADICIONAR, CUA_ENDDATE, su01_reset_password.py) ou ajustes na interface do Assistente de Autorizações.
---

# Skill: Análise de Perfis de Autorização SAP (CUA, SU01 & RFC)

Esta skill documenta as regras de negócio, padrões de arquitetura, métodos de extração SAP GUI/ALV e diretrizes de UI para o processo de **Análise de Perfis de Autorização SAP** e **Automação de Dados de Utilizador CUA / SU01**.

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
- **Ordenação Alfabética Obrigatória (Regra Global A-Z):**
  Todas as listas de sugestões de rotinas e sub-rotinas (ex: menu inicial do assistente, menus de sub-processos) DEVEM ser ordenadas estritamente de **A a Z** pelo valor do texto (`val`). Quando nova rotina for criada ou adicionada, a ordenação alfabética DEVE ser mantida via `.sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }))`.
- **Nomes de Sistemas em Negrito:** Códigos de sistema (**`S4D`**, **`S4P`**, **`S4Q`**, **`SPA`**) usam `font-weight: 700`.
- **Tipografia Compacta:** Pílulas e cartões usam fontes proporcionais (`0.82rem` / `0.86rem`) com `padding: 8px 12px`.
- **Shell Responsiva:** O contentor do chat (`.auth-chat-shell`) preenche **100% da largura útil disponível** (`width: 100%; max-width: 100%`).

---

## 5. Fluxo de Ações Pós-Análise (Follow-up Actions & CUA_ENDDATE)

- **Pergunta Pós-Resumo de Análise:**
  Quando a pesquisa do processo de **Análise de Autorizações SAP** é concluída (ex: mensagem de resumo `"Análise de autorizações concluída via RFC"` ou CUA), o assistente/bot DEVE perguntar numa nova interação se o utilizador deseja **seguir com alguma ação** sobre o processo de Análise de Autorizações SAP.

- **Reaproveitamento de Dados da Análise Anterior:**
  Se o utilizador responder afirmativamente (ex.: *"sim"*, *"quero"*, *"seguir com ação"*, etc.):
  1. O bot deve apresentar novamente as opções de processos de Perfil de Autorização (ex.: `CUA_ENDDATE`, `CUA_REMOVE`, etc.).
  2. A nova ação deve reaproveitar automaticamente a lista de funções/roles selecionadas/exibidas na análise anterior.

- **Execução do `CUA_ENDDATE` sobre a Lista:**
  - O sistema executa o ajuste de data fim (`UPDATE_TO_DAT`) para cada uma das funções presentes na lista apresentada na análise anterior.
  - A data fim colocada por defeito pelo processo `CUA_ENDDATE` é a **data de ontem (hoje - 1 dia)** no formato SAP (`DD.MM.YYYY`).

- **Ações Diretas Abaixo da Lista Filtrada (CUA_ENDDATE / CUA_REMOVE):**
  Ao filtrar a lista por **`✅ Listar funções ativas`** ou **`❌ Listar funções expiradas`**, o assistente exibe a tabela filtrada e posiciona imediatamente abaixo os cartões de ação direta ordenados A-Z:
  1. **`📅 Delimitar data fim (CUA_ENDDATE)`**: Executa automaticamente o ajuste de validade no CUA para as funções presentes na tabela filtrada.
  2. **`➖ Remover funções (CUA_REMOVE)`**: Executa a remoção CUA para as funções presentes na tabela filtrada.
  3. **`🔄 Nova análise`**: Reinicia o fluxo de pesquisa.

---

## 6. Automação CUA Login (`su01_reset_password.py`)

- **Prioridade Absoluta das Variáveis CUA (.env):**
  A leitura de configuração em `load_config()` DEVE priorizar estritamente as variáveis dedicadas ao CUA (`SPACLNT001` / `CUA`) sobre as variáveis genéricas:
  - `connection_name` / `main_system_name` $\rightarrow$ `SAP_CONNECTION_SPACLNT001` (`CUA (PRD)`).
  - `client` $\rightarrow$ `SAP_CLIENT_SPACLNT001` (`001`).
  - `user` $\rightarrow$ `SAP_USER_CUA` / `SAP_USER`.
  - `password` $\rightarrow$ `SAP_PASSWORD_SPACLNT001` / `SAP_PASSWORD_CUA`.
  - `language` $\rightarrow$ `SAP_CUA_LANGUAGE` / `SAP_LANGUAGE_CUA` (`EN`).

- **Correspondência Flexível de Sistemas (`is_matching_system`):**
  Reconhece a equivalência entre o nome de ligação no SAP Logon (`CUA (PRD)` / `SPACLNT001`) e o SID técnico retornado pela sessão SAP GUI (`SPA`).

- **Validação de Destino em Modo RFC (`pyrfc`):**
  - **Step 6 (Execução CUA):** Alteração efectuada via GUI na transação `SU01` do CUA central.
  - **Step 7 (Validação no Alvo):** A confirmação da ativação da senha no sistema alvo (`S4PCLNT100`, `S4DCLNT100`, `S4QCLNT100`) é realizada **em modo RFC direto (`pyrfc.Connection`)**, sem abrir janelas visuais secundárias.
  - Respostas RFC como `PASSWORD_EXPIRED` ou `MUST_CHANGE` confirmam que a senha temporária foi gravada no alvo com sucesso.

---

## 7. Sub-rotinas do Processo "Dados de Utilizador" & Recolha de Campos CUA

- **Estrutura de Interação:**
  1. **Seleção Inicial:** O utilizador clica em **`👤 Dados de utilizador`**.
  2. **Escolha da Sub-rotina:** Assistente apresenta as 3 opções:
     - **`➕ Criar utilizador`** (`L. CUA_CRIAR_USER.py` / Categoria `CUA_CRIAR_USER`)
     - **`🔑 Alterar Senha`** (`su01_reset_password.py` / Categoria `CUA Login`)
     - **`📅 Delimitar data fim`** (`I. CUA_ENDDATE.py` / Categoria `CUA_ENDDATE`)
  3. **Modo de Alteração:** **📊 Alteração Massiva (Excel)** ou **👤 Alteração Individual (Chat)**.
  4. **Recolha dos 6 Campos CUA & Prefixo 'S' no Utilizador SAP:**
     - **Formatação de Utilizador SAP:** Quando o utilizador CUA é derivado do número de colaborador (PERNR), deve ser obrigatoriamente adicionado o prefixo **`S`** na frente (ex: colaborador `80002000` $\rightarrow$ Utilizador `S80002000`).
     - Ao selecionar o colaborador do RH (ou utilizador por cópia), o assistente solicita/confirma interativamente no chat os 6 campos obrigatórios:
       - **Nome (`NAME_FIRST`)** — Pré-preenchido do RH ou editável
       - **Sobrenome (`NAME_LAST`)** — Pré-preenchido do RH ou editável
       - **Email (`SMTP_ADDR`)** — Pré-preenchido do RH ou editável
       - ⚠️ **Função (`FUNCTION`)** — *Não existe na tabela RH, solicitado no chat*
       - ⚠️ **Departamento (`DEPARTMENT`)** — *Não existe na tabela RH, solicitado no chat*
       - ⚠️ **Telefone (`MOB_NUMBER`)** — *Não existe na tabela RH, solicitado no chat*
   5. **Criação de Utilizador por Cópia (SU01 SAP GUI Scripting):**
      - **OKCode:** `/nsu01`
      - **Botão Copiar (Shift+F6):** `wnd[0]/tbar[1]/btn[17]`
      - **Pop-up de Cópia (`wnd[1]`):**
        - `txtGV_COPY_UNAME_SRC`: Utilizador de referência (ex: `S4423`)
        - `txtGV_COPY_UNAME_DST`: Utilizador de destino (ex: `S80002000`)
        - Confirmar Cópia: `wnd[1]/tbar[0]/btn[5]`
      - **Tab Endereço (`tabpADDR`):**
        - `txtSUID_ST_NODE_PERSON_NAME-NAME_LAST`: Sobrenome
        - `txtSUID_ST_NODE_PERSON_NAME-NAME_FIRST`: Nome
        - `cmbSUID_ST_NODE_PERSON_NAME-LANGU`: Idioma (`PT`)
        - `txtSUID_ST_NODE_WORKPLACE-FUNCTION`: Função
        - `txtSUID_ST_NODE_WORKPLACE-DEPARTMENT`: Departamento
        - `txtSUID_ST_NODE_COMM_DATA-MOB_NUMBER`: Telefone
        - `txtSUID_ST_NODE_COMM_DATA-SMTP_ADDR`: Email
      - **Tab Logon (`tabpLOGO`):** `pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD` e `PASSWORD2`
      - **Gravar (Ctrl+S):** `wnd[0]/tbar[0]/btn[11]`
   6. **Submissão Direta do Job (`POST /jobs`):**
      As sub-rotinas ativas de alteração **NÃO** chamam `/api/authorizations/start` (consulta USLA04). Disparam diretamente `POST /jobs` com `task: "sap_cockpit"`, `processo: "<Categoria>"`, `subprocesso: "<Script.py>"`, `ambiente: "CUA"`, o `target_system` (utilizador formatado com prefixo `S`), `reference_user` (quando aplicável) e os 6 parâmetros do utilizador (`first_name`, `last_name`, `email`, `function`, `department` e `mob_number`).

---

## 8. Padrão de Assinatura Web (`executar()`) nos Scripts

Todo o sub-script Python que for invocado a partir do SAP Cockpit Web DEVE exportar uma função de entrada padronizada com a seguinte assinatura:

```python
def executar(ambiente=None, target_user=None, session=None, target_system=None, target_env=None, **kwargs) -> dict:
    ...
```

- **Inspecção Segura (`sap_cockpit_web_ready.py`):** O motor do Cockpit valida a presença de `executar()` ou `main()` via `hasattr()` **antes** de chamar `inspect.signature()`, evitando exceções `AttributeError`.

---

## 9. Sub-rotinas do Processo "Perfil de Autorização"

- **Estrutura de Interação:**
  1. **Seleção Inicial:** O utilizador clica em **`🛡️ Perfil de autorização`**.
  2. **Escolha da Sub-rotina (Ordenadas A-Z):**
     - **`🔍 Analisar autorizações`**:
       - Solicita o utilizador SAP (ex: `CSILVA`).
       - Apresenta a escolha do sistema alvo (`S4P`, `S4D`, `S4Q` ou `SPA`/CUA).
       - **Sistemas de Destino (`S4P`, `S4D`, `S4Q`):** Executa consulta instantânea via **RFC (`pyrfc`)** diretamente na tabela **`AGR_USERS`** do sistema selecionado.
       - **Sistema Central CUA (`SPA` / `SPACLNT001`):** Executa consulta GUI (SE16) na tabela **`USLA04`**.
     - **`➕ Criar role composta`** (`D. PFCG_COMPOSTA_RFC.py` / Categoria `Funções PFCG`)
     - **`➕ Criar role simples`** (`A. PFCG_CREATE_RFC.py` / Categoria `Funções PFCG`)
     - **`❌ Eliminar role`** (`B. PFCG_DELETE_RFC.py` / Categoria `Funções PFCG`)
     - **`🛡️ Gerir objetos de autorização`** (`C. PFCG_AUTHORITY.py` / Categoria `Funções PFCG`)
     - **`➖ Remover role de utilizador`** (`J. CUA_REMOVE.py` / Categoria `Funções PFCG`)


