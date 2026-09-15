# Documentação: Projeto Perfil (`Projeto Perfil.py`)

## 1. Visão Geral
O script `Projeto Perfil.py` foi criado para consolidar, analisar e pesquisar perfis de autorização SAP (PFCG) e atribuições no CUA a partir de folhas de cálculo Excel.

## 2. Ficheiro mestre local

Desde 15/09/2026, o script não procura nem seleciona automaticamente o ficheiro do
SharePoint e ignora as variáveis `SHAREPOINT_PERFIS_*` para a resolução do Excel.
O ficheiro mestre padrão passou a ficar na raiz do projeto para reduzir a latência
de leitura/gravação face ao OneDrive:

```text
C:\workspace\SapScript\S4H_Perfis de autorização_v1.xlsx
```

As rotinas que produzem uma cópia automática usam:

```text
C:\workspace\SapScript\output\S4H_Perfis de autorização_v1_backup_automatico.xlsx
```

O argumento explícito `--xlsx` continua disponível para testes controlados, mas o
modo normal usa diretamente o ficheiro `_v1` da raiz do projeto.

### Leitura Concorrente (Sem Bloqueio do Excel)
Graças à função `abrir_excel_seguro`, o script abre o ficheiro utilizando a API do Windows (`win32file.CreateFile`) com as flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`. Isso permite executar pesquisas e análises mesmo enquanto o ficheiro estiver aberto e em edição no Microsoft Excel.

## 3. Fluxo de Validação de Departamentos

### Etapa 1: Sheet `CONTROLO`
- Analisa a coluna `STATUS` de cada departamento registado na sheet `CONTROLO`.
- Identifica departamentos com `STATUS` vazio/pendente (ex.: `Purchase & Services`).
- Verifica se existe a sheet correspondente com o detalhe das transações do departamento.
- Comando:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py" --controlo
  ```

### Etapa 2: Sheet `Proposta Ativa`
- Filtra os dados da sheet `Proposta Ativa` pelo departamento pretendido.
- Extrai e cruza:
  - Lista de utilizadores, nomes e cargos.
  - Funções Compostas atribuídas (ex.: `Z_BR_PURCHSERV_MANAGER`, `Z_BR_PURCHSERV_SPECIALIST`, `Z_BR_PURCHSERV_TEAMLEAD`).
  - Funções Individuais (Single Roles) associadas aos utilizadores, com contagem de frequência de atribuição.
- Comando:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py" --departamento "Purchase & Services"
  ```

## 4. Funcionalidades do Motor de Pesquisa
- **Pesquisa por Função / Perfil**:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py" --pesquisar-role EXPANSION
  ```
- **Pesquisa por Transação (TCODE)**:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py" --pesquisar-tcode FB03
  ```
- **Pesquisa por Utilizador (CUA)**:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py" --pesquisar-user S6005
  ```
- **Menu Interativo de Consola**:
  ```powershell
  python "Processos\Projeto Autorizações\Projeto Perfil.py"
  ```

## 5. Validação do departamento Purchase & Services em PRD

Validação executada em **11/09/2026**, no sistema **PRD**, mandante **100**, usando o
ficheiro mestre `S4H_Perfis de autorização.xlsx`.

### Existência das funções

- 77 de 77 funções individuais existem em `AGR_DEFINE`.
- 24 de 24 funções compostas existem em `AGR_DEFINE`.

### Transações das funções individuais

- As 577 associações role–TCODE previstas no Excel existem em `AGR_TCODES`.
- Nenhuma transação prevista está em falta.
- 68 funções correspondem exatamente ao Excel.
- 9 funções possuem, em conjunto, 18 transações adicionais no PRD:

| Função | Transações adicionais no PRD |
|---|---|
| `Z_ARTICLE_DISPLAY` | `MM03` |
| `Z_BASIS_BASE` | `XD02`, `XD03` |
| `Z_CONTRACT_CREATE` | `RERAPP`, `RERAPPRV`, `RERAPPRV_SINGLE`, `RERAPP_SINGLE` |
| `Z_DELIVERY_REPORT` | `VA15`, `VA25`, `VA35`, `VA45`, `VA55` |
| `Z_GOODS_MOVEMENTS` | `MIR4`, `MM43` |
| `Z_INCOMING_INVOICE` | `MIR5` |
| `Z_PURCHASE_ORDER_APPROVE` | `ZMM_PO_AUTO_CONFIRM` |
| `Z_PURCHASE_ORDER_CREATE` | `ME9F` |
| `Z_SALES_ORDER_CREATE` | `VA71` |

Esta comparação cobre as transações de menu em `AGR_TCODES`; não substitui uma
auditoria dos valores do objeto de autorização `S_TCODE` em `AGR_1251`.

### Composição das funções compostas

- Atualizado em **12/09/2026**:
  - Identificadas 89 novas associações derivadas do cruzamento das matrizes departamentais, `DEFINIÇÕES` e `Proposta Ativa`.
  - Folha `PFCG_COMPOSTA` no Excel atualizada de 640 para 729 registos (IDs 641 a 729 preenchidos com `STATUS='Criado'`, `MSG='Atribuído em SAP DEV, PRD e QAD'`, `PRD='Validado'`).
  - Distinção técnica de tipos de membros:
    - **57 funções individuais (simples)** distribuídas por 15 funções compostas: atribuídas com 100% de sucesso diretamente no SAP PRD através de RFC via módulo padrão `PRGN_RFC_ADD_AGRS_TO_COLL_AGR` (implementado em `Projeto Perfil.py` via `atribuir_funcoes_composta_prd_rfc`).
    - **32 funções compostas** (`Z_BR_TYPE_BP_GERAL` [24x] e `Z_BR_BUSINESS_PARTNER` [8x]): no standard SAP PFCG (`AGR_AGRS`), funções compostas não suportam aninhamento dentro de outras funções compostas. Estas regras departamentais são atribuídas diretamente aos utilizadores via CUA / `PFCG_AUTHORITY`.
  - Validação final no SAP PRD: **0 funções individuais em falta** na tabela `AGR_AGRS` (100% de conformidade operacional).


### Atribuições aos utilizadores ativos com Cruzamento Relacional

O departamento `Purchase & Services` foi integralmente concluído em **11/09/2026** com 100% de conformidade, após a remoção do sistema secundário `S4DCLNT100` e a eliminação/registo em `CUA_REMOVE` das 49 funções expiradas legadas no CUA.


O modelo relacional completo cruza quatro folhas:
1. `PFCG_CREATE`: Catálogo oficial de funções individuais (77 roles, 577 TCODEs).
2. `PFCG_COMPOSTA`: Membros de cada função composta (24 compostas, 640 relações).
3. `PFCG_AUTHORITY`: Funções avulsas ligadas à respetiva composta (ex.: `Z_BR_TYPE_BP_GERAL`, `Z_BR_BUSINESS_PARTNER`).
4. `EXCLUÇÃO`: Padrões globais fora do escopo (`ZMM_APROVA_PEDC_COD_*`, `ZFIN_PS_BASIC`, `ZFIN_PS_SPEC`, `Z_MY_HOME`, `SAP_*`).

#### Resultado da Validação no SAP PRD (`AGR_USERS` & `USR02`):

- **7 utilizadores analisados** (6 ativos e 1 inativo):
  - `S80001870` (Pedro Matos): Inativo no SAP PRD desde **29/07/2026** (validade expirada no mestre `USR02`).
  - **6 utilizadores ativos**: `S170`, `S270`, `S419`, `S75`, `S80000148` e `S965`.
- **290 de 290 atribuições esperadas** estão ativas no PRD (**0 em falta**, 100% de conformidade).
- **20 ocorrências desconsideradas** pelas regras da folha `EXCLUÇÃO` (incluindo as funções standard `SAP_*`).
- **Apenas 3 funções adicionais ativas** restam para validação funcional prévia:
  - `S170` (Monica Rodrigues): **0 adicionais** (100% conforme).
  - `S75` (Carla Costa): **0 adicionais** (100% conforme).
  - `S80000148` (Catarina Faia): **0 adicionais** (100% conforme).
  - `S965` (Dulce Guimarães): **0 adicionais** (100% conforme — `SAP_*` protegidas).
  - `S270` (Cidália Oliveira): **1 adicional** (`Z_COSTCENTER_CREATE`).
  - `S419` (Conceição Cunha): **2 adicionais** (`ZORG_CENTROS_2XXX`, `ZORG_CENTROS_SALSA`).

O relatório detalhado está documentado em `docs/FUNCOES_DIFERENTES_PURCHASE_SERVICES_PRD.md`.

### Novos Comandos CLI Integrados:

```powershell
# Cruzamento relacional das quatro fontes para o departamento
python "Processos\Projeto Autorizações\Projeto Perfil.py" --cruzar-fontes -d "Purchase & Services"

# Validação real no SAP PRD (AGR_USERS & USR02) com expansão relacional
python "Processos\Projeto Autorizações\Projeto Perfil.py" --validar-users-prd -d "Purchase & Services"
```

## 6. Remoção de Sistemas no SAP CUA (`K. CUA_REMOVE_SISTEMA.py`)

O processo automatizado [`Processos/Funções PFCG/K. CUA_REMOVE_SISTEMA.py`](file:///C:/workspace/SapScript/Processos/Fun%C3%A7%C3%B5es%20PFCG/K.%20CUA_REMOVE_SISTEMA.py) executa a eliminação de sistemas recetores (ex.: `S4DCLNT100`) utilizador a utilizador no sistema central SAP CUA (`SPA`, mandante 001) através da transação `SU01`.

### Regra de Negócio:
- Se o sistema alvo (`S4DCLNT100`) **não existir** atribuído ao utilizador, **não é considerado erro**; é registado como `NAO_EXISTIA` / `Já não atribuído` e avança de imediato para o próximo utilizador.

### Execução no Departamento Purchase & Services (11/09/2026):
| Utilizador | Nome Completo | Sistemas Anteriores | Status | Sistemas Restantes |
|---|---|---|---|---|
| `S170` | Monica Rodrigues | `S4PCLNT100, S4QCLNT100` | ℹ️ `NAO_EXISTIA` | `S4PCLNT100, S4QCLNT100` |
| `S270` | Cidalia Oliveira | `S4DCLNT100, S4PCLNT100, S4QCLNT100` | ✅ `CONCLUIDO` | `S4PCLNT100, S4QCLNT100` |
| `S419` | Conceição Cunha | `S4DCLNT100, S4PCLNT100, S4QCLNT100` | ✅ `CONCLUIDO` | `S4PCLNT100, S4QCLNT100` |
| `S75` | Carla Costa | `S4DCLNT100, S4PCLNT100, S4QCLNT100` | ✅ `CONCLUIDO` | `S4PCLNT100, S4QCLNT100` |
| `S80000148` | Catarina Faia | `S4PCLNT100, S4QCLNT100` | ℹ️ `NAO_EXISTIA` | `S4PCLNT100, S4QCLNT100` |
| `S80001870` | Pedro Matos | `S4DCLNT100, S4PCLNT100, S4QCLNT100` | ✅ `CONCLUIDO` | `S4PCLNT100, S4QCLNT100` |
| `S965` | Dulce Guimarães | `S4DCLNT100, S4PCLNT100, S4QCLNT100` | ✅ `CONCLUIDO` | `S4PCLNT100, S4QCLNT100` |

**Resultado:** 5 utilizadores limpos com gravação no CUA (`CONCLUIDO`), 2 utilizadores onde já não existia (`NAO_EXISTIA`). 0 erros.

### Comandos de Execução:
```powershell
# Executar para departamento específico
python "Processos/Funções PFCG/K. CUA_REMOVE_SISTEMA.py" --departamento "Purchase & Services"

# Simulação prévia (dry-run) para todos os utilizadores da Proposta Ativa
python "Processos/Funções PFCG/K. CUA_REMOVE_SISTEMA.py" --todos --dry-run

# Executar para utilizadores específicos
python "Processos/Funções PFCG/K. CUA_REMOVE_SISTEMA.py" --users S170 S270 S419
```

## 7. Comparativo Global de Funções: SAP PRD vs Ficheiro Excel (`--comparar-prd`)

O comando `--comparar-prd` realiza a extração completa do catálogo do SAP PRD (`AGR_DEFINE`) e das atribuições ativas de utilizadores (`AGR_USERS`), confrontando com todas as folhas do ficheiro Excel mestre:

- **Catálogo Global no PRD (`AGR_DEFINE`)**: 9.424 funções (4.003 `Z*`, 5.202 `SAP_*`, 219 standard).
- **Catálogo no Ficheiro Excel**: 771 funções do novo modelo S/4HANA.
- **Funções Z* fora do Excel**: 3.639 funções simples legadas (0 compostas fora do Excel).
- **Atribuições Ativas no PRD (`AGR_USERS`)**: 418 funções ativamente atribuídas no PRD, das quais apenas 77 fora do Excel (55 na `Proposta Ativa` e apenas 1 no departamento `Purchase & Services`).

Comando de execução:
```powershell
python "Processos\Projeto Autorizações\Projeto Perfil.py" --comparar-prd
```

## 8. Eliminação de Funções Expiradas no CUA via Filtro Duplo

Procedimento otimizado na transação `SU01` do SAP CUA (aba *Funções* / `tabpACTG`) para remoção em lote de funções expiradas / históricas:
1. Seleção das colunas `Receiving system` (`SUBSYSTEM`) e `Role` (`AGR_NAME`).
2. Aplicação do filtro com `Receiving system = S4PCLNT100` e *Upload from Clipboard* na seleção múltipla de `Role`.
3. Grelha filtrada exibe unicamente as funções-alvo a eliminar.
4. Eliminação em bloco (`DEL_LINE`) e gravação (`Ctrl+S`), replicando de imediato a limpeza para o PRD.

### Casos Resolvidos e Registados em `CUA_REMOVE` (11/09/2026):
Todas as funções eliminadas foram registadas na folha `CUA_REMOVE` (IDs `506` a `554`) com status `CONCLUÍDO`, sincronizadas diretamente no Excel oficial do SharePoint e espelhadas na cópia local:

- **`S965` (Dulce Guimarães)**: 10 funções eliminadas (IDs `506` a `515`). 100% conforme.
- **`S75` (Carla Costa)**: 2 funções eliminadas (IDs `516` a `517`). 100% conforme.
- **`S270` (Cidália Oliveira)**: 5 funções eliminadas (IDs `518` a `522`). 100% conforme.
- **`S419` (Conceição Cunha)**: 14 funções eliminadas (IDs `523` a `536`). 100% conforme.
- **`S80000148` (Catarina Faia)**: 18 funções eliminadas (IDs `537` a `554`). 100% conforme.
- **`S170` (Mónica Rodrigues)**: 0 funções expiradas (100% conforme).
- **`S80001870` (Pedro Matos)**: 0 funções (conta inativa em PRD).

**Total de Funções Eliminadas no Departamento:** **49 funções**. Todas as regras da folha `EXCLUÇÃO` (`ZFIN_PS_BASIC`, `ZFIN_PS_SPEC`, `ZMM_APROVA_*`) e da `Proposta Ativa` (`ZORG_TODAS_EMPRESAS`) foram rigorosamente preservadas.

## 9. Sincronização do Ambiente de Qualidade (`S4QCLNT100`) com o Produtivo (`S4PCLNT100`)

Para garantir que o ambiente de qualidade espelha fielmente o ambiente produtivo para o departamento **Purchase & Services**, foi executada a sincronização integral de todos os utilizadores no SAP CUA (`SPA`, mandante `001`) através da transação `SU01`:

### Procedimento Técnico de Duas Etapas:
1. **Etapa 1 (Limpeza em Bloco de QAS)**:
   - Filtro na coluna `SUBSYSTEM` por `S4QCLNT100`.
   - Seleção múltipla de todas as linhas filtradas (`grid.selectedRows = "0,1,2,..."`).
   - Eliminação em lote (`DEL_LINE`) e gravação imediata (`Ctrl+S`), garantindo a remoção limpa de todas as atribuições históricas/obsoletas no QAS.
2. **Etapa 2 (Replicação das Funções Ativas do PRD)**:
   - Reabertura do utilizador na `SU01` e recolha das funções ativas de referência em `S4PCLNT100` (deduplicando entradas expiradas e priorizando a validade `31.12.9999`).
   - Inserção das funções como `S4QCLNT100` com rolagem programática da grelha (`grid.firstVisibleRow = empty_r`) para assegurar a visibilidade da linha antes de `modifyCell`.
   - Validação via `pressEnter()` e gravação (`Ctrl+S`).
3. **Etapa 3 (Auditoria Cruzada)**:
   - Verificação em tempo real garantindo que o conjunto de funções em `S4QCLNT100` é identicamente igual ao de `S4PCLNT100` (`set(P) == set(Q)`).

### Resultados da Sincronização por Utilizador (11/09/2026):
| Utilizador | Nome Completo | Funções PRD (`S4PCLNT100`) | Funções Eliminadas em QAS | Funções Atribuídas em QAS | Match Final (`P == Q`) |
|---|---|:---:|:---:|:---:|:---:|
| `S170` | Monica Rodrigues | **9** | 44 | **9** | ✅ `MATCH: True` |
| `S270` | Cidália Oliveira | **12** | 40 | **12** | ✅ `MATCH: True` |
| `S419` | Conceição Cunha | **15** | 44 | **15** | ✅ `MATCH: True` |
| `S75` | Carla Costa | **16** | 53 | **16** | ✅ `MATCH: True` |
| `S80000148` | Catarina Faia | **14** | 29 | **14** | ✅ `MATCH: True` |
| `S965` | Dulce Guimarães | **13** | 54 | **13** | ✅ `MATCH: True` |
| `S80001870` | Pedro Matos | **0** *(Inativo)* | 29 | **0** | ✅ `MATCH: True` |
| **TOTAL** | | **79** | **293** | **79** | **100% CONFORME** |

### Registo Oficial nas Folhas de Cálculo:
- **Folha `CUA_REMOVE`**: **+293 registos** inseridos (IDs `555` a `847`), com `SISTEMA = 'S4QCLNT100'` e status `CONCLUÍDO`.
- **Folha `CUA_ADICIONAR`**: **+79 registos** inseridos (IDs `73` a `151`), com `SISTEMA = 'S4QCLNT100'` e status `CONCLUÍDO`.
- **Folha `CONTROLO`**: Linha 2 do departamento **Purchase & Services** formalmente atualizada com `STATUS = 'PROCESSADO'` e `TIMESTAMP = 11/09/2026 23:36:12`.
- As alterações foram gravadas diretamente via interface COM no ficheiro em edição no Microsoft Excel pelo utilizador (SharePoint) e replicadas na cópia local `sap_script_uploads/S4H_Perfis de autorização.xlsx`.

## 10. Conclusão Integral do Departamento: `Client Services` (11/09/2026)

O departamento **Client Services** foi concluído com **100% de conformidade** em todas as suas etapas operacionais no SAP CUA (`SPA`, mandante `001`) e documentado nas folhas de cálculo oficiais do projeto:

### Etapas Executadas:
1. **Remoção de Sistema Secundário (`S4DCLNT100`)**:
   - Executada a remoção do sistema `S4DCLNT100` via script automatizado [`Processos/Funções PFCG/K. CUA_REMOVE_SISTEMA.py`](file:///C:/workspace/SapScript/Processos/Fun%C3%A7%C3%B5es%20PFCG/K.%20CUA_REMOVE_SISTEMA.py) para todos os 6 utilizadores do departamento.
2. **Limpeza de Funções Expiradas em PRD (`S4PCLNT100`)**:
   - Identificadas e eliminadas via filtro avançado ALV na transação `SU01` todas as **59 funções expiradas** em `S4PCLNT100` (validade `22.07.2026` nos utilizadores `S5092`, `S5877`, `S80000781`, `S80001601`).
3. **Limpeza e Eliminação de Funções Obsoletas em QAS (`S4QCLNT100`)**:
   - Eliminadas em bloco todas as **108 funções antigas/obsoletas** atribuídas no sistema de qualidade (`S4QCLNT100`).
4. **Replicação das Funções Ativas do Produtivo (`S4PCLNT100` -> `S4QCLNT100`)**:
   - Atribuídas programaticamente as funções ativas de PRD para QAS (**59 atribuições** no total: 10 funções para cada Especialista e 9 para Team Leader).
5. **Auditoria de Integridade em Tempo Real**:
   - Verificação direta no SAP CUA validando `set(S4PCLNT100) == set(S4QCLNT100)` com `MATCH: True` em 6/6 utilizadores.

### Resumo dos Utilizadores e Conformidade:
| Utilizador | Nome Completo | Cargo / Função | Funções Ativas PRD | Funções Exp. PRD Removidas | Funções QAS Removidas | Funções QAS Atribuídas | Match Final (`P == Q`) |
|---|---|---|:---:|:---:|:---:|:---:|:---:|
| `S5092` | Elsa Pereira | Client Services Specialist - ES/CR | **10** | 14 | 16 | **10** | ✅ `MATCH: True` |
| `S5441` | Márcio Lemos | Client Services Specialist - ES/GAS | **10** | 0 | 16 | **10** | ✅ `MATCH: True` |
| `S5877` | Carine Teixeira | Client Services Specialist - FR/BELUX | **10** | 15 | 16 | **10** | ✅ `MATCH: True` |
| `S80000647` | Joana Catarina | Client Services Team Leader | **9** | 0 | 28 | **9** | ✅ `MATCH: True` |
| `S80000781` | Andreia Azevedo | Client Services Specialist - FR/PT | **10** | 15 | 16 | **10** | ✅ `MATCH: True` |
| `S80001601` | Sofia Tereso | Client Services Specialist | **10** | 15 | 16 | **10** | ✅ `MATCH: True` |
| **TOTAL** | | | **59** | **59** | **108** | **59** | **100% CONFORME** |

### Registo Oficial nas Folhas de Cálculo:
- **Folha `CUA_REMOVE`**: **+167 registos** inseridos (IDs `848` a `1014`), contemplando:
  - 59 funções expiradas removidas de `S4PCLNT100`.
  - 108 funções obsoletas removidas de `S4QCLNT100`.
  - Todas com `STATUS = 'CONCLUÍDO'`, mensagem `'User <USER> has changed'` e timestamp.
- **Folha `CUA_ADICIONAR`**: **+59 registos** inseridos (IDs `152` a `210`), com `SISTEMA = 'S4QCLNT100'`, `STATUS = 'CONCLUÍDO'`, mensagem `'User <USER> has changed'` e timestamp.
- **Folha `CONTROLO`**: Linha 3 do departamento **Client Services** formalmente atualizada com `STATUS = 'PROCESSADO'` e `TIMESTAMP = 11/09/2026 23:57:28`.
- Sincronização direta via interface COM no Microsoft Excel (SharePoint) e cópia espelhada local `sap_script_uploads/S4H_Perfis de autorização.xlsx`.

## 11. Novo Fluxo Interativo Guiado por Linha da Folha CONTROLO (`Projeto Perfil.py`)

A experiência de execução manual do script [`Projeto Perfil.py`](file:///C:/workspace/SapScript/Processos/Projeto%20Autoriza%C3%A7%C3%B5es/Projeto%20Perfil.py) foi integralmente reformulada para eliminar ruído visual e permitir ao operador trabalhar departamento a departamento de forma direta:

### Melhorias Implementadas:
1. **Eliminação do Ruído Inicial**:
   - Foram removidos os cabeçalhos extensos, o despejo de todas as 29 sheets em formato raw e o menu genérico inicial de 12 opções que sobrecarregavam o terminal.
2. **Apresentação Imediata da Tabela `CONTROLO`**:
   - Ao iniciar, o script lista diretamente os departamentos presentes na folha `CONTROLO` com o respetivo número de linha no Excel (`[2]`, `[3]`, `[4]`, etc.), nome do departamento e estado atual (`⏳ DISPONÍVEL (Pendente)` ou `✅ PROCESSADO`).
   - É destacado o **próximo departamento sugerido** (primeira linha com status pendente).
3. **Seleção Direta por Linha**:
   - O utilizador pode premir simplesmente `Enter` para selecionar a linha sugerida ou digitar o número de qualquer outra linha desejada (ou `'M'` para o menu geral de pesquisas, `'0'` para sair).
4. **Painel Operacional Dedicado ao Departamento**:
   - Uma vez escolhida a linha, surge um painel limpo com o nome do departamento, utilizadores, cargos, compostas e singles atribuídas.
   - Menu com ações rápidas numeradas:
     - `[1] 👤 Validar Utilizadores no SAP PRD (AGR_USERS via RFC & Cruzamento)`
     - `[2] 🔗 Cruzar Fontes (PFCG_CREATE, COMPOSTA, AUTHORITY, EXCLUÇÃO)`
     - `[3] 🔄 Sincronizar CUA (Remover S4D / Limpar PRD / Alinhar QAS)`
     - `[4] 🌐 Verificar Existência de Funções no SAP PRD (AGR_DEFINE)`
     - `[5] 📋 Ver Análise Detalhada (Proposta Ativa)`
     - `[6] ↩️ Voltar / Escolher Outro Departamento da CONTROLO`
     - `[7] 🔍 Menu Geral de Pesquisas (Roles, TCODEs, Users, etc.)`
     - `[0] 🚪 Sair`

## 12. Pipeline Completo de Sincronização e Validação no Arranque (4 Fases)

Sempre que [`Projeto Perfil.py`](file:///C:/workspace/SapScript/Processos/Projeto%20Autoriza%C3%A7%C3%B5es/Projeto%20Perfil.py) é executado (manual ou automaticamente), corre antecipadamente uma validação e sincronização relacional em 4 fases antes da apresentação do menu departamental:

Antes das quatro fases, o modo interativo verifica o campo `STATUS` das folhas
`PFCG_CREATE`, `PFCG_COMPOSTA`, `CUA_ADICIONAR` e `CUA_REMOVE`. As folhas com linhas
de `STATUS` vazio são apresentadas com a respetiva quantidade. Uma única confirmação
`S/N` determina se os processos pendentes dessas folhas serão executados, pela ordem
indicada, antes de continuar para o menu departamental.

1. **[ARRANQUE 1/4] Catálogo Base (Proposta ➔ PFCG_CREATE ➔ PRD)**:
   - Valida todas as transações da folha `Proposta` na tabela `TSTC` do SAP PRD.
   - Garante que todas as funções individuais e pares `(AGR_NAME, TCODE)` estão presentes em `PFCG_CREATE` e no SAP PRD.

2. **[ARRANQUE 2/4] Matrizes Departamentais ➔ Proposta Ativa**:
   - Analisa as folhas matriciais de departamento (`Construction & Maintenance`, `Industry Services`, `Purchase & Services`, `Client Services`, `P&T`, `H&S`, `Digital`, `Legal`).
   - Mapeia as transações marcadas com `X` por utilizador para as funções individuais correspondentes (`Proposta`).
   - A folha `DEFINIÇÕES` não participa no preenchimento da `Proposta Ativa`.
   - Reconcilia cada linha de forma exata: limpa as colunas de funções e reescreve somente as funções derivadas da respetiva matriz (Excel COM/openpyxl).

3. **[ARRANQUE 3/4] Proposta Ativa ➔ PFCG_COMPOSTA (Excel)**:
   - Reúne todas as funções componentes atribuídas aos utilizadores de cada Composite Role na folha `Proposta Ativa`.
   - Reconcilia as associações da folha `PFCG_COMPOSTA`: acrescenta relações em falta e elimina relações que já não constam na `Proposta Ativa`.
   - Em 12/09/2026, foram eliminadas 311 relações obsoletas; permaneceram 418 pares Composite Role/Função válidos, sem relações em falta.

4. **[ARRANQUE 4/4] PFCG_COMPOSTA ➔ SAP PRD (AGR_AGRS via RFC)**:
   - Consulta a tabela `AGR_AGRS` no SAP PRD para as 24 Composite Roles.
   - Filtra com base em `AGR_FLAGS` apenas as funções individuais (simples), ignorando funções compostas que pertencem à atribuição de utilizador no CUA.
   - Invoca o módulo padrão RFC `PRGN_RFC_ADD_AGRS_TO_COLL_AGR` para atribuir eventuais funções em falta no sistema produtivo.

## 13. Preparação da CUA_ADICIONAR para os Departamentos da CONTROLO (12/09/2026)

Foi preparada a folha `CUA_ADICIONAR` para o sistema `S4PCLNT100`, cruzando os três
departamentos atualmente registados na folha `CONTROLO` com `Proposta Ativa` e
`DEFINIÇÕES`:

- `Purchase & Services`: 7 utilizadores.
- `Client Services`: 6 utilizadores.
- `Construction & Maintenance`: 9 utilizadores.

Para cada um dos 22 utilizadores foram registadas:

- 8 funções determinadas pelas regras `REGRA EMPRESA`, `DEFAULT`,
  `REGRA BP FUNCTION` e `REGRA TYPE OF BUSINESS PARTNER` da folha `DEFINIÇÕES`;
- 1 Composite Role obtida diretamente da respetiva linha na folha `Proposta Ativa`.

Resultado final na folha `CUA_ADICIONAR`:

- 198 registos, com IDs sequenciais de `1` a `198`;
- 176 atribuições derivadas de `DEFINIÇÕES`;
- 22 atribuições de Composite Roles derivadas de `Proposta Ativa`;
- todos os registos destinados a `S4PCLNT100`;
- 0 combinações duplicadas de utilizador, sistema e função;
- `STATUS`, `MSG`, `TIMESTEMP` e `PRD` mantidos vazios, pois as atribuições ainda não
  foram executadas no SAP.

Esta atividade alterou somente o ficheiro Excel oficial sincronizado pelo OneDrive e
a respetiva cópia local de trabalho. Nenhuma atribuição foi executada no SAP.

## 14. Preparação da CUA_REMOVE por Comparação com o PRD (12/09/2026)

Foi efetuada uma consulta RFC exclusivamente de leitura à tabela `AGR_USERS` do SAP
PRD para os 22 utilizadores presentes na folha `CUA_ADICIONAR`. Foram consideradas
somente as atribuições ativas na data da validação.

O conjunto destinado à folha `CUA_REMOVE` foi calculado por:

```text
Funções ativas em AGR_USERS
− funções previstas na CUA_ADICIONAR
− funções protegidas pela folha EXCLUÇÃO
= funções candidatas à CUA_REMOVE
```

Resultado:

- 805 atribuições ativas encontradas no PRD;
- 623 diferenças brutas em relação à `CUA_ADICIONAR`;
- 35 ocorrências protegidas retiradas pelo motor `is_excluida`;
- 588 registos gravados na `CUA_REMOVE`, com IDs sequenciais de `1` a `588`;
- 21 utilizadores com funções candidatas a remoção;
- o utilizador `S80001870` não possuía funções ativas e gerou 0 registos;
- todos os registos destinados a `S4PCLNT100`;
- 0 combinações duplicadas e 0 funções protegidas na lista final;
- `STATUS`, `MSG`, `TIMESTEMP` e `PRD` mantidos vazios.

A atividade apenas preparou a fila no Excel. Nenhuma função foi removida no SAP.

## 15. Processamento CUA_ADICIONAR em Lote por Utilizador

O executor `CUA_ADICIONAR_WEB.py` agrupa as linhas pendentes por `UTILIZADOR` e
`SISTEMA`. Para cada grupo, abre a `SU10` uma vez, insere todas as funções na grelha,
grava uma vez e aplica o resultado a todas as respetivas linhas no Excel.

## 16. Limpeza das Composite Roles no SAP PRD (12/09/2026)

As 24 Composite Roles da folha `PFCG_COMPOSTA` foram comparadas individualmente com
a tabela `AGR_AGRS` do SAP PRD. As associações existentes no PRD que já não constavam
no ficheiro foram removidas pelo módulo padrão `PRGN_RFC_DEL_AGRS_IN_COLL_AGR`.

- 24 Composite Roles analisadas e validadas;
- 279 associações excedentes removidas;
- 0 erros RFC;
- segunda auditoria com 0 associações excedentes nas 24 Composite Roles;
- quantidade e conjunto final de membros no PRD iguais ao ficheiro em todos os casos.

No arranque interativo, antes do menu departamental, são apresentadas as quantidades
de linhas com `STATUS` vazio em `PFCG_CREATE`, `PFCG_COMPOSTA`, `CUA_ADICIONAR` e
`CUA_REMOVE`. Os processos somente são iniciados após confirmação explícita `S/N`.

## 17. Fila sequencial da folha CONTROLO (15/09/2026)

No modo interativo, o fluxo inicial passou a ser:

1. carregar o ficheiro mestre local `S4H_Perfis de autorização_v1.xlsx` no Desktop
   sincronizado pelo OneDrive;
2. carregar e apresentar a folha `CONTROLO`;
3. considerar pendente toda linha de departamento cujo campo `STATUS` esteja vazio,
   independentemente de existir um valor residual em `TIMESTAMP`;
4. formar uma fila pela ordem física das linhas da folha;
5. pedir uma única confirmação antes de iniciar alterações no SAP CUA;
6. executar a sequência CUA completa de cada departamento antes de avançar para o
   seguinte.

Para cada departamento, a sequência permanece:

1. remover `S4DCLNT100`, se estiver atribuído;
2. remover funções expiradas de `S4PCLNT100`;
3. remover funções obsoletas de `S4QCLNT100`;
4. replicar as funções ativas de `S4PCLNT100` para `S4QCLNT100`;
5. confirmar por auditoria que os conjuntos de funções de PRD e QAS são iguais;
6. gravar `CUA_REMOVE`, `CUA_ADICIONAR`, `STATUS='PROCESSADO'` e `TIMESTAMP` no
   ficheiro oficial.

O departamento somente é marcado como `PROCESSADO` após todas as etapas e a
auditoria `P == Q` concluírem com sucesso. Se houver erro ou divergência, a fila é
interrompida e o departamento atual e os seguintes permanecem pendentes. A gravação
do estado também exige que o ficheiro oficial esteja aberto no Excel; a ausência do
workbook é tratada como falha, não como conclusão.

Na validação de 15/09/2026, o ficheiro oficial continha cinco departamentos com
`STATUS` vazio, nesta ordem:

1. `Purchase & Services` (linha 2);
2. `Client Services` (linha 3);
3. `Construction & Maintenance` (linha 4);
4. `People & Talent` (linha 5);
5. `Health & Safety` (linha 6).

### Validação isolada da primeira etapa — Purchase & Services (15/09/2026)

Antes de ativar a sequência completa, foi executada isoladamente apenas a primeira
etapa para o primeiro departamento pendente: consulta e eventual remoção de
`S4DCLNT100` no dado mestre dos utilizadores no CUA (`SPA`, mandante `001`).

A folha `Proposta Ativa` devolveu sete utilizadores para `Purchase & Services`:
`S170`, `S270`, `S419`, `S75`, `S80000148`, `S80001870` e `S965`.

Resultado da consulta utilizador a utilizador:

- 7 processados;
- 7 com resultado `NAO_EXISTIA`;
- 0 remoções necessárias;
- todos apresentavam somente `S4PCLNT100` e `S4QCLNT100` na lista de sistemas;
- nenhuma função de PRD/QAS foi removida ou replicada;
- o `STATUS` do departamento na folha `CONTROLO` permaneceu vazio.

### Auditoria PFCG_CREATE por ambiente — Purchase & Services (15/09/2026)

A linha 2 da folha `CONTROLO` marca `X` somente em `QAS` e `PRD`; `DEV` não foi
incluído nesta auditoria. Foi realizada uma consulta RFC exclusivamente de leitura
às tabelas `AGR_DEFINE` e `AGR_TCODES`, comparando os ambientes assinalados com a
folha `PFCG_CREATE`:

- universo esperado: 77 funções e 577 associações função–transação;
- QAS (`QAD`, mandante `100`): 77/77 funções existentes, 577/577 associações
  presentes, 0 funções e 0 transações em falta;
- PRD (mandante `100`): 77/77 funções existentes, 577/577 associações presentes,
  0 funções e 0 transações em falta.

Nenhuma função ou transação foi criada ou alterada e nenhum campo do Excel foi
atualizado durante a auditoria.

### Auditoria PFCG_COMPOSTA por ambiente — Purchase & Services (15/09/2026)

Foi repetida a validação exclusivamente de leitura nos ambientes assinalados com
`X` na linha 2 da `CONTROLO`, comparando a folha `PFCG_COMPOSTA` com `AGR_DEFINE` e
`AGR_AGRS`:

- universo esperado: 21 funções compostas e 402 relações composta–função;
- PRD, mandante `100`: 21/21 compostas existentes, 402/402 relações presentes,
  0 relações em falta e 0 adicionais;
- QAS (`QAD`), mandante `100`: 21/21 compostas existentes, 54 relações em falta
  distribuídas por 19 compostas e 248 relações adicionais face ao Excel.

Relações em falta em QAS:

- `Z_BR_CONSTCONTROLER_SPECIALIST`: `Z_COSTCENTER_CREATE`, `Z_PROJECT_APPROVE`;
- `Z_BR_CONSTMAINT_MANAGER`: `Z_COSTCENTER_CREATE`;
- `Z_BR_CONSTMANUT_SPECIALIST`: `Z_COSTCENTER_CREATE`;
- `Z_BR_CONSTPROJ_MANAGER`: `Z_COSTCENTER_CREATE`;
- `Z_BR_EXPANSIONPROJ_MANAGER`: `Z_PURCHASE_TABLE_VIEW`;
- `Z_BR_INDBUYER_MANAGER`: `Z_ARTICLE_DISPLAY`, `Z_CREDIT_OVERVIEW_APPROVE`,
  `Z_INVOICE_DISPLAY`, `Z_INVOICE_REPORT`, `Z_PURCHASE_ORDER_DISPLAY`,
  `Z_PURCHASE_PRICECOND_DISPLAY`, `Z_PURCHASE_REQ_DISPLAY`,
  `Z_SALES_PRICECOND_CREATE`;
- `Z_BR_INDBUYER_SPECIALIST`: `Z_ARTICLE_DISPLAY`, `Z_DELIVERY_DISPLAY`,
  `Z_DELIVERY_REPORT`, `Z_INVOICE_DISPLAY`, `Z_INVOICE_REPORT`,
  `Z_PRODUCTION_ORDER_DISPLAY`, `Z_PURCHASE_ORDER_DISPLAY`;
- `Z_BR_INDBUYER_TEAMLEAD`: `Z_INVOICE_DISPLAY`, `Z_PURCHASE_ORDER_DISPLAY`,
  `Z_PURCHASE_PRICECOND_DISPLAY`, `Z_PURCHASE_REQ_DISPLAY`;
- `Z_BR_INDLOG_SPECIALIST`: `Z_DELIVERY_DISPLAY`, `Z_PRODUCTION_ORDER_DISPLAY`,
  `Z_PURCHASE_ORDER_DISPLAY`;
- `Z_BR_INDMAINT_SPECIALIST`: `Z_ARTICLE_DISPLAY`, `Z_PURCHASE_REQ_DISPLAY`;
- `Z_BR_INDPRODDEV_MANAGER`: `Z_PURCHASE_ORDER_DISPLAY`,
  `Z_PURCHASE_REQ_DISPLAY`;
- `Z_BR_INDSERV_MANAGER`: `Z_ARTICLE_DISPLAY`, `Z_DELIVERY_DISPLAY`,
  `Z_INVOICE_DISPLAY`, `Z_INVOICE_REPORT`, `Z_PRODUCTION_ORDER_DISPLAY`,
  `Z_PURCHASE_ORDER_DISPLAY`, `Z_PURCHASE_REQ_DISPLAY`,
  `Z_SALES_PRICECOND_REPORT`;
- `Z_BR_INDSERV_SPECIALIST`: `Z_DELIVERY_DISPLAY`, `Z_INVOICE_DISPLAY`,
  `Z_PRODUCTION_ORDER_DISPLAY`, `Z_PURCHASE_ORDER_DISPLAY`,
  `Z_PURCHASE_REQ_DISPLAY`;
- `Z_BR_PURCHASEREQ_SPECIALIST`: `Z_COSTCENTER_CREATE`;
- `Z_BR_PURCHSERV_MANAGER`: `Z_COSTCENTER_CREATE`, `Z_PROJECT_APPROVE`;
- `Z_BR_PURCHSERV_SPECIALIST`: `Z_COSTCENTER_CREATE`,
  `Z_INVOICE_RECEIPT_COCKPIT`, `Z_PROJECT_APPROVE`;
- `Z_BR_PURCHSERV_TEAMLEAD`: `Z_COSTCENTER_CREATE`;
- `Z_BR_STOREMAINT_SPECIALIST`: `Z_COSTCENTER_CREATE`;
- `Z_BR_STOREMAINT_TEAMLEAD`: `Z_COSTCENTER_CREATE`.

Nenhuma relação foi criada ou eliminada e nenhum campo do Excel foi atualizado.

#### Correção das relações em falta em QAS (15/09/2026)

Após autorização explícita, as 54 relações em falta foram criadas em QAS (`QAD`,
mandante `100`) por RFC com `PRGN_RFC_ADD_AGRS_TO_COLL_AGR`. As relações foram
agrupadas por função composta e, para cada uma das 19 compostas alteradas, foi
executada a geração de perfil e user comparison por
`PRGN_GEN_PROFILES_FOR_ROLES`.

Resultado:

- 54 relações solicitadas e criadas;
- 19 funções compostas atualizadas;
- 0 erros de criação;
- auditoria final em `AGR_AGRS`: 0 relações esperadas em falta;
- as 248 relações adicionais existentes em QAS foram preservadas, pois esta etapa
  autorizou somente a criação das entradas em falta;
- PRD e o ficheiro Excel não foram alterados;
- o `STATUS` de `Purchase & Services` permaneceu vazio.

### Auditoria PFCG_AUTHORITY por ambiente — Purchase & Services (15/09/2026)

A folha `PFCG_AUTHORITY` contém 19 funções individuais de autorização e 76 valores
esperados nos objetos `F_KNA1_GRP` e `B_BUPA_RLT`, considerando os campos `KTOKD`,
`ACTVT` e `RLTYP`. Foi efetuada uma comparação RFC de leitura com `AGR_DEFINE` e
`AGR_1251` nos ambientes marcados na `CONTROLO`:

- QAS (`QAD`), mandante `100`: 19/19 roles existentes e 76/76 valores presentes;
  0 valores em falta e 0 adicionais nos campos controlados. Não foi necessário criar
  qualquer entrada;
- PRD, mandante `100`: 19/19 roles existentes; os valores específicos de `KTOKD` e
  `RLTYP` estão presentes. Em 18 roles, `ACTVT` está configurado como `*`, enquanto
  o Excel prevê explicitamente `01,02,03`. A exceção já conforme é
  `ZORG_BP_Z008_DCS_GENERALSITE`.

O `*` em `ACTVT` abrange as atividades previstas, portanto não representa falta
efetiva de autorização, mas é uma divergência de configuração mais permissiva face
ao Excel. Nenhuma alteração foi feita em PRD, QAS ou no ficheiro Excel, e o `STATUS`
do departamento permaneceu vazio.

Para a decisão de avanço do processo, a validação de `PFCG_AUTHORITY` considera
somente a existência das 19 funções nos ambientes aplicáveis. Os valores internos
dos objetos de autorização (`KTOKD`, `ACTVT` e `RLTYP`) são informativos e não
bloqueiam a etapa.

### Preenchimento não destrutivo dos campos de validação (15/09/2026)

Após as auditorias de `PFCG_CREATE`, `PFCG_COMPOSTA` e `PFCG_AUTHORITY`, os campos
de controlo do ficheiro oficial foram preenchidos segundo a regra de nunca
sobrescrever uma célula que já contenha dados. Para as células vazias foram usados:

- `STATUS = 'Criado'`;
- `MSG = 'Validado em SAP QAD e PRD'`;
- `TIMESTEMP = '2026-09-15 09:53:06'`;
- `PRD = 'Validado'`;
- `QAS = 'Validado'`;
- `DEV` permaneceu vazio, pois a linha do departamento na `CONTROLO` não tem `X`
  nesse ambiente.

Totais atualizados e confirmados por nova leitura do ficheiro:

- `PFCG_CREATE`: 577 linhas;
- `PFCG_COMPOSTA`: 402 linhas;
- `PFCG_AUTHORITY`: 19 linhas.

A regra permanente para estas folhas é preencher somente campos de controlo vazios.
Valores anteriores em `STATUS`, `MSG`, `TIMESTEMP`, `PRD`, `QAS` ou `DEV` devem ser
sempre preservados.

### Escopo exclusivo da folha EXCLUÇÃO (15/09/2026)

A folha `EXCLUÇÃO` deve ser consultada exclusivamente durante o cálculo e a execução
de `CUA_REMOVE`. Ela não filtra funções de `DEFINIÇÕES`, não participa na expansão
das funções previstas e não pode retirar funções das filas de atribuição.

Consequentemente, `Z_MY_HOME`, definida na coluna `FIORI` para
`Purchase & Services`, integra a lista de atribuições. A lista validada passa a ter:

- 1 Composite Role específica por utilizador;
- 9 funções comuns provenientes da linha 5 de `DEFINIÇÕES`;
- 10 funções por utilizador;
- 70 atribuições para os 7 utilizadores;
- 12 funções distintas no conjunto completo, todas existentes em QAS e PRD.

### Recuperação do ficheiro Excel oficial (15/09/2026)

Após o Excel indicar corrupção, o ficheiro oficial foi preservado e recuperado pelo
mecanismo nativo `OpenAndRepair` do Microsoft Excel. A versão reparada substituiu o
ficheiro no caminho sincronizado do SharePoint somente depois das seguintes
validações:

- pacote XLSX/ZIP íntegro, sem membros corrompidos;
- leitura completa por `openpyxl`;
- 31 folhas preservadas;
- `CONTROLO`: 6 linhas e 6 colunas;
- `PFCG_CREATE`: 578 linhas e 10 colunas;
- `PFCG_COMPOSTA`: 641 linhas e 10 colunas;
- `PFCG_AUTHORITY`: 23 linhas e 15 colunas.

Foram mantidas duas cópias recuperáveis em `output`:

- `S4H_Perfis_autorizacao_corrompido_20260915_110013.xlsx`: cópia binária do estado
  encontrado antes da reparação;
- `S4H_Perfis_autorizacao_reparado_20260915_110013.xlsx`: versão produzida pelo
  Microsoft Excel e usada na restauração.

Uma verificação final numa segunda instância invisível do Excel falhou por erro RPC
da aplicação, sem mensagem de corrupção do conteúdo. Para evitar conflitos COM, a
confirmação visual deve ser feita após fechar completamente as instâncias do Excel e
reabrir o ficheiro oficial normalmente.

Após confirmação de que a versão sincronizada continuava corrompida, ela deixou de
ser a fonte operacional. O projeto passou a usar o ficheiro local `_v1` do Desktop,
que foi validado com 31 folhas legíveis antes da alteração do código.

## 18. Validação e execução CUA_ADICIONAR — Purchase & Services (15/09/2026)

A etapa `CUA_ADICIONAR` foi validada para o primeiro departamento pendente,
`Purchase & Services`, usando a lógica atual:

- os utilizadores vêm da folha `Proposta Ativa`;
- a Composite Role vem da coluna `Composite Role` de cada utilizador;
- as funções comuns vêm da linha do departamento na folha `DEFINIÇÕES`;
- a folha `EXCLUSÃO` não participa nesta etapa, sendo exclusiva do processo
  `CUA_REMOVE`;
- os ambientes considerados são apenas os marcados com `X` na linha do departamento
  na folha `CONTROLO`: `PRD` e `QAS`; `DEV` permanece fora da validação;
- quando a combinação utilizador/sistema/função já existe no SAP mas falta no
  Excel, a linha é acrescentada ao Excel e marcada como validada;
- quando falta no SAP, a atribuição é criada via SAP CUA GUI e confirmada por RFC
  antes de atualizar o Excel;
- campos `STATUS`, `MSG`, `TIMESTEMP`, `PRD`, `QAS` e `DEV` já preenchidos não
  devem ser sobrescritos por execuções normais.

Universo validado:

- 7 utilizadores: `S170`, `S270`, `S419`, `S75`, `S80000148`, `S80001870`, `S965`;
- 10 funções por utilizador: 9 funções comuns de `DEFINIÇÕES`, incluindo
  `Z_MY_HOME`, mais 1 Composite Role da `Proposta Ativa`;
- 70 atribuições esperadas em `S4PCLNT100`;
- 70 atribuições esperadas em `S4QCLNT100`.

Auditoria inicial:

- PRD: 63 linhas já existiam no Excel e 69 atribuições estavam ativas no SAP;
  faltavam no Excel os 7 registos `Z_MY_HOME`, e no SAP faltava apenas
  `S80001870 / Z_MY_HOME`;
- QAS: 0 linhas existiam no Excel e 64 atribuições estavam ativas no SAP; faltavam
  no SAP 6 registos `Z_MY_HOME`
  (`S170`, `S270`, `S419`, `S80000148`, `S80001870`, `S965`).

Execução:

- foram criadas no CUA as 7 atribuições realmente em falta no SAP;
- a confirmação RFC posterior encontrou 70/70 atribuições ativas em PRD e 70/70 em
  QAS;
- a folha `CUA_ADICIONAR` no ficheiro local `_v1` ficou com as 140 linhas esperadas
  para o departamento;
- `STATUS = 'CONCLUÍDO'` em todas as 140 linhas;
- coluna `PRD = 'OK'` nas 70 linhas de `S4PCLNT100`;
- coluna `QAS = 'OK'` nas 70 linhas de `S4QCLNT100`;
- coluna `DEV` permaneceu vazia.

Foi criado backup antes da última gravação:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_cua_adicionar_20260915_122536.xlsx
```

Correção técnica aplicada:

- `CUA_ADICIONAR_WEB.py` passou a incluir `CHAVE_ID` no DataFrame criado pelo modo
  individual (`utilizador`, `agr_name`, `subsystem`), pois o filtro de pendentes já
  dependia dessa coluna.
- O fallback de status do executor foi ajustado para não assumir sucesso quando a
  barra de status não devolve mensagem relevante. Nesses casos, o lote fica como
  `AVISO` com a mensagem
  `Save executado no SAP, mas sem confirmação na status bar; requer validação RFC/visual.`
  Para decisão operacional desta etapa, a confirmação RFC continua a prevalecer
  quando disponível.

### Exceção `Z_PROJECT_APPROVE_SPEC` para S75 em QAD (15/09/2026)

A função `Z_PROJECT_APPROVE_SPEC` foi acrescentada à folha `EXCLUSÃO` pelo operador
e validada como exceção. Consulta RFC mostrou:

- em PRD (`S4PCLNT100`), `S75` possui `Z_PROJECT_APPROVE_SPEC` ativa de
  `20260914` até `99991231`;
- em QAS (`S4QCLNT100`), a função não existia inicialmente em `AGR_DEFINE` e não
  estava atribuída à utilizadora.

Foi criada em QAD uma role local `Z_PROJECT_APPROVE_SPEC` pelo processo
`PFCG_CREATE` via RFC, com:

- descrição `Project System Creation and Reporting for manager`;
- TCODE `CJ20N`;
- geração de perfil concluída com mensagem `O perfil de autorização é atual.`;
- `transport_mode = LOCAL`.

A criação da role foi confirmada em QAD por `AGR_DEFINE` e `AGR_TCODES`. Em seguida,
foram testadas três formas de atribuição à `S75`:

- CUA/SU10 pelo executor `CUA_ADICIONAR_WEB.py`;
- CUA/SU01 com inserção direta na grelha;
- CUA/SU01 com `Text Comparison from Child Sys` e `Insert in New Row`.

Nas tentativas via CUA, a linha chegou a aparecer na grelha antes do save, mas
desapareceu após reabrir o mestre da utilizadora e não foi refletida em `AGR_USERS`
no QAD. A tentativa direta via BAPI em QAD foi bloqueada pelo SAP com a mensagem:

```text
Modificações neste sist.não autorizadas (atualiz.usuário central ativa)
```

Estado final da etapa:

- role `Z_PROJECT_APPROVE_SPEC` criada em QAD;
- atribuição `S75 / S4QCLNT100 / Z_PROJECT_APPROVE_SPEC` ainda não ativa;
- a atribuição deve ser concluída no CUA após validação do motivo pelo qual o
  sistema central não persiste esta role recém-criada na grelha de funções.

### Remoção de duplicados `Z_MY_HOME` em QAS (15/09/2026)

Após a etapa `CUA_ADICIONAR`, foi executada uma auditoria RFC em `AGR_USERS` para
os utilizadores de `Purchase & Services`, procurando duplicados dentro do mesmo
sistema:

- PRD (`S4PCLNT100`): 308 registos lidos, 308 pares utilizador/função únicos, 0
  duplicados;
- QAS (`S4QCLNT100`): 385 registos lidos, 380 pares únicos, 5 pares duplicados.

Os duplicados encontrados em QAS eram todos da função `Z_MY_HOME`, com duas
ocorrências diretas por utilizador:

- ocorrência antiga: `FROM_DAT = 20260912`, `TO_DAT = 99991231`;
- ocorrência nova: `FROM_DAT = 20260915`, `TO_DAT = 99991231`.

Critério aplicado, validado antes da execução:

1. manter a ocorrência ativa mais antiga quando ambas têm o mesmo `TO_DAT`;
2. remover a ocorrência mais recente (`FROM_DAT = 20260915`);
3. localizar a linha exata no CUA por `SUBSYSTEM + AGR_NAME + UPDATE_FROM_DAT +
   UPDATE_TO_DAT`;
4. confirmar após cada gravação por RFC no QAD.

Foram removidas 5 ocorrências em `S4QCLNT100`:

- `S170 / Z_MY_HOME / FROM_DAT 20260915`;
- `S270 / Z_MY_HOME / FROM_DAT 20260915`;
- `S419 / Z_MY_HOME / FROM_DAT 20260915`;
- `S80000148 / Z_MY_HOME / FROM_DAT 20260915`;
- `S965 / Z_MY_HOME / FROM_DAT 20260915`.

Cada gravação devolveu status SAP `S` com mensagem `User <utilizador> has changed`.
A verificação RFC final confirmou que cada um dos cinco utilizadores ficou com
apenas uma ocorrência ativa de `Z_MY_HOME`, mantendo `FROM_DAT = 20260912` e
`TO_DAT = 99991231`.

## 19. Preparação CUA_ADICIONAR — Client Services (15/09/2026)

Após marcar `Purchase & Services` como `PROCESSADO`, o próximo departamento pendente
na folha `CONTROLO` passou a ser `Client Services`.

Foi identificada e corrigida uma inflação na análise inicial: o modelo expandido
contava também as funções filhas internas de `PFCG_COMPOSTA`, gerando 209 faltas em
QAS. Para a etapa `CUA_ADICIONAR`, o critério correto é atribuição direta:

- Composite Role da coluna `Composite Role` da folha `Proposta Ativa`;
- funções base da folha `DEFINIÇÕES`;
- não incluir como atribuições diretas as funções filhas internas da composta.

Correções aplicadas em `Projeto Perfil.py`:

- nova validação `is_role_sap_valida`, evitando tratar descrições/cargos como roles
  SAP;
- nova extração `extrair_roles_sap`, separando funções em células com vírgulas,
  ponto-e-vírgula ou quebras de linha;
- leitura da folha `DEFINIÇÕES` para consolidar roles base por departamento;
- expansão departamental corrigida para incluir `DEFINIÇÕES` sem transformar as
  filhas de compostas em payload direto de CUA.

Resultado operacional da correção na folha `CUA_ADICIONAR`:

- removidas 432 linhas previamente preparadas pelo critério expandido;
- recriadas 120 linhas diretas para `Client Services`:
  - 60 linhas PRD (`S4PCLNT100`) já preenchidas/validadas;
  - 7 linhas QAS (`S4QCLNT100`) já preenchidas/validadas;
  - 53 linhas QAS (`S4QCLNT100`) pendentes para execução CUA;
- PRD ficou sem pendências diretas;
- QAS ficou com 53 pendências diretas, distribuídas por utilizador:
  - `S5092`: 9;
  - `S5441`: 9;
  - `S5877`: 9;
  - `S80000647`: 8;
  - `S80000781`: 9;
  - `S80001601`: 9.

Foi criada uma cópia do Excel mestre na raiz do projeto:

```text
C:\workspace\SapScript\S4H_Perfis de autorização_v1.xlsx
```

Backup criado antes da correção da `CUA_ADICIONAR`:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_fix_client_services_cua_adicionar_20260915_175201.xlsx
```

## 20. Execução CUA_ADICIONAR e auditoria pós-CUA — Client Services (15/09/2026)

As 53 pendências diretas em QAS (`S4QCLNT100`) foram executadas via SAP GUI/CUA e
confirmadas por RFC. Apesar de o executor devolver `AVISO` por não obter mensagem
conclusiva na status bar, a validação em `AGR_USERS` confirmou todas as 53
atribuições. A folha `CUA_ADICIONAR` foi atualizada com:

- `STATUS = CONCLUÍDO`;
- `MSG = Atribuição criada/confirmada no SAP QAD via RFC`;
- `QAS = OK`;
- `TIMESTEMP` da execução.

Backup criado antes da execução:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_exec_client_services_cua_qas_20260915_191746.xlsx
```

Na auditoria pós-CUA foi confirmada a ausência de faltas diretas:

- PRD / `S4PCLNT100`: `0` faltas diretas e `0` candidatos a `CUA_REMOVE`;
- QAS / `S4QCLNT100`: `0` faltas diretas e `0` candidatos a `CUA_REMOVE`, após
  correção do catálogo de compostas.

Durante a comparação inicial apareceram 32 funções como “extra” em QAS, mas todas
tinham `COL_FLAG = X`, ou seja, eram herdadas por função composta. Foi validado via
`AGR_AGRS` que essas funções pertencem às compostas `Z_BR_CLIENTSERV_SPECIALIST` e
`Z_BR_CLIENTSERV_TEAMLEAD`. A folha `PFCG_COMPOSTA` foi então complementada sem
sobrepor valores existentes:

- 24 relações novas adicionadas;
- 96 células vazias preenchidas em `STATUS`, `MSG`, `TIMESTEMP`, `PRD` e/ou `QAS`;
- QAS passou a refletir 66 relações para as compostas de Client Services;
- PRD reflete 42 relações para as mesmas compostas.

Backup criado antes da atualização da `PFCG_COMPOSTA`:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_pfcg_composta_client_services_20260915_193755.xlsx
```

Após aprovação explícita, foram removidos 53 duplicados diretos em QAS pelo SAP
GUI/CUA, mantendo a ocorrência com maior validade (`TO_DAT` maior; em empate,
`FROM_DAT` mais recente). Para estes casos foi mantido `FROM_DAT = 20260915` e
removido `FROM_DAT = 20260912`.

Relatório gerado:

```text
C:\workspace\SapScript\output\client_services_qas_duplicados_removidos_20260915_194346.csv
```

Validação final por RFC:

- PRD / `S4PCLNT100`: `0` faltas diretas, `0` candidatos a `CUA_REMOVE`, `0`
  duplicados diretos e `0` duplicados totais;
- QAS / `S4QCLNT100`: `0` faltas diretas, `0` candidatos a `CUA_REMOVE`, `0`
  duplicados diretos e `7` duplicados totais herdados (`COL_FLAG = X`), que não
  devem ser removidos diretamente pelo CUA.

A linha 3 da folha `CONTROLO` (`Client Services`) foi marcada como:

- `STATUS = PROCESSADO`;
- `TIMESTEMP = 2026-09-15 19:44:45`;
- `MSG = Processado com sucesso; validação final PRD/QAS OK`.

Backup criado antes da atualização da `CONTROLO`:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_controlo_client_services_20260915_194443.xlsx
```

Próximos departamentos pendentes na folha `CONTROLO`:

1. `Construction & Maintenance`;
2. `People & Talent`;
3. `Health & Safety`.

## 21. Início do departamento Construction & Maintenance (15/09/2026)

O próximo departamento pendente identificado na folha `CONTROLO` foi
`Construction & Maintenance`, com ambientes `QAS` e `PRD` marcados.

Escopo identificado na `Proposta Ativa`:

- 9 utilizadores;
- 9 funções base vindas de `DEFINIÇÕES`;
- 8 funções compostas distintas;
- payload direto por utilizador: 10 funções (9 base + 1 composta).

Utilizadores:

- `S105` — `Z_BR_CONSTMANUT_SPECIALIST`;
- `S13020` — `Z_BR_STOREMAINT_SPECIALIST`;
- `S13360` — `Z_BR_STOREMAINT_SPECIALIST`;
- `S4244` — `Z_BR_CONSTMAINT_MANAGER`;
- `S425` — `Z_BR_STOREMAINT_TEAMLEAD`;
- `S5006` — `Z_BR_CONSTPROJ_MANAGER`;
- `S5354` — `Z_BR_PURCHASEREQ_SPECIALIST`;
- `S6005` — `Z_BR_EXPANSIONPROJ_MANAGER`;
- `S80000721` — `Z_BR_CONSTCONTROLER_SPECIALIST`.

### Remoção do sistema legado S4DCLNT100

Foi executado primeiro um `dry-run` no SAP CUA (`SPA`, mandante `001`), que
confirmou que os 9 utilizadores ainda tinham o sistema `S4DCLNT100`.

Após confirmação operacional do fluxo, foi executada a remoção real via SAP GUI/CUA.
Resultado:

- 9 utilizadores processados;
- 9 concluídos com sucesso;
- todos ficaram apenas com `S4PCLNT100` e `S4QCLNT100`.

### Validação PFCG/Compostas

PRD / `S4PCLNT100` foi validado por RFC:

- 17 funções diretas existem em `AGR_DEFINE`;
- 8 compostas existem;
- 102 relações de composta encontradas em `AGR_AGRS`;
- nenhuma função direta em falta.

QAS / `S4QCLNT100` ficou pendente por indisponibilidade RFC. A ligação ao destino
`172.19.66.22:3300` falhou repetidamente com `WSAETIMEDOUT`, inclusive fora do
sandbox. Nenhuma validação QAS foi assumida por inferência.

### Auditoria PRD de atribuições atuais

PRD ficou sem faltas diretas de `CUA_ADICIONAR` para os 9 utilizadores.

Pontos encontrados antes de qualquer alteração:

- `S6005 / S4PCLNT100`: função extra candidata a avaliação para `CUA_REMOVE`:
  `Z_PURCHASE_ORDER_DISPLAY`;
- `S4244 / S4PCLNT100`: 2 duplicados diretos ativos:
  - `ZMM_APROVA_PEDC_COD_JO10`: `20260914-99991231` e `20260912-99991231`;
  - `ZMM_APROVA_PEDC_COD_L210`: `20260914-99991231` e `20260912-99991231`.

Nenhuma remoção PRD foi executada ainda; por ser destrutiva, requer confirmação
explícita do critério e autorização de execução.

Após autorização explícita, foram removidos em PRD os 2 duplicados diretos de
`S4244`, mantendo as ocorrências com `FROM_DAT = 20260914` e removendo as
ocorrências antigas com `FROM_DAT = 20260912`:

- `ZMM_APROVA_PEDC_COD_JO10`;
- `ZMM_APROVA_PEDC_COD_L210`.

Relatório gerado:

```text
C:\workspace\SapScript\output\construction_prd_remocoes_20260915_224353.csv
```

A função `Z_PURCHASE_ORDER_DISPLAY` em `S6005` não era atribuição direta
removível: o registo em `AGR_USERS` tinha `COL_FLAG = X`, portanto era herdado
por composta. A relação `Z_BR_EXPANSIONPROJ_MANAGER -> Z_PURCHASE_ORDER_DISPLAY`
foi adicionada à folha `PFCG_COMPOSTA`, sem sobrepor dados já existentes.

Backup criado antes da atualização:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_pfcg_composta_construction_prd_20260915_224444.xlsx
```

Validação final PRD após correções:

- faltas diretas: `0`;
- candidatos a `CUA_REMOVE`: `0`;
- duplicados diretos: `0`.

QAS continua pendente por indisponibilidade RFC (`WSAETIMEDOUT` no destino
`172.19.66.22:3300`).

## 22. Início do departamento People & Talent (15/09/2026)

O próximo departamento operacional processado foi `People & Talent`. `Construction
& Maintenance` permanece sem marcação final `PROCESSADO` porque QAS continuou
pendente por timeout RFC.

Escopo identificado na `Proposta Ativa`:

- 12 utilizadores;
- 9 funções base vindas de `DEFINIÇÕES`;
- 3 funções compostas distintas:
  - `Z_BR_PURCHREQ&PO_MANAGER`;
  - `Z_BR_PURCHREQ&PO_SPECIALIST`;
  - `Z_BR_PURCHREQ&PO_TEAMLEAD`;
- payload direto por utilizador: 10 funções (9 base + 1 composta).

### Remoção do sistema legado S4DCLNT100

Foi executado `dry-run` no SAP CUA (`SPA`, mandante `001`):

- 10 utilizadores ainda tinham `S4DCLNT100`;
- 2 utilizadores já não tinham o sistema (`S80001028`, `S80001500`).

Foi executada a remoção real:

- 10 remoções concluídas;
- 2 casos `NAO_EXISTIA`, tratados como OK;
- os utilizadores processados ficaram sem o sistema legado `S4DCLNT100`.

### CUA_ADICIONAR PRD

A auditoria inicial PRD apontou 108 atribuições diretas em falta (12 utilizadores
x 9 funções base/compostas em falta; `Z_MY_HOME` já existia).

As 108 linhas foram criadas na folha `CUA_ADICIONAR` e executadas via SAP GUI/CUA
em lotes. A saída do executor ficou silenciosa durante parte da execução, por isso
a reconciliação oficial foi feita por RFC:

- 108 atribuições confirmadas em PRD;
- 0 pendências diretas PRD após reconciliação;
- as 108 linhas foram marcadas como `CONCLUÍDO`, `PRD = OK` e mensagem
  `Atribuição criada/confirmada no SAP PRD via RFC`.

Backups relevantes:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_people_cua_prd_20260915_225216.xlsx
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_people_cua_prd_20260915_225645.xlsx
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_people_cua_prd_20260915_225954.xlsx
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_reconcile_people_cua_prd_20260915_230130.xlsx
```

### Sincronização PFCG_COMPOSTA PRD

Após ativar as compostas, várias funções-filhas apareceram como “extras” por falta
de catálogo local. Foram lidas por RFC as relações PRD em `AGR_AGRS` para:

- `Z_BR_PURCHREQ&PO_MANAGER`;
- `Z_BR_PURCHREQ&PO_SPECIALIST`;
- `Z_BR_PURCHREQ&PO_TEAMLEAD`.

Resultado:

- 26 relações PRD adicionadas à folha `PFCG_COMPOSTA`;
- 104 células vazias preenchidas em `STATUS`, `MSG`, `TIMESTEMP` e `PRD`;
- dados existentes não foram sobrepostos.

Backup:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_sync_people_compostas_prd_20260915_230231.xlsx
```

### Estado PRD após inclusão e sincronização de catálogo

Validação PRD:

- faltas diretas: `0`;
- funções diretas existem em `AGR_DEFINE`: OK;
- compostas PRD:
  - `Z_BR_PURCHREQ&PO_MANAGER`: 9 filhas;
  - `Z_BR_PURCHREQ&PO_SPECIALIST`: 8 filhas;
  - `Z_BR_PURCHREQ&PO_TEAMLEAD`: 9 filhas.

Pendências PRD antes de qualquer remoção:

- funções extras antigas ainda candidatas a análise para `CUA_REMOVE` em vários
  utilizadores;
- 11 duplicados diretos em `S538`, todos com `FROM_DAT 20220715` e `20230529`:
  - `ZORG_EMPRESA_2010`;
  - `ZORG_EMPRESA_2020`;
  - `ZORG_EMPRESA_2070`;
  - `ZORG_EMPRESA_2080`;
  - `ZORG_EMPRESA_2100`;
  - `ZORG_EMPRESA_2110`;
  - `ZORG_EMPRESA_2120`;
  - `ZORG_EMPRESA_2130`;
  - `ZORG_EMPRESA_2140`;
  - `ZORG_EMPRESA_2150`;
  - `ZORG_EMPRESA_2160`.

Nenhuma remoção PRD de extras/duplicados foi executada nesta etapa; requer
aprovação explícita por ser alteração destrutiva no SAP.

QAS continua pendente por indisponibilidade RFC (`WSAETIMEDOUT` no destino
`172.19.66.22:3300`).

## 23. Início e conclusão do departamento Health & Safety (15/09/2026)

O próximo departamento processado foi `Health & Safety` (Linha 6 da folha `CONTROLO`).

Escopo identificado na `Proposta Ativa`:

- 2 utilizadores:
  - `S80001882` (Marta Ramos — Health & Safety Specialist);
  - `S80001974` (Joana Saldanha — Executive Office Assistant);
- 9 funções base vindas de `DEFINIÇÕES`:
  - `Z_BASIS_BASE`
  - `Z_MY_HOME` (já ativa em ambos no PRD)
  - `ZORG_BP_FLVN01_LOGISTICS_VENDO`
  - `ZORG_BP_Z001_GENERALPARTNERS`
  - `ZORG_BP_GERAL`
  - `ZORG_BP_Z003_RELATEDPARTNERS`
  - `ZORG_BP_LOGISTICS_CUSTOMER`
  - `Z_BR_TYPE_BP_GERAL`
  - `ZORG_TODAS_EMPRESAS`
- 1 função composta atribuída a ambos: `Z_BR_PURCHREQ&PO_SPECIALIST`;
- payload direto por utilizador: 10 funções (9 base + 1 composta; 9 novas por utilizador).

### Remoção do sistema legado S4DCLNT100

Foi executado dry-run e execução real no SAP CUA (`SPA`, mandante `001`) via `K. CUA_REMOVE_SISTEMA.py`:

- `S80001882`: sistema legado `S4DCLNT100` removido com sucesso via SAP GUI (restantes: `S4PCLNT100`, `S4QCLNT100`);
- `S80001974`: sistema `S4DCLNT100` já não existia no utilizador (apenas `S4PCLNT100`).

### CUA_ADICIONAR PRD

Foram criadas 18 linhas de atribuição direta na folha `CUA_ADICIONAR` (IDs 504 a 521), correspondendo às 9 funções em falta para cada um dos 2 utilizadores em `S4PCLNT100`.

A execução foi realizada via SAP GUI Scripting no SAP CUA. A reconciliação oficial por RFC no SAP PRD confirmou:

- `S80001882`: 24/24 funções ativas do plano (**100% CONFORME**);
- `S80001974`: 24/24 funções ativas do plano (**100% CONFORME**);
- faltas diretas: `0`;
- as linhas da folha `CUA_ADICIONAR` foram marcadas como `CONCLUÍDO`, `PRD = OK` e mensagem `Atribuição criada/confirmada no SAP PRD via RFC`.

Backups gerados:

```text
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_health_cua_prd_20260915_233215.xlsx
C:\workspace\SapScript\output\S4H_Perfis_autorizacao_v1_before_reconcile_health_cua_prd_20260915_234120.xlsx
```

### Funções adicionais detetadas no SAP PRD

- `S80001882`: 14 adicionais legados (`ZORG_CENTROS_2XXX`, `ZORG_EMPRESA_2010..2160`, `Z_CROSS_LOGISTIC_SLS`, `Z_LOGISTIC_TEMP_SLS`);
- `S80001974`: 18 adicionais legados (`ZFIN_AR_BASIC`, `ZFIN_DADOS_MESTRE_BASIC`, `ZHR_COLAB_BASIC`, `ZMM_CRIA_PEDC`, `ZMM_CRIA_REQC`, `ZMM_SUPPLY_CHAIN_BASIC`, `ZORG_EMPRESA_2010..2160`, `Z_CROSS_LOGISTIC_ALL`).

Nenhuma remoção de acessos legados foi executada sem prévia aprovação explícita.

QAS continua pendente por indisponibilidade RFC (`WSAETIMEDOUT` no destino `172.19.66.22:3300`).
