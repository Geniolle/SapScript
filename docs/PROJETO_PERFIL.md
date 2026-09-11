# Documentação: Projeto Perfil (`Projeto Perfil.py`)

## 1. Visão Geral
O script `Projeto Perfil.py` foi criado para consolidar, analisar e pesquisar perfis de autorização SAP (PFCG) e atribuições no CUA a partir de folhas de cálculo Excel.

## 2. Integração com SharePoint e OneDrive
O script conecta-se diretamente ao repositório corporativo no SharePoint via sincronização local do OneDrive (configurada no `.env`):
- `SHAREPOINT_PERFIS_URL`: URL oficial da pasta no SharePoint (`006. Perfis Autorização`).
- `SHAREPOINT_PERFIS_LOCAL_DIR`: Caminho da pasta sincronizada localmente no Windows via OneDrive.
- `SHAREPOINT_PERFIS_FILE`: Caminho completo do ficheiro mestre `S4H_Perfis de autorização.xlsx`.

### Leitura Concorrente (Sem Bloqueio do Excel)
Graças à função `abrir_excel_seguro`, o script abre o ficheiro utilizando a API do Windows (`win32file.CreateFile`) com as flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`. Isso permite executar pesquisas e análises mesmo enquanto o ficheiro estiver aberto e em edição no Microsoft Excel.

## 3. Fluxo de Validação de Departamentos

### Etapa 1: Sheet `CONTROLO`
- Analisa a coluna `STATUS` de cada departamento registado na sheet `CONTROLO`.
- Identifica departamentos com `STATUS` vazio/pendente (ex.: `Purchase & Services`).
- Verifica se existe a sheet correspondente com o detalhe das transações do departamento.
- Comando:
  ```powershell
  python "Projeto Perfil.py" --controlo
  ```

### Etapa 2: Sheet `Proposta Ativa`
- Filtra os dados da sheet `Proposta Ativa` pelo departamento pretendido.
- Extrai e cruza:
  - Lista de utilizadores, nomes e cargos.
  - Funções Compostas atribuídas (ex.: `Z_BR_PURCHSERV_MANAGER`, `Z_BR_PURCHSERV_SPECIALIST`, `Z_BR_PURCHSERV_TEAMLEAD`).
  - Funções Individuais (Single Roles) associadas aos utilizadores, com contagem de frequência de atribuição.
- Comando:
  ```powershell
  python "Projeto Perfil.py" --departamento "Purchase & Services"
  ```

## 4. Funcionalidades do Motor de Pesquisa
- **Pesquisa por Função / Perfil**:
  ```powershell
  python "Projeto Perfil.py" --pesquisar-role EXPANSION
  ```
- **Pesquisa por Transação (TCODE)**:
  ```powershell
  python "Projeto Perfil.py" --pesquisar-tcode FB03
  ```
- **Pesquisa por Utilizador (CUA)**:
  ```powershell
  python "Projeto Perfil.py" --pesquisar-user S6005
  ```
- **Menu Interativo de Consola**:
  ```powershell
  python "Projeto Perfil.py"
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

- As 640 associações composta–individual previstas existem em `AGR_AGRS`.
- As 24 funções compostas correspondem exatamente ao Excel, sem membros em falta ou adicionais.

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
python "Projeto Perfil.py" --cruzar-fontes -d "Purchase & Services"

# Validação real no SAP PRD (AGR_USERS & USR02) com expansão relacional
python "Projeto Perfil.py" --validar-users-prd -d "Purchase & Services"
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
python "Projeto Perfil.py" --comparar-prd
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

