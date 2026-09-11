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

### Atribuições aos utilizadores ativos

O departamento em aberto na folha `CONTROLO` era `Purchase & Services`.

- 7 utilizadores constavam da proposta.
- 6 utilizadores ativos estavam conformes: `S170`, `S270`, `S419`, `S75`, `S80000148` e `S965`.
- As 210 atribuições esperadas para esses utilizadores estavam ativas em `AGR_USERS`.
- `S80001870` foi excluído da divergência porque a validade no mestre `USR02` terminou em **29/07/2026**; o último logon também ocorreu nessa data.

Conclusão: as funções e os utilizadores ativos do departamento foram validados para a etapa de atribuição.

### Scripts de auditoria

- `scratch/validar_tcodes_perfil_prd.py`
- `scratch/validar_membros_compostas_prd.py`
- `scratch/validar_utilizadores_departamentos_abertos_prd.py`
- `scratch/consultar_mestre_utilizador_prd.py`

Todos os scripts acima executam apenas consultas RFC de leitura.
