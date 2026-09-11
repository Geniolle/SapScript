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
