# AGENTS.md — SAP Script Workspace

## Contexto do Projeto

Este workspace contém scripts Python para automação de processos SAP S/4HANA,
especificamente para gestão de roles PFCG e utilizadores CUA.

O ponto de entrada principal é o **SAP Cockpit** (`SAP Cockpit.py`), que lista
e executa os sub-scripts organizados por processo.

---

## Arquitectura de Scripts

### Entry Point
- `SAP Cockpit.py` — Menu interactivo que descobre e executa sub-processos
  automaticamente por pasta. Aceita argumento de ambiente: `DEV`, `QAD`, `PRD`.

### Processos (pasta `Processos/`)
Cada subpasta corresponde a um menu no Cockpit. Os scripts são descobertos automaticamente.

#### `Processos/Funções PFCG/`

| Script | Método | Descrição |
|--------|--------|-----------|
| `A. PFCG_CREATE.py` | GUI (SAP Scripting) | Criar roles simples via PFCG GUI |
| `A. PFCG_CREATE_RFC.py` | **RFC** | Criar/actualizar roles simples via RFC (sem GUI) |
| `B. PFCG_DELETE.py` | GUI | Eliminar roles |
| `C. PFCG_AUTHORITY.py` | GUI | Gerir objectos de autorização |
| `D. PFCG_COMPOSTA.py` | GUI (SAP Scripting) | Criar roles compostas via PFCG GUI |
| `D. PFCG_COMPOSTA_RFC.py` | **RFC** | Criar/actualizar roles compostas via RFC (sem GUI) |
| `H. CUA_ADICIONAR.py` | GUI | Adicionar utilizadores CUA |
| `I. CUA_ENDDATE.py` | GUI | Alterar data de fim de roles CUA |
| `J. CUA_REMOVE.py` | GUI | Remover roles de utilizadores CUA |

---

## Padrão RFC (sem GUI SAP)

Os scripts `*_RFC.py` são a versão optimizada que **não requer GUI SAP aberta**.
Comunicam directamente com o servidor SAP via `pyrfc` (SAP NW RFC SDK).

### Credenciais RFC (.env)
```
SAP_ASHOST=<host>
SAP_SYSNR=00
SAP_USER=<user>
SAP_PASSWORD_S4DCLNT100=<password>
```

### Mapeamento de Ambientes
```python
MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
```

---

## Ficheiro Excel de Dados
O ficheiro `S4H_Perfis de autorização.xlsx` é o input principal.
Contém múltiplas sheets, uma por processo:

| Sheet | Processo |
|-------|----------|
| `PFCG_CREATE` | Criar roles simples (TCODEs) |
| `PFCG_COMPOSTA` | Criar roles compostas (roles filhas) |
| `CUA_ADICIONAR` | Adicionar utilizadores |

### Colunas obrigatórias (todas as sheets)
- `AGR_NAME` — Nome da role
- `TEXT` — Descrição da role  
- `STATUS` — Estado de processamento (preenchido pelo script)
- `MSG` — Mensagem de resultado
- `TIMESTEMP` — Timestamp de execução

---

## Convenções de Código

### Gravação Excel
Sempre usar `win32com.client` (COM Automation) como método principal,
com fallback para `openpyxl`. Isto evita corrupção de ficheiros Excel.

```python
def gravar_resultados_excel(caminho, sheet, header_map, records, resultados):
    # 1. Tentar COM
    # 2. Fallback: openpyxl
```

### Logging de Progresso
Usar `rich.progress` para barra de progresso visual:
```python
from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn
```

### Encoding
Todos os scripts devem começar com:
```python
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
```
E correr com `.venv/Scripts/python.exe -X utf8`.

### Checkpoint Excel
Gravar resultado no Excel **após cada role**, não só no final.
Isto garante que progresso não se perde se o script for interrompido.

---

## Skill de Referência
Ver `.agents/skills/sap_rfc_automation/SKILL.md` para a documentação
completa das RFCs SAP utilizadas e padrões de implementação.
