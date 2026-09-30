# Relatório Técnico: Investigação Somente Leitura - Manutenção de Condições de Pagamento (OBB8 / T052 / V_T052) no SAP DEV (S4D)

**Data da Análise:** 30/09/2026  
**Ambiente:** SAP DEV (S4D / Host: 172.19.66.4 / Client 100)  
**Modo:** Somente Leitura (Read-Only)  
**Projeto:** `SapScript` (`C:\workspace\SapScript`)

---

## 1. Como Funciona a Transação OBB8

A transação **OBB8** não é um programa de diálogo ABAP isolado ou customizado, mas sim uma **Transação de Parâmetro (Parameter Transaction)** standard do SAP que executa diretamente a transação de manutenção de views `SM30`.

### Mapeamento Técnico de Invocação:
- **Tabela `TSTC`:**
  - `TCODE`: `OBB8`
  - `PGMNA`: `SDTM` (Chamador standard de transações de parâmetros)
- **Tabela `TSTCP`:**
  - `PARAM`: `/*SM30 VIEWNAME=V_T052;UPDATE=X;`

Quando o utilizador executa `OBB8`, o SAP chama a função `VIEW_MAINTENANCE_CALL` (ou `VIEW_MAINTENANCE_SINGLE_ENTRY`) do grupo de funções `SVIM` (`SAPLSVIM`), passando a visão de manutenção `V_T052` em modo de atualização (`UPDATE = 'X'`).

---

## 2. Objetos Técnicos Envolvidos

### Grupo de Funções e Programas:
- **Grupo de Funções Principal:** `0F30` (`SAPL0F30` - *Generated View Maintenance Function Pool*).
- **Módulo de Processamento Geral de Views:** Grupo de Funções `SVIM` (`SAPLSVIM`).
- **Includes do Grupo `0F30`:**
  - `L0F30TOP`: Declarações de dados globais da visão.
  - `L0F30UXX`: Declaração dos Function Modules da visão.
  - `L0F30F00`: Subprogramas gerados automaticamente pelo TMG.
  - `L0F30I00`: Módulos PAI (Process After Input) gerados.
  - `LSVIMFXX` / `LSVIMOXX` / `LSVIMIXX`: Módulos genéricos do gerador de visões (PBO/PAI/Subroutines).
  - `L0F30I01`: Módulo PAI customizado pelo SAP standard para gestão de textos multilíngua (tradução de explicações `TEXT1`).

### Function Modules Gerados no Grupo `0F30`:
- `VIEWFRAME_V_T052` / `VIEWPROC_V_T052`
- `VIEWFRAME_V_T052S` / `VIEWPROC_V_T052S`
- `VIEWFRAME_V_TVZBT` / `VIEWPROC_V_TVZBT`
- `TABLEFRAME_0F30` / `TABLEPROC_0F30`

---

## 3. Tabelas e Views

| Nome Técnico | Tipo | Descrição | Chave Primária |
| :--- | :--- | :--- | :--- |
| **`T052`** | Tabela Transparente (Classe C) | Condições de Pagamento (Dados Principais / Cabeçalho) | `MANDT`, `ZTERM`, `ZTAGG` |
| **`T052U`** | Tabela Transparente (Classe C) | Explicação das Condições de Pagamento (Textos Multilíngua) | `MANDT`, `SPRAS`, `ZTERM`, `ZTAGG` |
| **`TVZBT`** | Tabela Transparente (Classe C) | Textos de Condição de Pagamento em Vendas (SD) | `MANDT`, `SPRAS`, `ZTERM` |
| **`V_T052`** | Visão de Manutenção (Maintenance View) | Visão combinada de `T052` + `T052U` para a transação `OBB8`/`SM30` | `MANDT`, `ZTERM`, `ZTAGG` |

---

## 4. BAPIs Encontradas

Foi realizada uma pesquisa exaustiva no repositório de objetos e dicionário do SAP DEV.
- **Resultado:** **Nenhuma BAPI standard** existe para criação/manutenção de Condições de Pagamento (`T052`/`T052U`).
- **Motivo Arquitetural:** No SAP R/3, ECC e S/4HANA, a tabela `T052` é classificada como **Dados de Customizing (Delivery Class C)**. O SAP não disponibiliza BAPIs standard (como `BAPI_PAYMENT_TERMS_CREATE`) para manipulação direta de tabelas de customizing do sistema.

---

## 5. Function Modules Encontrados

Foram inspecionadas as funções relacionadas a `ZTERM`, `T052` e `PAYMENT_TERMS`:

| Nome Técnico | Tipo | Standard/Z | Remote-Enabled | Altera Dados | Descrição / Função |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `VIEW_MAINTENANCE_CALL` | FM | Standard | **Não (Local)** | Sim (via UI/TMG) | Invoca o gerador de manutenção de views em tela. |
| `VIEW_MAINTENANCE_SINGLE_ENTRY` | FM | Standard | **Não (Local)** | Sim (via UI/TMG) | Permite edição/criação de uma entrada única na view. |
| `VIEW_MAINTENANCE_GIVEN_DATA` | FM | Standard | **Não (Local)** | Sim | Gravação direta de dados passados em tabela sem tela. |
| `VIEWPROC_V_T052` | FM | Standard | **Não (Local)** | Sim | Processador interno da visão `V_T052` no grupo `0F30`. |
| `FI_PRINT_ZTERM` | FM | Standard | **Não (Local)** | Não | Formatação de texto da condição de pagamento para impressão. |
| `FI_TEXT_ZTERM` | FM | Standard | **Não (Local)** | Não | Leitura e construção da descrição/texto da condição. |
| `WCB_READ_T052` | FM | Standard | **Não (Local)** | Não | Leitura de `T052`. |
| `FKK_VBKD_ZTERM_CHECK_RFC` | RFC FM | Standard | **Sim (RFC)** | Não | Validação de ZTERM no módulo FI-CA / IS-U (não cria ZTERM). |

---

## 6. Classes / Métodos Encontrados

Pesquisa realizada no Dicionário ABAP (`SEOCLASS` / `SEOCLASSTX`):
- Foram encontradas classes utilitárias para consulta de Business Partner e contratos (ex: `CL_EX_HRAT_INFT0527_RA`, `FSBP_BUPA_MO_BUT052`), porém **nenhuma classe standard pública** existe para criação ou alteração de registros da `T052` / `T052U`.

---

## 7. Objetos Remote-Enabled (RFC)

- **`VIEW_MAINTENANCE_CALL_EXT`**: É o único módulo `VIEW_MAINTENANCE_*` com flag Remote-Enabled (`FMODE = 'R'`), contudo é restrito à integração CRM Order CD2 e não é adequado para manutenção geral da `V_T052`.
- As funções padrão de manutenção de telas (`VIEW_MAINTENANCE_CALL`, `VIEW_MAINTENANCE_SINGLE_ENTRY`) são **locais** e **não podem ser chamadas diretamente via RFC** por clientes externos (Python / `pyrfc`).

---

## 8. Fluxo de Gravação Standard

Quando uma nova condição de pagamento é gravada na OBB8:

```text
OBB8 (Parameter TCode)
  ↓
SM30 (View Maintenance Execution)
  ↓
VIEW_MAINTENANCE_CALL (FG SVIM)
  ↓
PAI / Validation (FG 0F30 - L0F30I01)
  - Validação de chaves ZTERM e ZTAGG
  - Validação de consistência de prazos (ZTAG1 <= ZTAG2 <= ZTAG3)
  - Validação de percentuais de desconto (ZPRZ1, ZPRZ2)
  ↓
TR_SYS_PARAMS / TR_CHECK_TYPE (Transport Organizer)
  - Verificação se o mandante permite alteração (T000)
  - Solicitação/Atribuição de Customizing Request (KORRNUM)
  ↓
UPDATE V_T052 (T052 + T052U) em memória / extract
  ↓
COMMIT WORK (Efetuado pelo controle do TMG ao salvar)
```

---

## 9. Dependências de Transport / Customizing

1. **Classe de Entrega (Delivery Class):** `T052` e `T052U` possuem classe `C` (Customizing).
2. **Transport Request (Order de Transportes):**
   - Em **DEV**, qualquer inclusão/alteração exige a gravação da chave da tabela (`T052` chave: `MANDT + ZTERM + ZTAGG`) numa **Order de Customizing** (Customizing Request).
   - Em **PRD**, as tabelas de customizing são normalmente bloqueadas contra alteração direta (`SCC4` / `T000` client non-modifiable), devendo as alterações ser transportadas de DEV -> QAD -> PRD.

---

## 10. Alternativas Possíveis para Integração RFC

Como não existe BAPI/RFC standard pronta para criar `T052`, foram analisadas as alternativas técnicas para o projeto Python / `SapScript`:

### Alternativa A: RFC Z Wrapper (Recomendada)
- **Requisitos Técnicos:** Criar no SAP DEV um Function Module Z Remote-Enabled (ex: `ZFI_PAYMENT_TERMS_MAINTAIN`).
- **Funcionamento:** O FM Z recebe os parâmetros em estrutura Python, valida os dados, efetua a inclusão controlada nas tabelas `T052` e `T052U` e regista a chave na Order de Customizing através das funções do Transport Organizer (`TR_OBJECT_INSERT` ou `TR_APPEND_TO_COMM_LOCK`).
- **Limitações:** Requer desenvolvimento ABAP em DEV e transporte do objeto Z.
- **Risco:** Baixo, desde que as validações da `T052` sejam replicadas no FM Z.
- **Utilização:** Funcional em DEV/QAD/PRD (com a devida gestão de customizing request).

### Alternativa B: Automação SAP GUI Scripting
- **Requisitos Técnicos:** Execução via `sap_session.py` (Scripting GUI desktop ativo).
- **Funcionamento:** O Python abre a sessão SAP GUI, executa a transação `OBB8`, insere os campos na tela SM30, preenche a pop-up de Order de Transporte e salva.
- **Limitações:** Requer janela SAP GUI aberta e visível no Windows. Não funciona em background puro/headless.
- **Risco:** Mínimo em termos de integridade, pois utiliza 100% da interface standard.
- **Utilização:** DEV (onde a gravação direta com popup de transport é permitida).

---

## 11. Campos Necessários para Criar Z030 (30 Dias Líquidos)

Para criar a condição simples `Z030` (Pagamento a 30 dias líquidos sem desconto):

### Tabela `T052` (Cabeçalho da Condição):
- `MANDT`: `100` (Mandante)
- `ZTERM`: `'Z030'` (Chave de 4 caracteres)
- `ZTAGG`: `00` (Dia limite = 00)
- `ZTAG1`: `30` (Dias do 1º prazo = 30)
- `ZPRZ1`: `0.000` (Percentual de desconto 1 = 0,00%)
- `ZTAG2`: `0`
- `ZPRZ2`: `0.000`
- `ZTAG3`: `0`
- `ZFAEL`: `00` (Dia fixo = 00)
- `ZMONA`: `00` (Meses adicionais = 00)
- `KOART`: `''` (Vazio = válido para todas as contas; ou `'D'`/`'K'`)

### Tabela `T052U` (Descrição / Texto Multilíngua):
- `MANDT`: `100`
- `SPRAS`: `'P'` (Português)
- `ZTERM`: `'Z030'`
- `ZTAGG`: `00`
- `TEXT1`: `'Pagamento 30 dias'` (Texto explicativo)

---

## 12. Conclusão Técnica

1. **Inexistência de BAPI Standard:** O SAP S/4HANA não disponibiliza BAPIs ou RFCs standard prontas para criação de Condições de Pagamento (`T052`/`T052U`), por se tratar de dados de Customizing.
2. **Natureza da OBB8:** A OBB8 é um atalho de parâmetro para a `SM30` sobre a visão `V_T052`. As funções de gravação envolvidas (`VIEW_MAINTENANCE_CALL`, `VIEW_MAINTENANCE_SINGLE_ENTRY`) são **apenas locais** (não-RFC).
3. **Caminho para o Projeto SapScript:**
   - **Cenário Remoto (sem tela):** Exige a criação de uma RFC Customizada Z (`ZFI_PAYMENT_TERMS_MAINTAIN`) em DEV.
   - **Cenário GUI:** Pode ser automatizado via SAP GUI Scripting (`sap_session.py` / `OBB8`).

---
*Relatório gerado automaticamente no âmbito da investigação técnica em modo somente leitura no SAP DEV.*
