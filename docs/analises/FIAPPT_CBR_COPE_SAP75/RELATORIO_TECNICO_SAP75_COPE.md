# Relatório Técnico de Investigação: FIAPPT_CBR & CL_FIAPPT_CBR (COPE / Balança Comercial)

- **Tickets Associados**: [IZ-52928](https://salsajeans.atlassian.net/browse/IZ-52928) & [SAP-75](https://salsajeans.atlassian.net/browse/SAP-75)
- **Data da Extração**: 14/10/2026
- **Ambientes Analisados via RFC**: `QAD` (Qualidade) e `PRD` (Produção)
- **Autor do Fix**: Marinho Pereira (`MPEREIRA`)
- **Ordem de Transporte**: `S4DK953607` (Tarefa `S4DK953608`)
- **Data da OT**: 04.08.2026

---

## 1. Localização dos Ficheiros de Código-Fonte Extraídos

O código-fonte integral de todos os programas, includes, classes e módulos de função envolvidos foi descarregado diretamente via RFC do SAP e arquivado em:
- **QAD**: [`docs/analises/FIAPPT_CBR_COPE_SAP75/QAD/`](file:///C:/workspace/SapScript/docs/analises/FIAPPT_CBR_COPE_SAP75/QAD/)
- **PRD**: [`docs/analises/FIAPPT_CBR_COPE_SAP75/PRD/`](file:///C:/workspace/SapScript/docs/analises/FIAPPT_CBR_COPE_SAP75/PRD/)

Total de ficheiros extraídos por ambiente: **83 ficheiros ABAP**.

---

## 2. Anatomia do Problema & Causa Raiz

### O Problema Reportado no Jira (IZ-52928 / Cátia Eira)
1. **Determinação incorreta de código BC**: Na execução do relatório COPE (`FIAPPT_CBR`), pagamentos com a conta de razão **`27112000`** (*Fornecedores Invest-Ativ Fixos Tang-Contas Gerais*) estavam a determinar o código de Balança Comercial **`E2010`** em vez de **`E2020`**.
2. **Rejeição / Falta de preenchimento**: Em certos casos, o código do Banco de Portugal vinha completamente vazio ou era rejeitado na validação do ficheiro.

### O Diagnóstico Técnico do Marinho Pereira (SAP-75)
- Na geração dos registos de saldo e transações do COPE pela classe standard **`CL_FIAPPT_CBR`**, o tipo de conta bancária (`BANK_ACC_TYPE`) vinha preenchido com **`E`** (ou valores diferentes de `X`).
- Segundo a validação do Banco de Portugal para o layout do COPE, o campo `BANK_ACC_TYPE` apenas permite o valor **`X`**.

---

## 3. O que a OT `S4DK953607` alterou no SAP

Através da inspeção RFC nas tabelas do Change & Transport System (`E070` e `E071`) e no repositório ABAP (`REPOSRC`), identificámos os 4 objetos transportados:

| Objeto | Tipo | Nome | Descrição |
|---|---|---|---|
| `ENHO` | Enhancement Implementation | **`ZFI_ENH_CHANGE_ACC_TYPE`** | Ponto de extensão no relatório/classe COPE para forçar o tipo de conta |
| `ENHO` | Enhancement Implementation | **`ZFI_ENH_CHANGE_BANK_ID_3`** | Ponto de extensão para ajuste do Bank ID e País da Conta |
| `FUGR` | Function Group | **`ZFI_CHANGE_BANKID_FG`** | Grupo de funções que suporta as regras customizadas de COPE |
| `FUNC` | Function Module | **`ZFI_CHANGE_ACC_TYPE`** | Módulo de função que itera a tabela e força `BANK_ACC_TYPE = 'X'` |

### Código Implementado em QAD:

#### 1. Módulo de Função `ZFI_CHANGE_ACC_TYPE` (em `LZFI_CHANGE_BANKID_FGU03.abap`):
```abap
FUNCTION ZFI_CHANGE_ACC_TYPE.
*"----------------------------------------------------------------------
*"  CHANGING
*"     REFERENCE(ET_FIAPPTD_BALANCE) TYPE  FIAPPT_T_BALANCE
*"----------------------------------------------------------------------

LOOP AT ET_FIAPPTD_BALANCE ASSIGNING FIELD-SYMBOL(<FS>) WHERE BANK_ACC_TYPE <> 'X'.
  <FS>-BANK_ACC_TYPE = 'X'.
ENDLOOP.

ENDFUNCTION.
```

#### 2. Enhancement `ZFI_ENH_CHANGE_ACC_TYPE` (em `ZFI_ENH_CHANGE_ACC_TYPE=======E.abap`):
```abap
ENHANCEMENT 1 .

CALL FUNCTION 'ZFI_CHANGE_ACC_TYPE'
  CHANGING
    ET_FIAPPTD_BALANCE = ET_FIAPPTD_BALANCE.

ENDENHANCEMENT.
```

---

## 4. Comparação entre QAD e PRD

- **Em QAD (`S4Q`)**: A OT `S4DK953607` já foi importada. O módulo de função `ZFI_CHANGE_ACC_TYPE` e as enhancements `ZFI_ENH_CHANGE_ACC_TYPE` e `ZFI_ENH_CHANGE_BANK_ID_3` estão **ativos**.
- **Em PRD (`S4P`)**: A OT ainda **não existe**. O módulo de função `ZFI_CHANGE_ACC_TYPE` não existe em PRD.

---

## 5. Mapeamento da Conta 27112000

- **Designação da Conta (`SKAT`)**: `0027112000` — *Fornecedores Invest-Ativ Fixos Tang-Contas Gerais* (Plano de Contas `1000` / `GC01`).
- **Tabela de Custos/Saldos SCB (`FIAPPTD_BAL_SCB`)**:
  - Foi inspecionada a tabela inteira (232 registos).
  - A tabela possui parametrizações para a empresa `2020` e `2010`.
  - A conta `27112000` é uma conta de Fornecedores de Imobilizado. No relatório standard `FIAPPT_CBR`, se a linha for tratada como transação de fornecedor/G/L com conta `27112000`, o tipo de conta bancária deixava de ser `X`, provocando o erro de validação.
  - A correção agora força `BANK_ACC_TYPE = 'X'` em todas as linhas de balanço geradas antes da criação do XML para o Banco de Portugal.

---

## 6. Procedimento Recomendado para Teste no SAP GUI (S4Q)

1. Abrir sessão no ambiente **`S4Q`**.
2. Executar a transação **`FIAPPT_CBR`**.
3. Parâmetros de Seleção:
   - Empresa: `1000` (ou a empresa do pagamento da Cátia Eira, ex: `2010` / `2020`).
   - Exercício: Ano do pagamento (ex: `2025` / `2026`).
   - Período: Mês correspondente.
   - Flag: Ativar **Modo de Teste** (para não gerar numeração definitiva do Banco de Portugal).
4. Executar e validar no ALV gerado:
   - Se os documentos com a conta **`27112000`** apresentam agora o código de BC correto (`E2020`).
   - Se o campo `Tipo de Conta` / `BANK_ACC_TYPE` está preenchido com **`X`**.
   - Se a criação do XML do COPE decorre sem erros de validação.
