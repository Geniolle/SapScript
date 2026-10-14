# Relatório de Diagnóstico: Saldos de Contas Bancárias CNY no COPE (FIAPPT_CBR)

**Data:** 14/10/2026  
**Empresa:** 2010 (Irmãos Vila Nova, SA)  
**Processo:** COPE / Balança Comercial (Banco de Portugal)  
**Transação:** `FIAPPT_CBR` / Programa `RFFOPT_CBR` / Classe `CL_FIAPPT_CBR`  
**Referência Jira:** [IZ-52928](https://salsajeans.atlassian.net/browse/IZ-52928) / [SAP-75](https://salsajeans.atlassian.net/browse/SAP-75)

---

## 1. Problema Reportado
Na extração mensal do relatório COPE, as contas bancárias devem refletir a soma de:
- **Conta Firme:** `...10` (ex.: `12040010` para CNY / `12010710` para EUR)
- **Conta Movimento / Transitória:** `...11` (ex.: `12040011` para CNY / `12010711` para EUR)

Nas contas em EUR (`CBK01`, `CGD02`), os saldos sempre foram reportados corretamente.  
Nas contas em CNY (`ICB01` - Industrial and Commercial Bank of China):
1. O relatório apresentava apenas o saldo da conta movimento `12040011` (ex.: Período 2: `-203.962,00`, Período 6: `21.091,06`, Período 7: `16.570,00`).
2. A conta firme `12040010` (que continha saldo acumulado positivo em todos os períodos) era completamente ignorada.
3. No Período 3 (onde a conta movimento tinha saldo zero), a moeda CNY desaparecia integralmente do relatório COPE.

---

## 2. Investigação Técnica

### 2.1 Análise do Código ABAP (`CL_FIAPPT_CBR->GET_BANK_BALANCES_2`)
O método `GET_BANK_BALANCES_2` (`CM00B_CM00B.abap`) efetua a seguinte lógica para cada conta do razão lida da estrutura de balanço `ZCOP` / `FAGLFLEXT`:
1. Tenta identificar o House Bank na tabela clássica `T042I`:
   ```abap
   SELECT SINGLE hbkid INTO lv_hbkid FROM t042i WHERE zbukr = ls_glt0-bukrs AND ukont = ls_glt0-racct.
   ```
2. Se não encontrar (`sy-subrc <> 0`), recorre à classe de Bank Account Management (BAM) do S/4HANA:
   ```abap
   DATA(lt_t012k) = cl_fclm_house_bank_account=>get_t012k_t(
                      it_bukrs = VALUE #( ( ls_glt0-bukrs ) )
                      it_hkont = VALUE #( ( ls_glt0-racct ) ) ).
   READ TABLE lt_t012k ASSIGNING FIELD-SYMBOL(<fs_t012k>) INDEX 1.
   IF sy-subrc = 0.
     lv_hbkid = <fs_t012k>-hbkid.
   ENDIF.
   ```
3. Se `lv_hbkid` estiver preenchido, acumula o saldo em `lt_gl_balance`. Se `lv_hbkid` estiver em branco, **a conta é sumariamente descartada**.

### 2.2 Causa Raiz na Base de Dados (PRD)
A conta de movimento `12040011` está cadastrada na `T042I` (determinação de vias de pagamento FBZP), obtendo `HBKID = 'ICB01'` sem passar pelo BAM.

A conta firme `12040010` não está na `T042I` e depende exclusivamente do BAM (`CL_FCLM_HOUSE_BANK_ACCOUNT`).  
Ao inspecionar as tabelas `FCLM_BAM_ACLINK2` e `FCLM_BAM_AMD`:
- **Conta `1000000093` (`ICB01` / CNY):**
  - Revisão `0000`: `VALID_TO = 30.05.2022`, **`STATUS = 15` (`OBSOLETED`)**
- **Conta `1000000113` (`CBK01` / EUR):**
  - Revisão `0000`: `VALID_TO = 31.12.9999`, **`STATUS = 02` (`ACTIVE`)**

O método do BAM filtra estritamente por contas não obsoletas na revisão `0000`. Como a conta CNY estava com status obsoleto desde maio de 2022, o BAM retornava vazio, fazendo com que a conta firme `12040010` fosse descartada na leitura de saldos.

---

## 3. Resolução e Validação
- **Ação:** Reativação da conta bancária / banco empresa no Bank Account Management (BAM), retirando-a do status obsoleto e ativando a revisão principal.
- **Validação:** Confirmado pelo utilizador que após a remoção do status obsoleto do banco empresa, o relatório COPE passou a calcular e consolidar perfeitamente os saldos da conta firme com a transitória.
