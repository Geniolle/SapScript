# Investigação SAP PRD: Conflito entre Sumarização Contabilística (TTYPV) e Split de Documento 999 Linhas (LFACIFSI)

**Data:** 15/10/2026  
**Ambiente Consultado:** SAP PRD (S4P) / Mandante 100  
**Objeto Principal:** Function Group `FACI` (Include `LFACIFSI`)  
**Documento Analisado:** `BELNR = 7470000002`, `BUKRS = 2010`, `GJAHR = 2025`, `BLART = RY`, `AWTYP = VBRK` (Fatura SD `2586801118`)

---

## 1. Resumo Executivo da Descoberta

Existe uma **incompatibilidade arquitetural direta e intencional no standard SAP** entre os dois mecanismos projetados para lidar com grandes volumes de linhas em FI:

1. **A Sumarização Contabilística FI (`TTYPV` / `TTYPVX` / `COMPRESS` / `COLLECT`)**:  
   Agrupa múltiplas linhas com a mesma chave numa única posição da BSEG através de `COLLECT P_ACC`.
2. **O Split de Fatura em Múltiplos Documentos (`SPLIT_INVOICE` / BAdI `FI_BILL_ISSUE_SPLIT`)**:  
   Particiona um lançamento de mais de 999 linhas em documentos contábeis subsequentes (`$$1`, `$$2`, com tipo de documento `ZV` e contrapartida técnica em conta `SPL`).

### A Regra Fundamental no Código:
> **Se a sumarização FI estiver ativa na tabela `TTYPV` para a aplicação/fatura (`AWTYP = 'VBRK'`), o SAP ABORTA IMEDIATAMENTE qualquer tentativa de particionamento/split de documento.**

Se a sumarização falhar (por exemplo, porque as linhas possuem diferenciação em campos como `KDPOS`, `HRKFT`, `MATNR`, `WERKS`, etc.) e o número total de posições ultrapassar 999, **o documento NÃO volta a passar pelo split**, caindo irremediavelmente no erro:
$$\mathbf{F5\ 727:\ \text{Número máximo de posições em FI atingido (999)}}$$

---

## 2. Evidência do Código ABAP Standard no PRD

### A. Ponto de Aborto do Split (`LFACIFSI`, linhas 196 a 203)
Dentro da rotina principal `FORM SPLIT_INVOICE`:

```abap
*do not perform this kind of split if summarization in FI is active
    PERFORM IS_SUMMARIZATION_ACTIVE CHANGING LV_NO_SPLIT.
    IF NOT LV_NO_SPLIT IS INITIAL.
      IF LV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '103' WITH ACCHD_FI-AWTYP.
      ENDIF.
      EXIT.
    ENDIF.
```

### B. Rotina de Verificação de Sumarização (`LFACIFSI`, linhas 1053 a 1094)
A rotina consulta dinamicamente a tabela `TTYPV` (e `TTYPVX`) para determinar se a sumarização está configurada:

```abap
*&---------------------------------------------------------------------*
*&      Form  IS_SUMMARIZATION_ACTIVE
*&---------------------------------------------------------------------*
*       Check whether summarization in FI is active
*----------------------------------------------------------------------*
  FORM IS_SUMMARIZATION_ACTIVE CHANGING C_ACTIVE TYPE XFELD.

    PERFORM READ_TTYPV.
    DESCRIBE TABLE LOGDN LINES SY-TFILL.

    IF SY-TFILL <> 0.
      LOOP AT LOGDN.
        LOOP AT GT_TTYPV WHERE AWTYP = ACCHD_FI-AWTYP
                         AND ( BUKRS = LOGDN-BUKRS
                         OR    BUKRS = SPACE )
                         AND ( BLART = LOGDN-BLART
                         OR    BLART = SPACE ).
          EXIT.
        ENDLOOP.
        IF SY-SUBRC IS INITIAL.
          C_ACTIVE = 'X'.
          EXIT.
        ENDIF.
      ENDLOOP.
    ELSE.
      CLEAR: SAVE.
      LOOP AT ACCIT_FI.
        IF ACCIT_FI-BUKRS <> SAVE-BUKRS
        OR ACCIT_FI-BLART <> SAVE-BLART.
          SAVE-BUKRS = ACCIT_FI-BUKRS.
          SAVE-BLART = ACCIT_FI-BLART.
          LOOP AT GT_TTYPV WHERE AWTYP = ACCHD_FI-AWTYP
                           AND ( BUKRS = SAVE-BUKRS
                           OR    BUKRS = SPACE )
                           AND ( BLART = SAVE-BLART
                           OR    BLART = SPACE ).
            EXIT.
          ENDLOOP.
          IF SY-SUBRC IS INITIAL.
            C_ACTIVE = 'X'.
            EXIT.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDIF.

  ENDFORM.                    " is_summarization_active
```

---

## 3. Fluxo de Decisão Técnico

```
                              [ Início: Fatura SD -> FI ]
                                           │
                                           ▼
                                 FI_DOCUMENT_PROJECT
                                           │
                       ┌───────────────────┴───────────────────┐
                       ▼                                       ▼
             PERFORM SPLIT_INVOICE                  PERFORM PROJECT_TO_V_TABLES
                       │                                       │
                       ▼                                       ▼
      PERFORM IS_SUMMARIZATION_ACTIVE                PERFORM FI_DOCUMENT_COLLECT
                       │                                       │
          ┌────────────┴────────────┐                          ▼
          │ Existe TTYPV ativa      │                 PERFORM COMPRESS
          │ para VBRK/BUKRS/BLART?  │                          │
          └────────────┬────────────┘               ┌──────────┴──────────┐
                       │                            │                     │
              SIM ─────┴───── NÃO             COMPRES = 'X'         COMPRES = ' '
               │               │                    │                     │
               ▼               ▼                    ▼                     ▼
        LV_NO_SPLIT = 'X'   Prossegue        COLLECT P_ACC         APPEND P_ACC
               │            com o Split             │                     │
               ▼            (gera $$1, $$2)         ▼                     ▼
             EXIT                               (Agrupa linhas)     (Mantém separadas)
        (PULA O SPLIT!)                             │                     │
               │                                    └──────────┬──────────┘
               │                                               ▼
               └───────────────────────────────────────► Checa > 999 linhas?
                                                               │
                                                       SIM ────┴──── NÃO
                                                        │              │
                                                        ▼              ▼
                                                    ERRO F5 727    Contabiliza OK
```

---

## 4. Análise do Caso Concreto: Documento `7470000002`

- **Fatura SD:** `2586801118` | **Doc FI:** `7470000002` | **Empresa:** `2010` | **Tipo Doc:** `RY`
- **Configuração TTYPVX:** Configurada para `VBRK / RY / 2010` (limpando `MATNR` e `SGTXT`).
- **Comportamento em `SPLIT_INVOICE`:**
  - `IS_SUMMARIZATION_ACTIVE` encontra o registro `VBRK / 2010 / RY` em `GT_TTYPV`.
  - Retorna `LV_NO_SPLIT = 'X'`.
  - O split é abortado imediatamente (`EXIT`).
- **Comportamento em `FI_DOCUMENT_COLLECT`:**
  - O SAP executa `COMPRESS` e limpa `MATNR`.
  - No entanto, a projeção intermediária e as estruturas de conta contêm diferenças em outros campos que não estão na `TTYPVX` para limpeza (como campos de classificação da linha ou determinação de objeto de custo/origem, ex. `KDPOS`/`HRKFT` ou chaves de linha).
  - Como a chave técnica diferiu, o `COLLECT` não somou as linhas da receita `0071720100`, gerando duas posições distintas (`003` e `004`).
- **Se essa fatura tivesse mais de 999 posições não agrupadas:**
  - Como `SPLIT_INVOICE` já havia sido descartado no início devido à flag de sumarização ativa, o documento não se dividiria em dois e estouraria com `F5 727`.

---

## 5. Arquivos ABAP Atualizados no Repositório

Os seguintes includes do Function Group `FACI` foram extraídos diretamente do SAP PRD (100) e arquivados em `Analise_Programas/FI_BILL_ISSUE_SPLIT_999_LINHAS/abap/`:

| Arquivo | Descrição |
| :--- | :--- |
| [`abap/LFACIFSI.abap`](./abap/LFACIFSI.abap) | Rotina principal `SPLIT_INVOICE` e `IS_SUMMARIZATION_ACTIVE`. |
| [`abap/LFACIFSP.abap`](./abap/LFACIFSP.abap) | Rotina de split para movimentos de mercadoria (`MKPF`) e adiantamentos. |
| [`abap/LFACIF4W.abap`](./abap/LFACIF4W.abap) | Rotina `FI_DOCUMENT_COLLECT` com `APPEND P_ACC` vs `COLLECT P_ACC`. |
| [`abap/LFACIF4V.abap`](./abap/LFACIF4V.abap) | Rotina `COMPRESS` que lê `TTYPV` e limpa os campos para o `COLLECT`. |
| [`abap/LFACIF0P.abap`](./abap/LFACIF0P.abap) | Rotina `READ_TTYPV` que carrega as regras da tabela de sumarização. |
