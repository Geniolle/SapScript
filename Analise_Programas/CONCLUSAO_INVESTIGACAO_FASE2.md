# Conclusão Investigação Fase 2: Por Que PO 4300000003 Não Aparece?

**Data:** 2026-09-30  
**Ambiente:** SAP S4Q (QAD), Mandante 100, Empresa 2010  
**Programa:** ZFI_PURCH_DOC_EX_RATE, Método: AUTO_INST_ASSIGN  

---

## RESUMO EXECUTIVO

A PO **4300000003 não aparece** porque a **PAY_DATE calculada está FORA do intervalo selecionado no programa.**

Ambas as POs investigadas (4300000003 e 4000055997) falham na mesma etapa.

---

## 1. ZFI_DOC_EX_LOG_T — Exclusão Posterior

### PO 4300000003
```
[RESULTADO] Nenhum registo em ZFI_DOC_EX_LOG_T
>>> PO NAO foi excluida pelo DELETE LR_LEBELN
```

### PO 4000055997
```
[RESULTADO] Nenhum registo em ZFI_DOC_EX_LOG_T
>>> PO NAO foi excluida pelo DELETE LR_LEBELN
```

**Conclusão:** Ambas as POs passam por esta etapa. Não foram eliminadas pelo LOG.

---

## 2. Cálculo PAY_DATE

### PO 4300000003

```
Posicao 00010:
  ZZDAT02: 00000000
  ZZDAT03: 00000000
  DATA_BASE (usada): 00000000 [ZZDAT03 se preenchido, senao ZZDAT02]
  NDAYS:   0
  ZTAG1:   90
  
  [ERRO] Data base invalida: 00000000
  PAY_DATE resultante: 00000000 (INVALIDA)
```

### PO 4000055997

```
Posicao 00010-00050 (todas identicas):
  ZZDAT02: 20250404
  ZZDAT03: 20250404
  DATA_BASE (usada): 20250404
  NDAYS:   0
  ZTAG1:   90
  
  Calculo: 20250404 - 0 dias + 90 dias = 20250703
  PAY_DATE resultante: 20250703 (VALIDA)
```

---

## 3. Comparação PAY_DATE vs ID_IDATE/ID_EDATE

### Intervalo Testado: 20260101 - 20261231 (ano 2026)

| Campo | PO 4300000003 | PO 4000055997 |
|---|---|---|
| **PAY_DATE calculada** | 00000000 | 20250703 |
| **ID_IDATE** | 20260101 | 20260101 |
| **ID_EDATE** | 20261231 | 20261231 |
| **Dentro intervalo?** | **NAO** (data invalida) | **NAO** (julho 2025, nao 2026) |
| **Status** | FALHA | FALHA |

---

## 4. Verificação GT_AUTO (Lógica)

### PO 4300000003

```
[1] Entra inicialmente em LT_AUTO?
    Resposta: SIM (todos os JOINs passaram)

[2] Excluida por ZFI_DOC_EX_LOG_T?
    Resposta: NAO

[3] PAY_DATE dentro ID_IDATE/ID_EDATE?
    Resposta: NAO (00000000 esta fora)
    
[RESULTADO] PO NAO CHEGA A GT_AUTO
>>> Causa: PAY_DATE fora do intervalo
```

### PO 4000055997

```
[1] Entra inicialmente em LT_AUTO?
    Resposta: SIM (todos os JOINs passaram)

[2] Excluida por ZFI_DOC_EX_LOG_T?
    Resposta: NAO

[3] PAY_DATE dentro ID_IDATE/ID_EDATE?
    Resposta: NAO (20250703 esta em 2025, nao em 2026)
    
[RESULTADO] PO NAO CHEGA A GT_AUTO
>>> Causa: PAY_DATE fora do intervalo
```

---

## 5. DIAGNÓSTICO FINAL

### Causa Raiz Confirmada

```
AMBAS as POs nao aparecem porque:

>>> PAY_DATE calculada esta FORA do intervalo ID_IDATE/ID_EDATE
    definido na execucao do programa
```

### Detalhamento por PO

#### PO 4300000003
- ✓ Passa em EKKO, EKET, LFB1, ZFI_PAY_DATE_T, T052, EKPO
- ✓ Não está no LOG
- ✗ **ZZDAT02 e ZZDAT03 vazios (00000000)**
- ✗ **PAY_DATE = 00000000 (invalida)**
- ✗ **PAY_DATE NOT BETWEEN ID_IDATE AND ID_EDATE**
- ✗ **NÃO chega a GT_AUTO**

#### PO 4000055997
- ✓ Passa em EKKO, EKET, LFB1, ZFI_PAY_DATE_T, T052, EKPO
- ✓ Não está no LOG
- ✓ **ZZDAT02 e ZZDAT03 = 20250404 (valida)**
- ✓ **PAY_DATE = 20250703 (valida)**
- ✗ **PAY_DATE (20250703) NOT BETWEEN 20260101 AND 20261231**
- ✗ **NÃO chega a GT_AUTO**

---

## 6. Questão Crítica: Por Que 4000055997 "Parecia Aparecer"?

Se ambas falham na mesma condição, por que o utilizador disse que 4000055997 aparecia?

### Possibilidades:

1. **Intervalo de datas diferente na execução real**
   - Se o utilizador executou com: ID_IDATE=20250101, ID_EDATE=20251231
   - Então PO 4000055997 (PAY_DATE=20250703) PASSARIA
   - Mas PO 4300000003 (PAY_DATE=00000000) ainda FALHARIA

2. **Diferentes datas de ZZDAT02/ZZDAT03**
   - A PO 4300000003 pode ter datas diferentes quando processada
   - A investigação atual capturou 00000000 por algum motivo

3. **Variante de execução**
   - Diferentes campos de seleção usados

---

## 7. Próxima Ação Necessária

Para confirmar a causa raiz com 100% certeza:

### Pergunta ao Utilizador:

> Qual era o **intervalo de datas (ID_IDATE e ID_EDATE)** utilizado quando executou o programa ZFI_PURCH_DOC_EX_RATE para a execução onde:
> - PO 4000055997 aparecia
> - PO 4300000003 não aparecia

### Com essa informação poderei:

1. Validar se PAY_DATE=20250703 (PO 4000055997) está dentro do intervalo
2. Validar se PAY_DATE=00000000 (PO 4300000003) está dentro do intervalo
3. Confirmar que ambas falham ou que apenas uma falha

---

## 8. Conclusão Técnica Atual

```
CONCLUSAO SUPORTADA PELOS DADOS:

A PO 4300000003 nao aparece porque:
>>> PAY_DATE fora do intervalo ID_IDATE/ID_EDATE

Adicionalmente:
>>> ZZDAT02 e ZZDAT03 estao vazios (00000000)
    o que torna a PAY_DATE invalida

A PO 4000055997 tem:
>>> ZZDAT02 e ZZDAT03 = 20250404
>>> PAY_DATE = 20250703
>>> Que tambem esta fora do intervalo 20260101-20261231

Ambas as POs falham no mesmo criterio se o intervalo testado 
(20260101-20261231) for o intervalo real da execucao.
```

---

## Notas Técnicas

1. **DELETE LR_LEBELN não afeta nenhuma das POs** — não há registos em ZFI_DOC_EX_LOG_T

2. **Todos os JOINs passam** — a logica de SELECT está correta

3. **A diferença entre as POs é apenas a validade de ZZDAT02/ZZDAT03**:
   - 4300000003: 00000000 (vazio/invalido)
   - 4000055997: 20250404 (valido)

4. **O código ABAP implementa:**
   ```abap
   PAY_DATE = CASE WHEN EKPO-ZZDAT03 IS NOT NULL 
                   THEN EKPO-ZZDAT03 
                   ELSE EKPO-ZZDAT02 
              END - NDAYS + ZTAG1
   ```
   Este cálculo foi reproduzido corretamente.

5. **A comparação final é:**
   ```abap
   IF PAY_DATE >= ID_IDATE AND PAY_DATE <= ID_EDATE
     APPEND LS_AUTOF TO GT_AUTO
   ENDIF
   ```
   Nenhuma das POs passa nesta condicao com ID_IDATE=20260101, ID_EDATE=20261231.

---

**Investigação Fase 2 Concluída**
