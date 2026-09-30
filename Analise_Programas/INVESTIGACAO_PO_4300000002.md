# Investigação: Por que PO 4300000002 não aparece em AUTO_INST_ASSIGN?

## 📋 Resumo Executivo

A PO **4300000002** não aparece na lista `AUTO_INST_ASSIGN` do programa `ZFI_PURCH_DOC_EX_RATE` porque os campos **EKKO-ZZCOORI** e **EKKO-ZZEXPVZ estão VAZIOS**.

Esta condição causa a falha no **INNER JOIN com ZFI_PAY_DATE_T**, eliminando automaticamente a PO do resultado do SELECT ABAP.

---

## 🔍 Investigação Técnica

### Ambiente
- **Sistema:** SAP S4Q (QAD)
- **Mandante:** 100
- **Empresa:** 2010
- **Programa:** ZFI_PURCH_DOC_EX_RATE
- **Método:** AUTO_INST_ASSIGN

### Código ABAP Analisado
**Ficheiro:** `ZFI_PURCH_DOC_EX_RATE_LCL.abap` (linhas 304-325)

```abap
SELECT A~BUKRS,
       A~WAERS,
       A~EBELN,
       B~EINDT,
       A~LIFNR,
       E~NDAYS,
       F~ZTAG1,
       G~NETWR,
       CASE WHEN G~ZZDAT03 IS NOT NULL THEN G~ZZDAT03 ELSE G~ZZDAT02 END AS ZZDAT02
FROM EKKO AS A
INNER JOIN EKET AS B ON B~EBELN EQ A~EBELN
INNER JOIN LFB1 AS D ON D~LIFNR EQ A~LIFNR AND D~BUKRS EQ A~BUKRS
INNER JOIN ZFI_PAY_DATE_T AS E ON E~ZZCOORI EQ A~ZZCOORI AND E~ZZEXPVZ EQ A~ZZEXPVZ
INNER JOIN T052 AS F ON F~ZTERM EQ D~ZTERM
INNER JOIN EKPO AS G ON G~EBELN = A~EBELN AND G~EBELP = B~EBELP
INTO TABLE @DATA(LT_AUTO)
WHERE A~BUKRS EQ @IV_BUKRS
AND A~WAERS EQ @IV_WAERS
AND A~EBELN IN @LR_EBELN
AND A~BEDAT IN @LR_BEDAT
AND A~AEDAT IN @LR_AEDAT
ORDER BY A~EBELN.
```

### Causa Raiz

**Ponto de Falha:** Linha 316
```abap
INNER JOIN ZFI_PAY_DATE_T AS E ON E~ZZCOORI EQ A~ZZCOORI AND E~ZZEXPVZ EQ A~ZZEXPVZ
```

**Valores em EKKO para PO 4300000002:**
- `ZZCOORI = ''` (VAZIO)
- `ZZEXPVZ = ''` (VAZIO)

**Resultado:** O JOIN procura correspondência em `ZFI_PAY_DATE_T` com ZZCOORI='' E ZZEXPVZ=''
- Nenhuma linha encontrada → **INNER JOIN retorna 0 registos**
- A PO é excluída do resultado final do SELECT

---

## 📊 Análise Comparativa

| PO | ZZCOORI | ZZEXPVZ | EKPO | PSTYP | ZFI_PAY_DATE_T | Resultado |
|----|---------|---------|------|-------|---|----------|
| **4300000002** | VAZIO | VAZIO | 1 pos | 5 (Serviço) | 0 linhas | ❌ NÃO aparece |
| **4000055997** | CN | 1 | 5 pos | 0 (Material) | 1 linha | ✅ Aparece |
| **4300000003** | CN | 1 | 1 pos | 5 (Serviço) | 1 linha | ✅ DEVERIA aparecer |

---

## 🎯 Achados Importantes

### 1. PSTYP=5 (Serviço) NÃO é a causa
- A PO 4300000003 também tem PSTYP=5
- Mas deveria aparecer porque ZZCOORI e ZZEXPVZ estão preenchidos
- **Conclusão:** O tipo de item (material vs serviço) não é o fator eliminatório

### 2. ZZCOORI e ZZEXPVZ são obrigatórios
- Campos de origem do Cockpit financeiro
- Utilizados para lookup em ZFI_PAY_DATE_T
- Se vazios → o INNER JOIN falha

### 3. Fluxo lógico de exclusão
1. EKKO: PO 4300000002 existe ✓
2. EKET: Tem 1 posição ✓
3. LFB1: Fornecedor encontrado ✓
4. **ZFI_PAY_DATE_T: 0 linhas ✗** ← AQUI A PO É ELIMINADA
5. (T052 e EKPO nunca são alcançados)

---

## 📂 Scripts de Validação

**Scripts executados:**
1. `analyze_po_4300000003.py` — Validação da PO 4300000003
2. `compare_two_pos.py` — Comparação lado a lado
3. `find_root_cause_automatic.py` — Investigação automática dos JOINs
4. `continue_joins_test.py` — Testes sequenciais

**Comando para executar:**
```bash
cd C:\workspace\SapScript
.\.venv\Scripts\python.exe Analise_Programas\find_root_cause_automatic.py
```

---

## ✅ Conclusão

A PO 4300000002 **não deveria aparecer** em AUTO_INST_ASSIGN porque:

> **Os campos EKKO-ZZCOORI e EKKO-ZZEXPVZ estão vazios, causando a falha no INNER JOIN com ZFI_PAY_DATE_T.**

Esta é uma exclusão **por design** do programa, não um bug.

---

## 📌 Recomendações

1. **Verificação de Dados:** Validar por que a PO 4300000002 foi criada sem ZZCOORI/ZZEXPVZ
2. **Documentação:** Adicionar comentário no código ABAP explicando esta dependência
3. **Testes:** Incluir casos de teste para POs com e sem ZZCOORI/ZZEXPVZ preenchidos

---

**Data da Investigação:** 2026-09-30  
**Ambiente Testado:** SAP S4Q, Mandante 100  
**Modo:** SOMENTE LEITURA (sem modificações a dados)
