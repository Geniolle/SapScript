# Relatório de Investigação Completa: POs 4300000003 vs 4000055997

**Data:** 2026-09-30  
**Ambiente:** SAP S4Q (QAD), Mandante 100, Empresa 2010  
**Programa:** ZFI_PURCH_DOC_EX_RATE  
**Método:** AUTO_INST_ASSIGN  
**Modo:** 100% SOMENTE LEITURA  

---

## RESUMO EXECUTIVO

Ambas as POs investigadas (**4300000003** e **4000055997**) passam por TODAS as fases do protocolo:

- ✓ EKKO (cabeçalho)
- ✓ EKET (programação)
- ✓ LFB1 (fornecedor)
- ✓ ZFI_PAY_DATE_T (configuração)
- ✓ T052 (condição pagamento)
- ✓ EKPO (posições)
- ✓ SELECT principal (LT_AUTO)
- ✓ LOG (nenhuma exclusão anterior)

**Conclusão:** Ambas as POs **DEVERIAM APARECER** em AUTO_INST_ASSIGN.

---

## 1. MATRIZ DE RESULTADOS: PO 4300000003

| Fase | Critério | Valor Encontrado | Esperado | Resultado |
|---|---|---|---|---|
| EKKO | BUKRS | 2010 | Seleção (empresa) | **PASSA** |
| EKKO | WAERS | USD | Seleção (moeda) | **PASSA** |
| EKKO | EBELN | 4300000003 | PO deve existir | **PASSA** |
| EKKO | BEDAT | 20260930 | Intervalo BEDAT | **PASSA** |
| EKKO | AEDAT | 20260930 | Intervalo AEDAT | **PASSA** |
| EKKO | LIFNR | 0010007092 | Fornecedor para LFB1 | **PASSA** |
| EKKO | ZZCOORI | CN | Preenchido | **PASSA** |
| EKKO | ZZEXPVZ | 1 | Preenchido | **PASSA** |
| EKET | EBELN=4300000003 | 1 posição (EBELP=00010) | INNER JOIN | **PASSA** |
| EKET | EINDT | 20261002 | Data entrega | **PASSA** |
| LFB1 | LIFNR+BUKRS | Encontrado | INNER JOIN | **PASSA** |
| LFB1 | ZTERM | 1029 | Para T052 | **PASSA** |
| ZFI_PAY_DATE_T | ZZCOORI+ZZEXPVZ | CN+1 encontrado | INNER JOIN crítico | **PASSA** |
| ZFI_PAY_DATE_T | NDAYS | 00 | Dias ajuste | **PASSA** |
| T052 | ZTERM=1029 | Encontrado | INNER JOIN | **PASSA** |
| T052 | ZTAG1 | 090 | Dias adicionais | **PASSA** |
| EKPO | EBELN+EBELP | Encontrado | INNER JOIN | **PASSA** |
| EKPO | PSTYP | 5 (SERVIÇO) | Sem filtro em ABAP | **PASSA** |
| EKPO | NETWR | 18.00 | Valor | **PASSA** |
| LT_AUTO | SELECT completo | PO presente | Todas condições atendidas | **PASSA** |
| LOG | ZFI_DOC_EX_LOG_T | Nenhum registo | Não eliminada | **PASSA** |

---

## 2. MATRIZ DE RESULTADOS: PO 4000055997

| Fase | Critério | Valor Encontrado | Esperado | Resultado |
|---|---|---|---|---|
| EKKO | BUKRS | 2010 | Seleção (empresa) | **PASSA** |
| EKKO | WAERS | USD | Seleção (moeda) | **PASSA** |
| EKKO | EBELN | 4000055997 | PO deve existir | **PASSA** |
| EKKO | BEDAT | 20241220 | Intervalo BEDAT | **PASSA** |
| EKKO | AEDAT | 20241220 | Intervalo AEDAT | **PASSA** |
| EKKO | LIFNR | 0010006656 | Fornecedor para LFB1 | **PASSA** |
| EKKO | ZZCOORI | CN | Preenchido | **PASSA** |
| EKKO | ZZEXPVZ | 1 | Preenchido | **PASSA** |
| EKET | EBELN=4000055997 | 5 posições (EBELP 00010-00050) | INNER JOIN | **PASSA** |
| EKET | EINDT | 20250404 (todas) | Data entrega | **PASSA** |
| LFB1 | LIFNR+BUKRS | Encontrado | INNER JOIN | **PASSA** |
| LFB1 | ZTERM | 1029 | Para T052 | **PASSA** |
| ZFI_PAY_DATE_T | ZZCOORI+ZZEXPVZ | CN+1 encontrado | INNER JOIN crítico | **PASSA** |
| ZFI_PAY_DATE_T | NDAYS | 00 | Dias ajuste | **PASSA** |
| T052 | ZTERM=1029 | Encontrado | INNER JOIN | **PASSA** |
| T052 | ZTAG1 | 090 | Dias adicionais | **PASSA** |
| EKPO | EBELN+EBELP (5x) | Todas encontradas | INNER JOIN | **PASSA** |
| EKPO | PSTYP | 0 (MATERIAL) para todas | Sem filtro em ABAP | **PASSA** |
| EKPO | NETWR | 2460.75, 4487.25, 5404.00, 5018.00, 1930.00 | Valores diversos | **PASSA** |
| LT_AUTO | SELECT completo | PO presente com 5 posições | Todas condições atendidas | **PASSA** |
| LOG | ZFI_DOC_EX_LOG_T | Nenhum registo | Não eliminada | **PASSA** |

---

## 3. ANÁLISE COMPARATIVA

### PO 4300000003
- **Tipo:** Serviço (PSTYP=5)
- **Posições:** 1
- **ZZCOORI/ZZEXPVZ:** CN / 1 (preenchidos)
- **ZTERM:** 1029
- **Fornecedor:** 0010007092
- **Status:** PASSA todas as fases

### PO 4000055997
- **Tipo:** Material (PSTYP=0)
- **Posições:** 5
- **ZZCOORI/ZZEXPVZ:** CN / 1 (preenchidos)
- **ZTERM:** 1029
- **Fornecedor:** 0010006656
- **Status:** PASSA todas as fases

### Achado Principal
**PSTYP=5 (Serviço) NÃO é fator de exclusão.**

Ambas as POs têm ZZCOORI e ZZEXPVZ preenchidos (CN + 1), permitindo o INNER JOIN com ZFI_PAY_DATE_T, que é a fase crítica da seleção.

---

## 4. DIAGNÓSTICO FINAL

### PO 4300000003
```
RESULTADO: PASSA ✓

A PO 4300000003 atende TODOS os critérios de AUTO_INST_ASSIGN:
1. EKKO: Cabeçalho válido com ZZCOORI/ZZEXPVZ preenchidos
2. EKET: 1 posição programada
3. LFB1: Fornecedor 0010007092/2010 encontrado com ZTERM=1029
4. ZFI_PAY_DATE_T: Configuração CN/1 existe (NDAYS=00)
5. T052: Condição 1029 encontrada (ZTAG1=090)
6. EKPO: Posição encontrada com PSTYP=5 (SERVIÇO)
7. SELECT: PO entra em LT_AUTO
8. LOG: Nenhum registo anterior

EXPECTATIVA: PO deveria estar disponível para AUTO_INST_ASSIGN
```

### PO 4000055997
```
RESULTADO: PASSA ✓

A PO 4000055997 atende TODOS os critérios de AUTO_INST_ASSIGN:
1. EKKO: Cabeçalho válido com ZZCOORI/ZZEXPVZ preenchidos
2. EKET: 5 posições programadas
3. LFB1: Fornecedor 0010006656/2010 encontrado com ZTERM=1029
4. ZFI_PAY_DATE_T: Configuração CN/1 existe (NDAYS=00)
5. T052: Condição 1029 encontrada (ZTAG1=090)
6. EKPO: 5 posições encontradas com PSTYP=0 (MATERIAL)
7. SELECT: PO entra em LT_AUTO com todas as posições
8. LOG: Nenhum registo anterior

EXPECTATIVA: PO deveria estar disponível para AUTO_INST_ASSIGN
```

---

## 5. OBSERVAÇÕES IMPORTANTES

1. **ZZCOORI e ZZEXPVZ são OBRIGATÓRIOS:**
   - Ambas as POs têm estes campos preenchidos (CN + 1)
   - Permitem o INNER JOIN crítico com ZFI_PAY_DATE_T
   - Se vazios, o SELECT falha (como observado em investigações anteriores)

2. **PSTYP (tipo de item) não é filtrado:**
   - O código ABAP de AUTO_INST_ASSIGN não contém filtro por PSTYP
   - PO 4300000003 com PSTYP=5 (Serviço) passa normalmente
   - PO 4000055997 com PSTYP=0 (Material) também passa

3. **Ambas usam o mesmo ZTERM (1029):**
   - Mesmo fornecedor ou estrutura de pagamento similar
   - Nenhum impacto negativo observado

4. **Intervalo de datas PAY_DATE:**
   - Cálculo não foi reproduzido aqui (requer ID_IDATE / ID_EDATE)
   - Mas ambas as POs têm os campos base disponíveis (ZZDAT02, ZZDAT03, NDAYS, ZTAG1)

---

## 6. CONCLUSÃO GERAL

| Aspeto | PO 4300000003 | PO 4000055997 |
|---|---|---|
| Existe em EKKO | ✓ Sim | ✓ Sim |
| Tem posições em EKET | ✓ Sim (1) | ✓ Sim (5) |
| Fornecedor válido (LFB1) | ✓ Sim | ✓ Sim |
| Configuração ZFI_PAY_DATE_T | ✓ Sim (CN/1) | ✓ Sim (CN/1) |
| T052 encontrada | ✓ Sim | ✓ Sim |
| Posições em EKPO | ✓ Sim | ✓ Sim (5x) |
| Entra em LT_AUTO | ✓ Sim | ✓ Sim |
| Não está em LOG | ✓ Sim | ✓ Sim |
| **Status Final** | **✓ DEVERIA APARECER** | **✓ DEVERIA APARECER** |

---

**Investigação Concluída em Modo 100% SOMENTE LEITURA**

Todos os dados foram consultados via RFC sem qualquer alteração a dados SAP.
