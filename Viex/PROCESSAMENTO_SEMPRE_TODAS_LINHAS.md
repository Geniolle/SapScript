# Validador VIES - Processamento Sempre de Todas as Linhas

**Data:** 2026-09-29  
**Hora de Build:** 17:04:22  
**Status:** ✅ RECOMPILADO  
**Versão:** 1.1.2

---

## Alteração Implementada

### Objetivo

Remover a lógica de "apenas pendentes" que verificava status anterior das linhas. Agora a aplicação **processa SEMPRE todas as linhas**, independentemente de terem sido processadas antes.

### Razão

- O serviço VIES pode estar ativo hoje e indisponível amanhã
- Os dados podem mudar (empresa encerrada, ativação de conta)
- Nenhuma suposição sobre "já foi processado"
- O timestamp no Excel é apenas referência histórica

---

## Mudanças Técnicas

### Ficheiro 1: `vies_excel.py`

#### Antes
```python
def obter_registos(self, apenas_pendentes: bool = True) -> List[Dict]:
    """
    Se apenas_pendentes=True (padrão):
        - Processa apenas linhas com "Válido no VIES" vazia ou "ERRO"
    Se apenas_pendentes=False:
        - Processa todas as linhas
    """
    # ... lógica complexa de verificação de status anterior
```

#### Depois
```python
def obter_registos(self) -> List[Dict]:
    """
    Lê todos os VATs da coluna NIFS para processamento.
    
    Processa SEMPRE todas as linhas, independentemente de terem sido
    processadas antes. O timestamp no Excel indica a última validação.
    """
    # ... simples leitura de todos os VATs
```

**Mudanças:**
- ✅ Removido parâmetro `apenas_pendentes`
- ✅ Removida lógica de verificação de status (SIM/NÃO/ERRO)
- ✅ Sempre retorna todas as linhas com VAT

### Ficheiro 2: `validador_vies_gui.py`

#### Antes
```python
# Primeira passagem - processar pendentes
registos = excel.obter_registos(apenas_pendentes=True)

if not registos:
    window["LOG_MASSA"].print("Nenhum VAT pendente encontrado.")
    ...

# Estatísticas
total_ficheiro = len(excel.obter_registos(apenas_pendentes=False))
```

#### Depois
```python
# Primeira passagem - processar todos os VATs
registos = excel.obter_registos()

if not registos:
    window["LOG_MASSA"].print("Nenhum VAT encontrado no ficheiro.")
    ...

# Estatísticas
total_ficheiro = len(registos)
```

**Mudanças:**
- ✅ Removidas chamadas com `apenas_pendentes=True/False`
- ✅ Uma única leitura de registos
- ✅ Mensagens ajustadas

---

## Fluxo de Processamento (Novo)

```
1. Abrir Excel
2. Ler TODAS as linhas (sem verificação de status)
3. Mostrar lista completa de VATs
4. PRIMEIRA PASSAGEM
   ├─ Processar cada VAT
   ├─ Até 4 tentativas com retry
   ├─ Atualizar Excel com resultado
   └─ Guardar após cada VAT
5. SEGUNDA PASSAGEM (se houver erros técnicos)
   ├─ Aguardar 10 segundos
   ├─ Reprocessar VATs com ERRO
   └─ Até 4 tentativas novamente
6. Mostrar resumo final
7. Guardar Excel
```

**Garantia:** Nenhuma linha é pulada por estar marcada como SIM/NÃO.

---

## Exemplos de Comportamento

### Cenário 1: Excel com Histórico Anterior

```
| NIFS        | Válido no VIES | Data consulta    |
|-------------|----------------|------------------|
| PT504...    | SIM            | 2026-09-28       | ← SERÁ REPROCESSADO
| FR704...    | NÃO            | 2026-09-28       | ← SERÁ REPROCESSADO
| ES123...    | ERRO           |                  | ← SERÁ REPROCESSADO
| DK271...    |                |                  | ← SERÁ PROCESSADO
```

**Resultado após nova execução:**
- Todos os 4 VATs são consultados novamente
- O timestamp é atualizado com a nova data/hora
- Status anterior é substituído pelo resultado atual

### Cenário 2: Dados Podem Mudar

```
Execução 1 (2026-09-28):
PT504772694 → VÁLIDO

Execução 2 (2026-09-29):
PT504772694 → Empresa encerrada (INVÁLIDO)
              → Excel atualizado com novo resultado
```

A aplicação detecta mudanças que ocorreram entre execuções.

---

## Impacto no Tempo de Processamento

### Execução Anterior (com verificação)
```
100 VATs no ficheiro
20 já processados (pulados)
80 processados
Tempo: ~3-4 minutos
```

### Execução Agora (sem verificação)
```
100 VATs no ficheiro
100 processados (nenhum pulado)
Tempo: ~4-5 minutos
```

**Diferença:** +1 minuto para garantir atualização completa.

---

## Excel: Timestamp como Referência Única

O Excel agora funciona como:

```
Coluna "Data consulta":
├─ Timestamp da última validação VIES
├─ Atualizado a cada execução
└─ Sem verificação "já foi processado"
```

**Uso:**
- Ver quando foi validado pela última vez
- Comparar com data de hoje
- Decidir manualmente se fazer nova execução

---

## Novo Executável

```
Caminho:    C:\workspace\SapScript\Viex\dist\ValidadorVIES\ValidadorVIES.exe
Tamanho:    4.07 MB
Build:      2026-09-29 17:04:22  ← Novo timestamp
Status:     ✅ Pronto para distribuição
```

---

## Resumo de Mudanças

| Aspecto | Antes | Depois |
|---------|-------|--------|
| Verificação de status anterior | Sim (pula SIM/NÃO) | Não (processa tudo) |
| Parâmetro `apenas_pendentes` | Sim | Removido |
| Complexidade de lógica | Alta (verificações) | Baixa (leitura simples) |
| Linhas sempre processadas | Não | Sim |
| Timestamp no Excel | Referência + controle | Referência apenas |
| Tempo de processamento | ~3-4 min | ~4-5 min |
| Capacidade de detectar mudanças | Limitada | Total |

---

## Garantias

✅ Nenhuma linha é ignorada por status anterior  
✅ Cada execução revalida 100% dos VATs  
✅ Timestamp sempre atualizado  
✅ Retry automático mantido (até 4 tentativas)  
✅ Segunda passagem mantida (para erros técnicos)  
✅ Gravação incremental mantida (após cada VAT)  

---

## Compatibilidade

- ✅ Ficheiros Excel já processados funcionam normalmente
- ✅ Dados anteriores são sobrescritos com novos resultados
- ✅ Sem quebra de funcionalidade
- ✅ Interface mantida igual

---

**Conclusão:** Aplicação agora processa TODAS as linhas sempre, com timestamp como referência. Sem suposições sobre status anterior. Pronta para distribuição.
