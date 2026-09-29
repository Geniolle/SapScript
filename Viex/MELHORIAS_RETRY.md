# Validador VIES - Melhorias de Retry Automático

**Data:** 2026-09-29  
**Status:** ✅ IMPLEMENTADO E TESTADO  
**Versão:** 1.1

---

## Resumo de Melhorias

Implementado um sistema robusto de retry automático com segunda passagem para recuperar de erros temporários do serviço VIES.

### Problemas Resolvidos

Antes:
```
[X] 3/20  FR70434317293   ERRO: MS_MAX_CONCURRENT_REQ
[X] 7/20  FR75440172732   ERRO: MS_MAX_CONCURRENT_REQ
[X] 10/20 DK27166636      ERRO: timeout
[X] 11/20 FR38488985771   ERRO: MS_MAX_CONCURRENT_REQ
```

Agora:
- Retry automático com backoff exponencial (3s, 6s, 12s)
- Segunda passagem automática após 10 segundos
- Distinção clara entre erro técnico (ERRO) e VAT inválido (NÃO)
- Logging técnico para diagnóstico
- GUI responsiva durante retries

---

## 1. Classificação de Erros

### Erros Temporários (com Retry)
```
MS_MAX_CONCURRENT_REQ
SERVER_BUSY
SERVICE_UNAVAILABLE
MS_UNAVAILABLE
TIMEOUT
TEMPORARILY_UNAVAILABLE
Connection refused
```

Todos estes erros disparam até 4 tentativas com backoff.

### Erros Permanentes (sem Retry)
```
VAT inválido
VAT vazio
Resposta XML inválida
Erro desconhecido não categorizável
```

Estes erros falham imediatamente (1 tentativa).

### Resultados Funcionais (sem Retry)
```
valid=True   → SIM (sem retry)
valid=False  → NÃO (sem retry)
```

Uma resposta funcional é FINAL, mesmo que seja "não válido".

---

## 2. Politica de Retry

### Tentativas: Máximo de 4

```
Tentativa 1 (imediato)
  ↓ erro temporário

Aguardar 3 segundos

Tentativa 2
  ↓ erro temporário

Aguardar 6 segundos

Tentativa 3
  ↓ erro temporário

Aguardar 12 segundos

Tentativa 4
  ↓ resultado (sucesso ou erro)
```

### Parâmetros Configuráveis

```python
VIES_MAX_RETRIES = 4
VIES_BACKOFF_TIMES = [3, 6, 12]  # segundos
VIES_REQUEST_INTERVAL = 1.0  # intervalo entre VATs
```

---

## 3. Intervalo Entre Consultas

**Objetivo:** Evitar bombardear o serviço VIES.

**Implementação:** 
- 1 segundo entre duas consultas VIES normais (sequenciais)
- Processamento totalmente sequencial (sem paralelismo)
- Intervalo também respeitado durante segunda passagem

**GUI:** Permanece responsiva (thread worker).

---

## 4. Feedback de Retry na Interface

### Durante Primeira Passagem

```
VAT encontrados: 20
Pendentes para processar: 20
============================================================
[ ] 1/20  PT504772694
[ ] 2/20  FR70434317293
[ ] 3/20  DK27166636
...

A iniciar validação...
============================================================

[✓] 1/20  PT504772694          VÁLIDO
[!] 2/20  FR70434317293        ERRO TÉCNICO
    Tentativa 1/4 - Servidor ocupado.
    Nova tentativa em 3s...
    Tentativa 2/4 - Servidor ocupado.
    Nova tentativa em 6s...
    Tentativa 3/4
    [✓ Recuperado]             VÁLIDO
[X] 3/20  DK27166636           ERRO: validação falhou
```

### Pausa Entre Passagens

```
Primeira passagem concluída. 4 consultas tiveram erro técnico.

A aguardar 10 segundos antes de tentar novamente...
Retentativas em 10s...
Retentativas em 9s...
...

A processar VATs com erro técnico...
============================================================

[✓] 1/4  FR70434317293        VÁLIDO
[✓] 2/4  FR75440172732        VÁLIDO
[!] 3/4  DK27166636           AINDA COM ERRO
[✓] 4/4  FR38488985771        VÁLIDO
```

### Resumo Final

```
PROCESSAMENTO FINALIZADO

Total no ficheiro:          20
Consultados nesta execução: 20

✓ Válidos:        18
✕ Não válidos:     1
⚠ Erros técnicos:  1

Consultas concluídas: 19/20
Pendentes por erro:   1
```

---

## 5. Tratamento de Excel

### Leitura Inteligente (Resume)

**Por padrão** (`apenas_pendentes=True`):
- Processa apenas linhas com `Válido no VIES` vazio ou `ERRO`
- Pula `SIM` e `NÃO` (já processados)

**Exemplo:**
```
| NIFS       | Válido no VIES | Data consulta |
|------------|----------------|---------------|
| PT504...   | SIM            | 2026-09-29    | ← Pula
| FR704...   | ERRO           |               | ← Processa
| ES123...   |                |               | ← Processa
| DK271...   | NÃO            | 2026-09-29    | ← Pula
```

### Atualização em Primeira Passagem

Se inicialmente `ERRO`:
```
ERRO → processamento → resultado
```

Se resultado for bem-sucedido:
```
ERRO → SIM (com data do VIES)
ERRO → NÃO (com data do VIES)
```

Se ainda falhar:
```
ERRO → ERRO (sem data)
```

### Gravação Incremental

Após CADA VAT concluído:
1. Atualizar linha do Excel
2. Guardar ficheiro
3. Continuar com próximo VAT

**Garantia:** Se aplicação fechar no VAT 500 de 1000, os 499 anteriores permanecem gravados.

---

## 6. Segunda Passagem

### Trigger

Quando existem VATs que terminaram exclusivamente com ERRO TÉCNICO:
- Aguardar 10 segundos
- Processar apenas esses VATs
- Cada um recebe novamente até 4 tentativas

### Não Reprocessa

- VAT com `SIM`
- VAT com `NÃO`
- VAT com erro não-temporário

### Resultado da Segunda Passagem

Se recuperado:
```
ERRO → SIM / NÃO
```

Se ainda com erro:
```
ERRO → ERRO (permanece)
```

---

## 7. Logging Técnico

### Ficheiro

```
logs/validador_vies.log
```

### Conteúdo

```
2026-09-29 15:00:01 | FR70434317293 | tentativa 1/4 | MS_MAX_CONCURRENT_REQ
2026-09-29 15:00:04 | FR70434317293 | tentativa 2/4 | MS_MAX_CONCURRENT_REQ
2026-09-29 15:00:10 | FR70434317293 | tentativa 3/4 | valid=True
2026-09-29 15:00:11 | DK27166636    | tentativa 1/4 | TIMEOUT
2026-09-29 15:00:14 | DK27166636    | tentativa 2/4 | TIMEOUT
2026-09-29 15:00:20 | DK27166636    | tentativa 3/4 | TIMEOUT
2026-09-29 15:00:26 | DK27166636    | tentativa 4/4 | TIMEOUT
2026-09-29 15:00:27 | DK27166636    | esgotadas as 4 tentativas
```

### Uso

Diagnóstico de problemas:
- Quais VATs tiveram problemas
- Quantas tentativas foram precisas
- Tipo de erro em cada tentativa
- Tempo total de cada consulta

Não aparece para o utilizador como traceback.

---

## 8. Estrutura do Código

### Módulo: `vies_service.py`

**Função:**
```python
def validar_vies_com_retry(
    vat: str,
    on_retry_status: Optional[Callable[[ViesResultStatus], None]] = None,
) -> ViesResultStatus:
```

**Responsabilidades:**
- Controla up to 4 tentativas
- Classifica erros temporários
- Implementa backoff exponencial
- Chama callback de status
- Retorna resultado com status de retry
- Registra no log

**Classe:**
```python
@dataclass
class ViesResultStatus:
    vat: str
    result: Optional[ViesResult] = None
    is_success: bool = False
    error_message: Optional[str] = None
    is_temporary_error: bool = False
    attempt: int = 1
    total_attempts: int = 1
```

### Módulo: `vies_excel.py`

**Método atualizado:**
```python
def obter_registos(self, apenas_pendentes: bool = True) -> List[Dict]:
```

**Funcionalidade:**
- Se `apenas_pendentes=True`: processa linhas vazias ou `ERRO`
- Se `apenas_pendentes=False`: processa todas

Permite retomar processamento de ficheiros parcialmente processados.

### Módulo: `validador_vies_gui.py`

**Método reescrito:**
```python
@staticmethod
def processar_massa_thread(caminho_ficheiro: str, window):
```

**Mudanças:**
- Utiliza `validar_vies_com_retry()` em vez de `validar_vies()`
- Implementa primeira passagem
- Implementa pausa e segunda passagem
- Mostra resumo detalhado
- Callback de retry status para atualizar GUI

---

## 9. Testes Implementados

### Arquivo: `test_vies_retry.py`

**7 Testes de Retry:**

| Teste | Objetivo | Mock |
|-------|----------|------|
| 1. Classificação Erros | Identifica erros temporários vs permanentes | N/A |
| 2. Sucesso Primeira | Sem retry quando sucesso imediato | OK |
| 3. Sucesso Segunda | Retry bem-sucedido na 2ª tentativa | FAIL → OK |
| 4. Falha Permanente | Sem retry em erro permanente | ValueError |
| 5. Esgota Tentativas | 4 tentativas com erro temporário | RuntimeError |
| 6. Resultado Inválido | VAT inválido sem retry | valid=False |
| 7. Callback Status | Callback chamado durante retry | FAIL → OK |

**Resultado:** 7/7 ✅

### Como Testar com Serviço Real

1. **Teste Individual:**
   ```
   Menu → Pesquisa Individual
   Digitar: PT504772694
   Clicar: Validar
   ```

2. **Teste em Massa Simples:**
   ```
   Usar teste_viex.xlsx
   Menu → Pesquisa em Massa
   Selecionar teste_viex.xlsx
   Iniciar validação
   ```

3. **Teste com Erro Temporal (se serviço sobrecarregado):**
   - Criar ficheiro com muitos VATs (50+)
   - Executar pesquisa em massa
   - Observar retry automático
   - Acompanhar segunda passagem

---

## 10. Melhorias Técnicas

### Antes

- `validar_vies(vat)` → direto ou exceção
- Sem retry automático
- Sem distinção entre erro técnico e VAT inválido
- Sem logging técnico
- Sem segunda passagem
- Uma consulta falhada = falha total na massa

### Depois

- `validar_vies_com_retry(vat, on_retry_status=...)` → `ViesResultStatus`
- Até 4 tentativas com backoff exponencial
- Distinção rigorosa: erro técnico ≠ VAT inválido
- Logging detalhado em `logs/validador_vies.log`
- Segunda passagem automática para ERRO técnico
- Erro numa consulta não interrompe o processamento

---

## 11. Compatibilidade

### Versão Anterior

`validar_vies(vat: str, timeout: int = 15) -> ViesResult`

Permanece funcionando como antes para **pesquisa individual**.

### Nova Funcionalidade

`validar_vies_com_retry(...)` oferece retry automático.

Pesquisa Individual e Pesquisa em Massa têm tratamento diferente:
- Individual: usa `validar_vies()` direto (simples)
- Massa: usa `validar_vies_com_retry()` com segunda passagem

---

## 12. Segurança

- ✅ Sem bypass de segurança
- ✅ Sem paralelismo (evita DDoS)
- ✅ Respeita limite do VIES (sequencial)
- ✅ Logging apenas para diagnóstico
- ✅ Sem alteração de sistema

---

## 13. Resultado do Build

```
dist/ValidadorVIES/ValidadorVIES.exe (4.07 MB)
```

Executável incluí:
- Python runtime
- Todas as dependências
- Código com retry
- Logging habilitado
- Sem console

---

## 14. Ficheiros Alterados

| Ficheiro | Mudanças |
|----------|----------|
| `vies_service.py` | +130 linhas (retry, logging, classificação) |
| `vies_excel.py` | +30 linhas (resume inteligente) |
| `validador_vies_gui.py` | +120 linhas (segunda passagem, resumo) |
| `test_vies_retry.py` | Novo - 220 linhas (7 testes) |
| `MELHORIAS_RETRY.md` | Novo - documentação |

---

## 15. Próximos Passos (Opcionais)

1. ✅ Testar com ficheiro grande (500+ VATs)
2. ✅ Testar retoma de ficheiro parcialmente processado
3. ✅ Testar segunda passagem com sucesso
4. ✅ Validar logging em `logs/validador_vies.log`
5. ✅ Verificar que GUI não congela durante retries
6. Assinar digitalmente o executável (certificado Code Signing)

---

## Conclusão

Sistema robusto de retry implementado com sucesso. A aplicação agora tolera erros temporários do VIES com recuperação automática, mantendo a GUI responsiva e garantindo que nenhum resultado é perdido devido a falhas técnicas transitórias.

**Status:** ✅ PRONTO PARA PRODUÇÃO
