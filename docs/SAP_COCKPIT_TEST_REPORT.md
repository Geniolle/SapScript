# SAP Cockpit — Relatório de Testes

## Fase 1 — Baseline

| Comando | Ambiente | Resultado |
|---|---|---|
| `sap_script_web_cockpit_v2\.venv\Scripts\python.exe tests\run_all.py` | venv do cockpit | 65 testes OK, `js_smoke` OK |

## Fase 2 — Segurança e contratos

### Execução da suite

| Comando | Ambiente | Resultado |
|---|---|---|
| `.venv\Scripts\python.exe tests\run_all.py` | venv do cockpit (sem `rich`/`pywin32`) | 82 executados: 65 OK + 17 saltados (`tests.test_sap_cockpit_core`, dependências em falta). `RESULTADO: OK` |
| `worker\.venv\Scripts\python.exe -m unittest tests.test_sap_cockpit_core` | venv do worker (tem `rich` e `pywin32`) | **17 OK** |

A suite completa do Agente Salsa IT não corre no venv do worker (falta `fastapi`); isso é esperado e não foi alterado.

### Prova de que os testes apanham os bugs

Os mesmos 17 testes foram executados contra `sap_cockpit_web_ready.py` de `HEAD` (pré-Fase 2) numa cópia em scratchpad:

- 3 FAIL e 9 ERROR, incluindo:
  - `test_typeerror_dentro_do_script_nao_reexecuta` → FAIL (o script corre duas vezes).
  - `test_pasta_irma_com_mesmo_prefixo_e_rejeitada` → FAIL (`Processos_evil` aceite).
  - `test_web_config_nao_le_fora_da_pasta` → FAIL (leitura fora da pasta).
  - Os restantes ERROR são `AttributeError` por o helper `_resolve_subprocesso_path` não existir antes da correção.

### Testes por correção

| Correção | Testes |
|---|---|
| Sem retry por `TypeError` na execução | `test_typeerror_dentro_do_script_nao_reexecuta`, `test_argumento_obrigatorio_em_falta_falha_antes_de_executar`, `test_script_sem_argumentos_extra_executa_uma_vez` |
| Travessia em subprocesso | `test_subprocesso_com_travessia_e_rejeitado`, `test_subprocesso_absoluto_e_rejeitado`, `test_nome_de_subprocesso_com_travessia_nao_carrega_modulo`, `test_web_config_nao_le_fora_da_pasta`, `test_subprocesso_valido_fica_dentro_da_pasta` |
| Prefixo de pasta | `test_pasta_irma_com_mesmo_prefixo_e_rejeitada`, `test_pasta_dentro_de_processos_e_aceite` |
| Contratos | `test_pedido_sem_processo_e_rejeitado`, `test_pedido_aceita_subprocesso_ou_script`, `test_pedido_nao_dict_e_rejeitado`, `test_resultado_tem_formato_que_o_worker_espera`, `test_run_sap_cockpit_sem_processo_falha_sem_tocar_no_sap` |
| Import sem efeitos | `test_importar_motor_nao_abre_sap`, `test_execucao_direta_continua_a_ser_o_unico_ponto_de_entrada` |

### Garantias de segurança dos testes

- Nenhum teste liga ao SAP. `GetObject` e `Dispatch` são substituídos por mocks que falham se forem chamados.
- Os testes de execução usam scripts temporários que apenas incrementam um contador em `builtins`.
- Ficheiros temporários em `tempfile.mkdtemp()`, removidos no `tearDown`.

## Não executado

- Testes de integração SAP GUI, RFC e worker Windows com o SAP real.
- Validação HTTP do container `sap-script-web` (sem deployment nesta fase).
- Comparação de corpos das funções duplicadas (Fase 3).

## Fase 3b — Segurança do desktop e núcleo de funções puras

| Comando | Ambiente | Resultado |
|---|---|---|
| `.venv\Scripts\python.exe testsun_all.py` | venv do cockpit | 95 executados: 77 OK, 18 saltados (dependências do worker em falta), `RESULTADO: OK` |
| `worker\.venv\Scripts\python.exe -m unittest tests.test_sap_cockpit_core tests.test_sap_cockpit_fase3b` | venv do worker | 30 OK |

Regressão: `DesktopSemRetryTest` e `DesktopListarRequestsSemRetryTest` executados contra `SAP Cockpit.py` de `HEAD` falham (2 ERROR, 1 FAIL). Contra o código novo passam.

Equivalência: 22 casos de `parse_env_line` e 17 de `is_scripting_disabled_error`, com saídas geradas a partir de `HEAD` (fixture `tests/fixtures/sap_cockpit_golden_puras.json`). Desktop, web e core coincidem em todos.

Não executado: testes SAP/RFC/GUI reais; nenhum teste liga ao SAP. O desktop não é importado pelos testes (executa código de topo); é lido por AST.
