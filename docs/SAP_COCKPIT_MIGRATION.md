# SAP Cockpit — Migração para Core + CLI + Web

Estado por fase. Ver também [SAP_COCKPIT_REFACTOR_AUDIT.md](SAP_COCKPIT_REFACTOR_AUDIT.md) e [SAP_COCKPIT_TEST_REPORT.md](SAP_COCKPIT_TEST_REPORT.md).

## Fase 1 — Inventário e testes de referência (concluída)

- Commit `58004e6`: auditoria, remoção de credenciais em claro de `sap_script_web_cockpit_v2/README.md` e `docs/MIGRACAO_ORACLE_CLOUD.md`.
- Baseline: 65 testes OK.
- Pendente do utilizador: rodar a senha antiga no servidor Oracle (continua no histórico Git).

## Fase 2 — Segurança e contratos (concluída, não commitada)

### Alterações

| Ficheiro | Alteração |
|---|---|
| `sap_script_web_cockpit_v2/sap_cockpit_web_ready.py` | Validação de caminhos, remoção de retries por `TypeError`, contratos de entrada/saída em `run_sap_cockpit` |
| `sap_script_web_cockpit_v2/cockpit_contracts.py` | Novo: `CockpitRunRequest`, `CockpitRunResult`, `CockpitContractError` |
| `sap_script_web_cockpit_v2/tests/test_sap_cockpit_core.py` | Novo: 17 testes específicos das correções |

### Riscos corrigidos

1. **Execução duplicada por `TypeError` (alta).** `_executar_um_script` chamava `exec_fn(ambiente, **kwargs)` e, em `TypeError`, voltava a chamar `exec_fn(ambiente)`. Um `TypeError` interno do script, depois de uma escrita SAP, fazia o script correr de novo. Agora os argumentos são validados com `inspect.signature(...).bind(...)` antes da chamada, e não há nova tentativa.
2. **Retry por `TypeError` em `listar_requests` (média).** `perguntar_opcao_request` repetia a chamada com menos argumentos. Removido: a assinatura do repositório aceita todos os argumentos.
3. **Prefixo de pasta (média).** `_resolve_processo_path` usava `startswith`, que aceitava `Processos_evil` como dentro de `Processos`. Agora usa `os.path.commonpath` com `normcase`.
4. **Travessia no nome do subprocesso (alta).** `subprocesso` e `script` do payload iam direto para `os.path.join`, tanto na execução como na leitura de `WEB_CONFIG`. Agora `_resolve_subprocesso_path` aceita apenas um nome de ficheiro `.py` (sem separadores, sem unidade, sem absolutos) dentro da pasta do processo.
5. **Credenciais em claro (alta).** Corrigido na Fase 1 (README e migração).

### Contratos

- `CockpitRunRequest.from_payload(payload)` valida que o payload é um objeto, que `processo` existe e normaliza `subprocesso`/`script`. Não altera as chaves extra, que continuam a ser passadas como antes.
- `CockpitRunResult(status, log).to_dict()` produz o mesmo `{"status", "log"}` que o worker já espera. `worker/sap_tasks.py` não foi alterado.

### Import sem efeitos

- `sap_cockpit_web_ready.py` só executa `main_terminal()` sob `if __name__ == "__main__"`. Teste `ImportSemEfeitosTest` confirma que importar não chama `GetObject` nem `Dispatch`.
- `SAP Cockpit.py` (desktop original) **ainda executa código de topo ao importar** (linhas ~1005–1159). Não foi alterado nesta fase: fica para a Fase 4, ao adaptar o terminal.

### Compatibilidade

- Terminal: `main_terminal` e `executar_processo` mantêm assinatura.
- Web/worker: `run_sap_cockpit(payload)` mantém entrada dict e saída dict. Módulo carregado por `SAP_COCKPIT_MODULE` sem alteração.
- Nenhuma alteração a `web_api/`, `sap_tasks.py` ou `sap_rfc/`.

### Pendências da Fase 2

- Comparação de corpos das 37 funções duplicadas: fica para a Fase 3.
- Verificação de `SAP_COCKPIT_MODULE` e de `WORKFLOW_PYTHON_EXEC` / `SAP_FI_BRIDGE_PYTHON`: não alteradas; sem verificação específica.
- Controlo de execuções simultâneas na mesma sessão SAP GUI: Fase 6.
- Caminho do ficheiro Excel (`caminho_ficheiro`) continua a ser validado apenas por existência. Um pedido web pode apontar para qualquer ficheiro existente no worker. Avaliar na Fase 6.

### Deployment

- Não efetuado. `sap_cockpit_web_ready.py` e `cockpit_contracts.py` mudam o comportamento do worker Windows. Quando autorizado, reiniciar o worker (processo Windows separado) e validar `/health` e uma execução de leitura.

## Fases 3 a 8 (pendentes)

Ver tabela de estado em [SAP_COCKPIT_REFACTOR_AUDIT.md](SAP_COCKPIT_REFACTOR_AUDIT.md#6-estado-das-fases).

## Fase 3b — Segurança do desktop e núcleo de funções puras (concluída, não commitada)

Ver detalhe na secção 7.11 de [SAP_COCKPIT_REFACTOR_AUDIT.md](SAP_COCKPIT_REFACTOR_AUDIT.md).

| Ficheiro | Alteração |
|---|---|
| `SAP Cockpit.py` | Sem `except TypeError` em `executar_processo` e `perguntar_opcao_request`; usa `validar_argumentos_exec`; importa `parse_env_line` e `is_scripting_disabled_error` |
| `sap_script_web_cockpit_v2/sap_cockpit_web_ready.py` | Usa o core para as duas funções puras e para a validação de assinatura; `_validar_argumentos_exec` passa a ser adaptador que converte para `SapCockpitError` |
| `sap_cockpit_core.py` (novo, raiz) | `parse_env_line`, `is_scripting_disabled_error`, `validar_argumentos_exec`, `ArgumentosExecInvalidos` |
| `sap_script_web_cockpit_v2/tests/test_sap_cockpit_fase3b.py` (novo) | 13 testes: equivalência, validação de assinatura, sem retry no desktop e na web |
| `sap_script_web_cockpit_v2/tests/fixtures/sap_cockpit_golden_puras.json` (novo) | Saídas originais de `HEAD` para o corpus de equivalência |

**Não aplicado (decisão pendente):** regra única de credenciais (dependências em `transportar_request.py`, `inspect_dmeex.py` e nos `.env.example`) e obrigatoriedade de `request_type` e `request_desc` (formulário web com defaults). Ver 7.11.

**Compatibilidade:** `SAP Cockpit.py` mantém os mesmos nomes e assinaturas públicas. `run_sap_cockpit` e `main_terminal` não mudaram. Nenhum deploy feito.
