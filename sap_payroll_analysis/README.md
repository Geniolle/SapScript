# sap_payroll_analysis

Diagnóstico read-only de Payroll (RH) → FI.

Investiga divergências entre o que o Payroll lançou na PCP0 e o que ficou
contabilizado em FI numa conta do Razão.

**Todo o pacote é estritamente de leitura.** Nenhuma função que escreva,
altere, elimine, faça `COMMIT`, poste documentos ou execute jobs é chamada.
Ver `security.py`.

## Execução

```bash
# 1) diagnóstico (liga, verifica tabelas, lista campos, lê amostras)
python analisar_payroll_fi.py --diagnostic

# 2) análise completa (parâmetros por omissão = caso 2010 / 06-2026 / 23120000)
python analisar_payroll_fi.py

# 3) parâmetros à medida
python analisar_payroll_fi.py --company 2010 --year 2026 --month 6 --account 23120000

# equivalente
python -m sap_payroll_analysis --diagnostic
```

Usar o interpretador com `pyrfc` instalado (neste projeto: `.venv-rfc`).

## Configuração (.env do projeto)

Lê **exclusivamente** do `.env`, pela ordem `SAP_R3_*` → `SAP_DEV_*` (fallback,
mesmo host). Nunca há credenciais no código.

```
SAP_R3_USER=
SAP_R3_PASSWD=
SAP_R3_ASHOST=10.1.1.101
SAP_R3_SYSNR=00
SAP_R3_CLIENT=100
SAP_R3_LANG=PT
```

## Output

`output/payroll_fi_<empresa>_<ano>_<mes>_<conta>.json` + CSV (`;`, UTF-8-SIG)
com os itens do posting RH e do FI.

## Módulos

| ficheiro | função |
|---|---|
| `config.py` | parâmetros do caso, `Decimal`, whitelists de tabelas e funções RFC |
| `security.py` | `SecurityError`, `safe_rfc_call` (único ponto de entrada RFC) |
| `sap_connection.py` | context manager `sap_connection()`, leitura do `.env` |
| `sap_reader.py` | `RFC_READ_TABLE` genérico com paginação, parsing SAP→`Decimal`, sinal D/C |
| `ddic.py` | `describe_table` (DD03L + fallback), heurística de campos por conceito |
| `payroll_posting.py` | PPDHD→PPDIT via `DOCNUM`; itens da conta por run/empresa |
| `payroll_wagetypes.py` | PPOIX: agrega `BETRG` por rubrica (`LGART`); **fase 2** liga cada linha PPOIX à linha de posting contabilística e cruza a determinação de contas (T52EL/T52EK/T030) |
| `payroll_cluster.py` | **fase 3** RGDIR (`HRPY_RGDIR`) + timeline/pares por PERNR + catálogo automático de tabelas de resultado + comparação do run 1299 |
| `manual_request.py` | **fase 4** shortlist mínima PERNR/SEQNR para consulta manual no `PC00_M99_CWTR` (não toca no cluster) |
| `wagetype_trace.py` | **fase 4.1** rastreio contabilístico de uma rubrica: PPOIX → (TSLIN=LINUM) PPDIX → PPDIT → HKONT; agrega o TSLIN inteiro e investiga o resíduo sem o classificar como "arredondamento" sem prova |
| `posting_delta_trace.py` | **fase 4.2** localiza o ESTÁGIO técnico do delta `SUM(PPOIX por linha de transferência) − PPDIT.WRBTR`: checkpoints PPOIX→PPOPX→PPDST→PPDIT, reconciliação LINUM a LINUM, análise `TSLIN=0`, mapa de `SEQNO`, runs anteriores, breakdowns por LGART/PERNR/MOMAG, `find_first_divergence()` e `classify_delta_origin()` |
| `payment_reconciliation.py` | **fase 5.0** reconcilia RH (`/559` corrente por PERNR) × programa de pagamentos `REGU*` no R/3: DDIC de REGUH/REGUP/REGUV, descoberta e escolha do payment run (multi-evidência, nunca só pelo total), identidade do colaborador (PERNR/EMPFG/LIFNR), `PayrollPaymentExpectation`, matching em níveis, `classify_reconciliation()`. **Exige `SAP_R3_*` completo — sem fallback.** |
| `fi_analysis.py` | ACDOCA (S/4) ou BSIS/BSAS/BSEG (ECC) + BAPI de saldo por período |
| `report.py` | reconciliação, relatório de terminal, JSON/CSV |
| `diagnostics.py` / `analysis.py` / `cluster_cli.py` / `cli.py` | orquestração |

## Testes

```bash
python -m pytest sap_payroll_analysis/tests
```

Não tocam no SAP (usam `tests/fakes.py`).

## Estrutura das tabelas HR posting (ECC, confirmada por DDIC)

```
PEVST  RUNID .............................. registo/estado do ciclo
PPDHD  RUNID -> DOCNUM (BUKRS, BUDAT) ..... cabeçalho do documento de posting
PPDIX  RUNID <-> DOCNUM/DOCLIN ............ índice
PPDIT  DOCNUM/DOCLIN: HKONT, BUKRS, WRBTR, WAERS, KTOSL, NEG_POSTNG
PPOIX  RUNID, PERNR, LGART, KOMOK, BETRG .. origem por rubrica salarial
```

* `PPDIT` **não tem** campo de run — liga-se por `DOCNUM` via `PPDHD`.
* `PPDIT.WRBTR` já traz sinal (ex.: `"727258.35-"`). `NEG_POSTNG='X'` inverte.
* Convenção de sinal em todo o pacote: **Débito = +, Crédito = −**.

## Fase 2 — composição por rubrica de uma linha de posting

Cadeia de ligação (confirmada no sistema real):

```
PPOIX.TSLIN  ==  PPDIX.LINUM
PPDIX.(DOCNUM, DOCLIN)  ==  PPDIT.(DOCNUM, DOCLIN)   <- linha da conta do Razão
```

`link_wage_types_to_posting_line(conn, params, payroll_report, run_id=None)`:

1. isola a(s) linha(s) PPDIT da conta `params.conta` para `params.primary_run`;
2. em PPDIX obtém os `LINUM` (linhas de transferência) que alimentam essa(s)
   linha(s);
3. soma `PPOIX.BETRG` por `LGART` para as linhas cujo `TSLIN` está nesses
   `LINUM`; separa `/558`+`/559` das restantes rubricas;
4. calcula o resíduo `PPOIX ligado − linha PPDIT`.

`resolve_account_determination(...)` cruza, só de leitura:

* `T52EL` rubrica → conta simbólica (`SYMKO`) + `SIGN`
* `T52EK` `KOART` da conta simbólica
* `T030`  `KTOPL` + `KTOSL` (=`PPDIT.KTOSL`, ex. `HRF`) + `BWMOD` (=rubrica ou
  conta simbólica) → `KONTS`/`KONTH`, i.e. a conta do Razão

CLI: `--run <RUNID>` escolhe o run analisado na fase 2 (por omissão 1298).
Output extra: `output/wage_link_<...>_run<RUNID>.csv`.

### Resultado no caso (run 1298, empresa 1010, conta 23120000)

| rubrica | linhas | montante EUR |
|---|--:|--:|
| /559 | 326 | −724.461,41 |
| /558 | 2 | −12,97 |
| /563 | 8 | −2.587,53 |
| /561 | 3 | +627,91 |
| 0029 | 2 | −1.090,00 |
| **/558 + /559** | | **−724.474,38** |
| **outras rubricas** | | **−3.049,62** |
| **total PPOIX ligado** | 341 | **−727.524,00** |
| linha PPDIT (FI) | | −727.258,35 |
| **resíduo** | | **−265,65** |

Determinação de contas confirmada: `S003` (KOART `F`); `T030` chart `PCPT`,
`KTOSL=HRF`, `BWMOD ∈ {/558,/559,/561,/563,0029}` → `23120000`.

## Fase 3 — contexto automático, RGDIR, timeline e retroactividade

Tudo automático — nenhum PERNR é pedido. `collect_payroll_context()` percorre
`PPDHD → PPDIT → PPDIX → PPOIX → PERNR → HRPY_RGDIR → PA0001`.

CLI: `--explain-cluster` (só a fase 3), `--rt-link-diagnostic` (RT↔PPOIX por
PERNR), `--payroll-timeline` (timeline + pares por PERNR), `--pernr <n>`,
`--no-cluster`.
Output: `payroll_cluster_analysis_<...>.json`, `rgdir_<...>.csv` (RGDIR completo
com classificação), `rgdir_pairs_<...>.csv` (pares original→recalculado),
`ppoix_rgdir_view_<...>.csv` (vista única PPOIX×RGDIR por PERNR),
`hrpy_catalog_<...>.csv` (catálogo descoberto).

### Contexto (descoberto automaticamente)

MOLGA **19** (Portugal) · ABKRS **Z2** · PERMO **01** (mensal) · IN-period
**202606** · RELID PCL2 **RP**. Diretório em `HRPY_RGDIR` (transparente).

### Catálogo de tabelas de resultado — descoberta automática (DD02L/DD03L)

`discover_payroll_result_tables()` varre `HRPY_%`, `HRPADNLP_%`, `P2RX_%`,
`PYD_D_RES%` e testa acesso/dados. Resultado neste sistema: **~72 tabelas
transparentes `P2RX_*` / `HRPADNLP_P2RX_*` existem no DDIC** (`P2RX_RT`,
`P2RX_CRT`, `P2RX_BT`, `P2RX_RT_PERSON`, …) **mas estão todas VAZIAS** — o
framework "Payroll Results Tables" não está activo. Só `HRPY_RGDIR` e
`HRPY_WPBP` têm dados.

### Limitação técnica — os montantes da RT continuam `MANUAL_REQUIRED`

| via | resultado |
|---|---|
| tabelas transparentes `P2RX_RT`/`P2RX_CRT`/… | existem mas **vazias** |
| `HR_GET_PAYROLL_RESULTS` | não é RFC-enabled (*"cannot be used for 'remote' calls"*) |
| `PYXX_READ_PAYROLL_RESULT` | `DA300 – "No active nametab"` (IMPORT PCL2 não funciona em RFC stateless) |

A RT vive só no cluster `PCL2(RP)`. Sem um **wrapper Z read-only** no SAP (ou
extracção por relatório `PC00_M99_CWTR`) os resíduos não se atribuem por PERNR
ao cêntimo. Nenhuma função de escrita foi usada.

### O que a fase 3 apura (RGDIR + timeline, 100 % automático)

* **Esta folha corre com desfasamento sistemático de 1 mês**: o resultado
  definitivo de cada FOR-period é produzido no run seguinte. Dos 321 PERNR com
  componente retro no run, **313 são retro só de rotina (1 mês)** e apenas **8
  têm correcção real (≥2 meses)**. Meses de retro por PERNR: 1×311, 2×6, 3×2,
  5×2.
* O run 1298 (IN-period 202606) transferiu, por PERNR, o resultado de 06/2026
  **mais** o recálculo de rotina de 05/2026. O PPOIX (−724.474,38 em /558+/559)
  é a RT tal como transferida; a referência RH (724.046,64) é outro recorte.
* `build_timeline()` + `pair_results()` produzem, por `(PERNR, FPPER)`, o par
  **original (SEQNR, INPER=FPPER) → actual (SEQNR, SRTZA=A)** e o estado
  (`RESULT_RECALCULATED` / `RESULT_UNCHANGED` / …). 656 pares alimentam o run.
* Os resíduos **427,74** (PPOIX vs RH) e **265,65** (PPOIX vs PPDIT) só se
  fecham ao cêntimo com a RT (definitivo vs provisório do período).
* **Run 1299 = repetição do 1298** (mesmos 324 PERNR, mesmos totais por rubrica,
  mesmo valor). `REPEAT_POSTING / RERUN` — não é estorno.

## Fase 4 — shortlist mínima para `PC00_M99_CWTR`

`python analisar_payroll_fi.py --run 0000001298 --manual-rt-request`

Não consulta o cluster: usa `output/payroll_cluster_analysis_*.json` (ou, com
RFC, `analyse_cluster(..., try_rt=False)` — só tabelas transparentes). Classifica
cada PERNR do run por evidência estrutural de recálculo (sem RT):

| categoria | critério | vai para a shortlist? |
|---|---|---|
| **B** | correcção real (FPPER recalculado ≥2 meses depois) **OU** PPOIX c/ `/561`·`/563`·`0029` | **sim** |
| **E** | finalização mensal de rotina (desfasamento de 1 mês) — não confirmável sem RT | não |
| **A** | retro de processamento sem par de versões | não |
| **C** / **D** | off-cycle / void-reversal | não |

Resultado no caso: dos **324 PERNR**, só **15 (categoria B)** precisam de
consulta manual → **24 pares OLD→NEW SEQNR**:
* prioridade 1 (5 PERNR): `00000069, 00002865, 80000006, 80000019, 80001637`
* prioridade 2 (10 PERNR): `00000005, 00000009, 00000197, 00005365, 00006637,
  00006881, 00007855, 80000208, 80000377, 80001145`
* 281 PERNR de rotina ficam de fora — só analisar se estes não fecharem.

Output:
* `output/manual_rt_shortlist_run<run>.csv` — `PERNR;FPPER;OLD_SEQNR;OLD_INPER;
  OLD_SRTZA;NEW_SEQNR;NEW_INPER;NEW_SRTZA;PPOIX_558…;PPOIX_TOTAL;RETRO_MONTHS;
  PRIORITY;REASON`, ordenada por prioridade → impacto `/558+/559` → nº de FPPER
  recalculados.
* `output/manual_rt_request_run<run>.txt` — um bloco `CASO n` por par, no
  formato *"Preciso da RT do SEQNR A e do SEQNR B para o PERNR X"*, com as
  rubricas a exportar (`LGART, BETRG, ANZHL, RTE, APZNR, C1ZNR, V0ZNR, ALZNR`).

## Fase 4.1 — rastreio contabilístico de uma rubrica

```
python analisar_payroll_fi.py --run 0000001298 --trace-wagetype \
    --pernr 00000005 --lgart 0029 [--compare-lgart /559]
```
(No Git-Bash: `MSYS_NO_PATHCONV=1` ou `--lgart='//559'` para rubricas `/NNN`;
o CLI também recupera `/559` de um caminho mangled.)

Segue a cadeia **PPOIX → (`TSLIN` = `PPDIX.LINUM`) → PPDIX → (`DOCNUM/DOCLIN`) →
PPDIT → `HKONT`**, agrega **todos** os TSLIN/LINUM que alimentam a linha da
conta (não só os da rubrica) e investiga o delta `SUM(PPOIX) − PPDIT.WRBTR`.
`explain_amount_sign_path()` documenta cada sinal em separado (BETRG, ACTSIGN,
NEG_POSTNG, WRBTR) — não deduz D/C só do texto. 100 % read-only, sem cluster.

Output: `output/trace_<run>_<pernr>_<lgart>.json` / `.csv`.

### Resultado — `0029` / PERNR 00000005 (run 1298)

| passo | valor (comprovado pelos dados) |
|---|---|
| PPOIX | 1 registo: SEQNO `01199`, BETRG `-265,00`, KOMOK `S003`, MOMAG `2`, TSLIN `0000000004` |
| PPDIX | LINUM `0000000004` → `DOCNUM 0000005392 / DOCLIN 0000000326` (destino único) |
| PPDIT | `5392/326`, HKONT `0023120000`, KTOSL `HRF`, WRBTR `-727.258,35` |
| **0029 → 23120000** | **SIM** · **0029 → PPDIT 5392/326** | **SIM** |
| determinação | `T52EL 0029 → SYMKO S003 (SIGN +)` · `T52EK S003 → KOART F` · `T030 PCPT/HRF/BWMOD=0029/KOMOK=2 → 23120000` (KOMOK=MOMAG) |
| `/559` do mesmo PERNR | SEQNO `01199` `-1.382,40` → **mesma** TSLIN 4 → **mesma** PPDIT 5392/326; SEQNO `01198` `-823,70` tem TSLIN `0` = **não transferido** |
| agregação do TSLIN {4, 347} | 341 linhas, SUM `-727.524,00` (/559 `-724.461,41`, /558 `-12,97`, /563 `-2.587,53`, /561 `+627,91`, 0029 `-1.090,00`) |
| PPDIT.WRBTR | `-727.258,35` |
| **DELTA** | **`-265,65`** |

### O resíduo de 265,65

* `265,65 − 265,00 (0029/PERNR5) = 0,65` é **coincidência aritmética**: a linha
  0029 já está incluída na soma PPOIX; excluí-la é arbitrário e deixa 0,65.
* Sem linha PPOIX `= delta` nem `= 0,65`, sem valores `≤ 1,00`, sem fracções de
  cêntimo, sem `ACTSIGN≠A` nem `NEG_POSTNG='X'`, sem split para `23110000`.
* **Classificação: `UNEXPLAINED`.** O delta nasce dentro do programa de posting
  HR ao construir a linha colectora. **Hipótese não provada**: netting de deltas
  de retro já lançados — o PPOIX carrega o resultado bruto e a PPDIT só a
  diferença. Só demonstrável com o trace do `RPCIPE00` ou o documento FI.

## Fase 4.2 — origem técnica do delta de 265,65

```
python analisar_payroll_fi.py --trace-posting-delta \
    --run 0000001298 --docnum 0000005392 --doclin 0000000326
python analisar_payroll_fi.py --analyze-zero-tslin  --run 0000001298
python analisar_payroll_fi.py --trace-seqno-history --run 0000001298
```

Reconstrói a cadeia com **checkpoints persistidos** e classifica *em que estágio*
o valor muda. Não procura "valores que somem 265,65" — exige relação técnica.
Output: `posting_delta_0000001298_5392_326.json` + `posting_delta_items_*.csv`,
`zero_tslin_0000001298.csv`, `seqno_history_0000001298.csv`,
`previous_run_trace_0000001298.csv`.

### Factos comprovados (RFC_READ_TABLE, sistema 10.1.1.101, run 1298)

| checkpoint | valor | evidência |
|---|---|---|
| **PPOIX** `TSLIN ∈ {4, 347}` (alimentam 5392/326 via PPDIX) | **−727.524,00** (341 linhas) | `[PROVED]` — LINUM 4 = −726.602,68 (340 l., MOMAG 2); LINUM 347 = −921,32 (1 l., MOMAG 3, PERNR 89000020) |
| **PPOPX** | *sem campo monetário* | `[PROVED]` — campos: MANDT/PERNR/SEQNO/RUNID/POSTNUM/TSLIN/ACTSIGN; 4909 l. no run, todas ACTSIGN='P'; **0** correspondências com as 341 linhas (4 variantes de chave) |
| **PPDST** | *vazio para o run* | `[PROVED]` — tem WRBTR mas 0 linhas para DOCNUM 5392 |
| **PPDIT** `5392/326` | **−727.258,35** | `[PROVED]` — HKONT 23120000, KTOSL HRF, 1 linha; única linha da conta no documento; doc balança a 0,00 |
| **DELTA** | **−265,65** | `PROVED_BETWEEN_PPOIX_AND_PPDIT` |

* **Estágio:** entre PPOIX e PPDIT **não existe checkpoint com montante**
  (PPDIX/PPOPX sem campo de valor; PPDST vazio). O delta nasce dentro do
  programa de posting HR (`RPCIPE00`/`SAPLHRPP`) ao construir a linha do
  documento — não é um registo observável.
* **Não é rounding** (0 linhas com fracção de cêntimo), **não é sinal**
  (ACTSIGN='A', NEG_POSTNG vazio em todas), **não é re-routing** (só há uma
  linha 23120000 no doc), **não é netting de run anterior** (a empresa 1010 só
  tem os runs 1298 e 1299 em 06/2026; 1299 é gémeo — mesmo doc-line 326, mesmo
  −727.258,35; não há run 1010 anterior), **não é PPOPX** (0 overlap).
* **Contexto `[OBSERVED]`:** este −265,65 é a fração, nesta linha, de uma
  diferença de **17.803,89** em *todo* o documento entre `SUM(PPOIX por linha de
  transferência)` e `PPDIT.WRBTR`. Essa diferença é espelhada pelos PPOIX com
  `TSLIN ∈ {0, 17}` (**−17.803,89**): working set do split (pares D/C do mesmo
  LGART que se anulam + resíduo por rubrica) que o programa redistribui pelas
  linhas lançadas e **não emite** como linha própria. `TSLIN=0`: 3157 l.,
  −25.774,89; `TSLIN=17` (/551, /552): 733 l., +7.971,00.
* **Regra exacta que produz precisamente 265,65: `UNEXPLAINED`** a partir das
  tabelas transparentes. Exigiria o *spool* do `RPCIPE00` ou debug de
  `PC00_M99_CIPE` (fora do âmbito read-only / sem PYXX).

## Fase 5.0 — reconciliação RH × REGU* (programa de pagamentos)

```
python analisar_payroll_fi.py --reconcile-payroll-regu \
    --run 0000001298 --company 1010 --period 202606
# opcional: --payment-run-date 20260628 --payment-run-id PAY001
```

Compara, **por colaborador**, o `/559` corrente do payroll (valor real de
transferência bancária — p.ex. PERNR 5 = `1.382,40`, **não** `/550` 1.647,40)
com o valor levado às tabelas `REGU*`. Só R/3:

> **Exige `SAP_R3_USER/PASSWD/ASHOST/SYSNR/CLIENT` completos no `.env`.**
> Se faltar algum, aborta com
> `ERROR: SAP_R3_* connection parameters required for payroll/payment reconciliation.`
> — **nunca** cai no fallback `SAP_DEV_*` (que aponta para outro sistema sem estes dados).

Passos: DDIC de `REGUH/REGUP/REGUV` → descoberta de payment runs na janela
`AAAAMM01 … (mês+1)15` (filtro `LAUFD` empurrado para o SELECT — filtrar só por
empresa em `REGUH` rebenta com `TSV_TNEW_PAGE_ALLOC_FAILED`) → escolha por
**múltiplas evidências** (empresa, período, nº de beneficiários ~ nº de
colaboradores, moeda, run real vs proposta; o total tem **peso mínimo**) com
`HIGH/MEDIUM/LOW_CONFIDENCE` → **conjunto de runs** (`select_payment_run_set`):
o pagamento pode estar dividido por vários `LAUFI` (lotes no mesmo dia +
pagamentos *off-cycle* individuais); parte do run âncora e junta, *greedy*, os
runs cuja **soma REGU por PERNR reproduz o `/559`** sem introduzir divergências
→ identidade do colaborador (`PERNR` directo / `EMPFG` = PERNR / via `LIFNR` =
`HYPOTHESIS`) → valor RH esperado (`/559` do SEQNO mais recente transferido;
`/559` com `TSLIN=0` = `PREVIOUS_VERSION`, não soma) → matching em níveis →
`STATUS` por PERNR (`EXACT_MATCH`/`DIFFERENCE`/`RH_ONLY`/`REGU_ONLY`/`AMBIGUOUS`)
com tolerância **0,00**. Montantes `REGUH.RBETR/RWBTR` vêm com sinal à direita
(`1854.80-`) → parse via `sap_str_to_decimal` + `abs()`. `427,74` só é referido
como `[CANDIDATE]` se surgir naturalmente.

Output: `output/payroll_regu_reconciliation_0000001298.json` (inclui
`selected_payment_run_set`) / `.csv`. Conclusão: `EXACT_MATCH` / `DIFFERENCE`
/ `PARTIAL`.

### Resultado (run 0000001298, empresa 1010, período 202606) — `[PROVED]`

| | |
|---|---|
| Payment runs | `10243P` + `10311P` (ambos `LAUFD 20260625`) + 5 runs *off-cycle* de 1 pagamento (`09454P` 12/06, `10475P` 17/06, `10595P` 23/06, `11044P` 30/06, `10085P` 02/07) |
| Identidade | `REGUH.PERNR` directo (`PERNR_FIELD`, 541/541 coincidem) — `PERNR == LIFNR` numérico |
| RH esperado (Σ `/559` corrente) | **724.033,67 EUR** — 324 colaboradores |
| REGU pago (Σ `RBETR`) | **724.033,67 EUR** — 324 beneficiários |
| Diferença | **0,00** |
| `EXACT_MATCH` | **324 / 324** — 0 `DIFFERENCE`, 0 `RH_ONLY`, 0 `REGU_ONLY` |
| Múltiplos pagamentos por PERNR | 0 (os 2 lotes de 25/06 particionam a população; não há net dividido) |
| `427,74` | não aplicável — reconciliação fecha a 0,00 |

**RH × REGU: `EXACT_MATCH`.** O que o RH mandou pagar por colaborador (`/559`
corrente) é exactamente o que o programa de pagamentos levou às `REGU*`.
Os runs `10260P` (16.128,84) e `10320P` (9.778,23) de 25/06 **não** são o net do
payroll (pagamentos de outra natureza) e são correctamente excluídos.

## Próximo passo

1. Correr a Fase 5.0 no R/3 (repor `SAP_R3_ASHOST/SYSNR/CLIENT` no `.env`) e
   responder: `SUM(/559) == SUM(REGU)` e por colaborador?
2. (adiado) `PC00_M99_CWTR` só se a regra do `RPCIPE00` (delta 265,65) for
   mesmo necessária.
3. Só depois: reconciliação com o S/4H / conta 23120000.
