# Análise técnica read-only — `/SBXC/ZCKPRLT01` (PRD)

## Âmbito confirmado

- Extração via RFC em modo exclusivamente de leitura, usando `RPY_PROGRAM_READ`, `RFC_READ_TABLE` e `RFC_PING`.
- 6 objetos do programa: report principal e includes `TOP`, `LCL`, `F01`, `O01`, `I01`.
- 5.956 linhas ABAP extraídas.
- Fontes adicionais extraídas para 19 classes globais referidas pelo código, incluindo `/SBXC/ZCL_READ_DOC_SAPHETY` e `/SBXC/CO_ICNDDOCUMENT_SERVICE`.
- Nenhuma chamada literal `BAPI_*`, `GET BADI`, `CALL BADI` ou enhancement explícito foi encontrada no report/includes.

## Finalidade e fluxo

O programa é um cockpit genérico para processamento de faturas provenientes de sistemas externos. O fluxo principal é:

1. `VALIDA_PROCESSOS`: lê `/SBXC/ZCKP_TAB00` e delega a autorização a `/SBXC/ZCKP_VALIDA_PROCESSO`.
2. `GET_CONFIGURATION`: carrega DDIC e parametrizações (`TAB04`, `TAB05`, `TAB10`, textos e controlos).
3. `GET_DATA`: lê dados do processo.
4. `INICIALIZA_CONTROL`: cria os controlos GUI.
5. `BUILD_ALVS_INIT`: constrói ALVs dinâmicos para cabeçalho, linhas e estado.
6. `CALL SCREEN 100`: entra no cockpit dialog.

## Dependências relevantes

- Funções `/SBXC/` literais: `ZCKP_VALIDA_PROCESSO`, `ZCKP_BASELINE_DATE`, `ZCKP_DBL_CLK_LIN`, `ZCKP_LOAD_DD03L`, `ZCKP_LOAD_TAB10`.
- Funções standard relevantes: `DDIF_FIELDINFO_GET`, `FUNCTION_EXISTS`, `LVC_FIELDCATALOG_MERGE`, `FI_F4_MWSKZ`, conversões ALPHA/ABPSP, popups e `SCMS_XSTRING_TO_BINARY`.
- Classes custom: `/SBXC/ZCL_READ_DOC_SAPHETY` e `/SBXC/CO_ICNDDOCUMENT_SERVICE`.
- Framework/UI: `CL_GUI_ALV_GRID`, containers/splitters, `CL_GUI_HTML_VIEWER`, `CL_SALV_*` e RTTS `CL_ABAP_*DESCR`.
- Principais tabelas custom: `/SBXC/ZCKP_TAB00`, `TAB03`, `TAB04`, `TAB05`, `TAB09`, `TAB10`, `TAB11`, `TAB12`, `TAB30`, `/SBXC/ZCKP_CTRL`, `/SBXC/ZCKP_HMAIL`.
- Principais tabelas standard referidas: `BKPF`, `DD03L`, `LFA1`, `T001`, `T005`, `T052`, `AUFK`, `PRPS`, `EKPO`, `A003`, `KONP`.

## Dificuldades e riscos técnicos encontrados

1. **Despacho dinâmico de lógica**: funções como `FM_ALTERACAO`, `FM_CAB_DISP`, `FM_LIN_DISP`, `FM_DBLCLK` e funções pré/principal/pós são chamadas por nome guardado em configuração. Uma análise apenas do código estático não determina todo o comportamento efetivo de cada processo.
2. **SQL e tipagem dinâmicos**: o código usa tabelas `(STR_CAB)` e `(STR_LIN)`, `CREATE DATA`, `ASSIGN (campo)` e componentes por nome. Erros de customizing/DDIC aparecem apenas em runtime e reduzem a verificabilidade estática.
3. **Escrita dinâmica no programa produtivo**: existem `UPDATE (str_cab)` e `COMMIT WORK AND WAIT` no tratamento de comandos do utilizador, além da persistência delegada ao FM configurado em `FM_ALTERACAO`. A extração não executou nenhum destes caminhos.
4. **Autorização indireta**: não há `AUTHORITY-CHECK` literal no report extraído; o controlo é delegado a `/SBXC/ZCKP_VALIDA_PROCESSO` e às tabelas de permissões/configuração. É necessário analisar esse FM para conhecer o objeto/campos de autorização efetivos.
5. **Acoplamento a convenções de campos**: o cockpit espera componentes como `PROCESSO`, `SEQNO`, `STATUS1`, `PSTNG_DATE`, `COMP_CODE`, `DOC_FI`, `REF_DOC_NO`, `CELLSTYLE` e `CELL_COLOR`. Uma estrutura configurada sem estas convenções pode causar falhas de `ASSIGN`, dados incompletos ou comportamento silenciosamente omitido.
6. **Responsabilidades concentradas**: `F01` e `LCL` concentram UI, DDIC dinâmico, regras financeiras, navegação, atualização e visualização de documentos. Isto torna testes isolados e diagnóstico de incidentes mais difíceis.
7. **Código legado/comentado**: há múltiplos blocos desativados e duplicação de lógica (por exemplo, cálculo/atualização de datas e reconstrução de ALV), aumentando o risco de divergência entre caminhos.
8. **Inventário RFC parcial**: PRD permitiu a fonte via `RPY_PROGRAM_READ`, mas recusou `RFC_READ_TABLE` para `SEOCLASS` e `TFDIR` com `AD 718`. Assim, as fontes de classes foram obtidas pelos class pools, mas metadados DDIC de classes/FMs e where-used completo não puderam ser confirmados por essas tabelas.
9. **Configuração runtime ainda não materializada**: duas tentativas read-only de obter metadados completos da configuração ficaram sem resposta e foram interrompidas. Portanto, os nomes concretos de todos os FMs dinâmicos por processo ainda são a principal lacuna antes de uma análise funcional exaustiva.

## Próxima análise recomendada

Prioridade alta: obter, por processo, os campos de roteamento de `/SBXC/ZCKP_TAB00` e depois extrair os FMs dinâmicos correspondentes, começando por autorização, seleção, alteração e ações pré/principal/pós. Em seguida, mapear cada processo como `configuração -> estrutura CAB/LIN -> FM -> tabelas alteradas -> COMMIT/ROLLBACK -> mensagens`.

## Limite da conclusão

Esta primeira fase comprovou o código ativo do report e dos includes/classes alcançáveis pela leitura RFC. A configuração dinâmica e o function pool foram obtidos na fase complementar abaixo.

---

## Complemento — configuração runtime e function pool

A lacuna de configuração foi posteriormente resolvida por uma leitura direta e limitada de `/SBXC/ZCKP_TAB00`.

### Configuração efetiva em PRD

- Foram encontrados 28 processos configurados.
- Todos usam `EST_CAB = /SBXC/ZCKP_INVH` e `EST_LIN = /SBXC/ZCKP_INVI`.
- Todos usam os mesmos FMs centrais:
  - `FM_CAB_DISP = /SBXC/ZCKP_CAB_FAT`
  - `FM_LIN_DISP = /SBXC/ZCKP_ITEM_FAT`
  - `FM_DBLCLK = /SBXC/ZCKP_DBL_CLK`
  - `FM_ALTERACAO = /SBXC/ZCKP_GUARDA_ALT_FAT`
- Entre os processos encontram-se `ARQUIVO`, `BP_CORE`, `PO_CORE`, `PO_NONCORE`, `SAPHETY`, `INVENTARIO`, `FI_C_IMP`, `PO_C_IMP`, `FI_SALSA` e `FI_LOSAN`.

### Autorização efetiva

`/SBXC/ZCKP_VALIDA_PROCESSO` executa:

- objeto de autorização: `ZCKP:PROCS`;
- campo `/SBXC/CPRO`: processo solicitado;
- campo `/SBXC/CAUT`: `DUMMY`.

A leitura da existência/configuração do processo presente no FM está comentada. Assim, o FM valida a autorização, mas não confirma diretamente se o processo existe na `TAB00`; a estrutura recebida pelo chamador é preenchida previamente pelo próprio report.

### Function pool completo

O grupo `/SBXC/ZCKP_F`, programa gerado `/SBXC/SAPLZCKP_F`, foi extraído com:

- 89 programas/includes;
- 16.724 linhas ABAP;
- 85 módulos de função referidos;
- 9 BAPIs;
- 26 classes;
- 96 tabelas;
- 11 transações;
- 1 `SUBMIT`.

BAPIs encontradas no function pool:

- `BAPI_ACC_DOCUMENT_POST`
- `BAPI_ACC_INVOICE_RECEIPT_POST`
- `BAPI_INCOMINGINVOICE_CANCEL`
- `BAPI_INCOMINGINVOICE_CREATE`
- `BAPI_INCOMINGINVOICE_PARK`
- `BAPI_PO_CHANGE`
- `BAPI_PO_CREATE1`
- `BAPI_TRANSACTION_COMMIT`
- `BAPI_TRANSACTION_ROLLBACK`

Não foram encontradas chamadas explícitas `GET BADI` ou `CALL BADI`. Isto não exclui BADIs executadas internamente pelas BAPIs/transações standard.

### Ações configuradas na toolbar

`/SBXC/ZCKP_TAB03` devolveu 50 ações configuradas, sem erros RFC. Os FMs distintos incluem:

- associação e receção MM: `ZCKP_ASSOCIA_PED_M`, `ZCKP_MM_REGISTA_FAC_CKP`, `ZCKP_MM_LIGA_ULT_VARIOS`, `ZCKP_MM_LIGA_ULTERIOR`;
- contabilização: `ZCKP_BAPI_INVOICE`, `ZCKP_RFBIBL00`;
- rejeição/cancelamento: `ZCKP_REJ`, `ZCKP_CANCEL`, `ZCKP_CANC_LIGA_ULTERIOR`;
- alterações: `ZCKP_GUARDA_ALT_TRAT`, `ZCKP_ALT_MASSA_C`, `ZCKP_ALTERA_MASSA`, `ZCKP_ITEM_NEW`, `ZCKP_ITEM_COPY`, `ZCKP_ITEM_DEL`;
- documentos/imagem/email: `ZCKP_VER_DOC_FI`, `ZCKP_IMG`, `ZCKP_READ_EMAIL`, `ZCKP_SEND_EMAIL`, `ZCKP_SEND_EMAIL_LIFNR`;
- integração: `ZCKP_SENDSAPHETY`;
- apoio: `ZCKP_REFRESH`, `ZCKP_LOG_MESS`, `ZCKP_M_TUDO`, `ZCKP_DM_TUDO`, `ZCKP_INVERTER`.

O cockpit constrói automaticamente nomes `<FM>_PRE` e `<FM>_POS`, testa a existência e executa a sequência pré → principal → pós. Isto significa que uma única ação pode atravessar três módulos e múltiplas LUWs. As fontes correspondentes estão incluídas no function pool extraído.

## Problemas concretos nos quatro FMs centrais

### `/SBXC/ZCKP_GUARDA_ALT_FAT`

1. Atualiza `/SBXC/ZCKP_INVH` sem validar `SY-SUBRC` e sem tratamento de erro.
2. Faz `MODIFY /SBXC/ZCKP_INVI` seguido de `COMMIT WORK` dentro do loop de itens. Isto quebra atomicidade: uma falha posterior pode deixar apenas parte da fatura gravada.
3. Usa `SELECT MAX(INVOICE_DOC_ITEM) + 1` para criar o próximo item, sem lock/enqueue. Duas sessões concorrentes podem calcular o mesmo número.
4. Faz vários commits adicionais, incluindo `COMMIT WORK AND WAIT`, mas não existe `ROLLBACK WORK` no FM.
5. A eliminação de itens removidos só ocorre quando `ITEM[] IS NOT INITIAL`. Se o utilizador eliminar todos os itens, os registos antigos podem permanecer na base.
6. O `DELETE /SBXC/ZCKP_INVI` também não valida retorno nem agrega mensagens de erro.

### `/SBXC/ZCKP_CAB_FAT`

1. Um FM nominalmente de apresentação (`FM_CAB_DISP`) altera `/SBXC/ZCKP_CTRL` e `/SBXC/ZCKP_INVH`, e executa `COMMIT WORK`. Portanto, abrir/refrescar a apresentação pode produzir efeitos persistentes.
2. Sincroniza estornos e estados consultando `BKPF`/`RBKP`; falhas ou documentos parcialmente encontrados podem alterar o estado do cockpit durante leitura.
3. As atualizações não verificam `SY-SUBRC` nem têm rollback.
4. Há leituras repetidas por item sobre `EKPO`/`EKET`, com potencial custo elevado para faturas com muitas linhas.
5. A determinação de condições/impostos depende de dados atuais e não contém uma camada explícita de rastreabilidade da regra aplicada.

### `/SBXC/ZCKP_ITEM_FAT`

1. Executa vários `SELECT SINGLE` dentro de loops (`EKPO`, `A003`, `KONP`, `MEAN/MARA`), padrão N+1 com impacto de desempenho.
2. Calcula `BPMNG = QUANTITY * BPUMZ / BPUMN` sem proteção explícita contra `BPUMN = 0` ou não encontrado.
3. `LV_IND` é inicializado com zero e não é incrementado; além disso, `INVOICE_DOC_ITEM` é atribuído ao auxiliar depois do `MOVE-CORRESPONDING` para `LINHA`. A renumeração pretendida pode não chegar à linha devolvida.
4. `DELETE ADJACENT DUPLICATES` depende da ordenação e de uma chave parcial; pode ocultar linhas legítimas ou manter duplicados fora dessa ordenação.
5. Recalcula imposto a partir de `A003/KONP`, mas não filtra explicitamente validade temporal ou todas as dimensões de determinação de condição.

### `/SBXC/ZCKP_DBL_CLK`

1. O duplo clique não é apenas navegação: recalcula imposto e valor de controlo antes de tratar o campo clicado.
2. A seleção em `A003` pode devolver várias condições, que são somadas exceto `MWVI`; a regra precisa ser validada funcionalmente para cada país/código fiscal.
3. Para `ICON_EMAIL`, chama `/SBXC/ZCKP_READ_EMAIL`; para documentos, delega navegação a FORMs que abrem FI/MM.
4. Não há indicação explícita de erro quando documentos, fornecedor ou utilizador não são encontrados.

## Diagnóstico consolidado

A principal dificuldade não é uma BADI isolada. O cockpit combina, no mesmo function pool, apresentação ALV, autorização, cálculo fiscal, persistência direta, contabilização FI/MM, criação/alteração de pedidos, cancelamentos e integração SAPHETY. A configuração central reduz variação entre os 28 processos, mas concentra o risco nos quatro FMs comuns e nas ações de toolbar que chamam funções pré/principal/pós.

As prioridades técnicas são:

1. tornar `ZCKP_GUARDA_ALT_FAT` atómico, com enqueue, uma única LUW e rollback;
2. remover efeitos de escrita de `ZCKP_CAB_FAT` ou separá-los numa sincronização explícita;
3. corrigir numeração de itens e o cenário de eliminação total;
4. reduzir `SELECT` dentro de loops;
5. inventariar `TAB03` e respetivos FMs pré/principal/pós para mapear todas as ações disponíveis na toolbar;
6. criar testes por processo/status para os fluxos FI, MM, impostos e SAPHETY.
