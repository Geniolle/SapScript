# DMEEX CtgyPurp/Cd SUPP - PRD Z_SEPA_CT vs DEV Z_PT_CGI_XML_CT_V9

Gerado em: 2026-09-24 16:05:56

Arvore esperada:

```text
PmtInf
`- PmtTpInf
   `- CtgyPurp
      `- Cd
         |- CASH
         `- SUPP
```

### PRD Z_SEPA_CT

| NODE_ID | path | type | map | const | field | exit | cond | flags |
|---|---|---|---|---|---|---|---|---|
| N_8186402610 | /Document/CstmrCdtTrfInitn/PmtInf | ELEM | STRUCTURE |  |  |  |  | DEFINED_IN_Z_SEPA_CT |
| N_1139725890 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf | ELEM | STRUCTURE |  |  |  |  | DEFINED_IN_Z_SEPA_CT |
| N_9359036710 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf | ELEM | STRUCTURE |  |  |  | = 'X' | DEFINED_IN_Z_SEPA_CT |
| N_4731479220 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/(TECH:CtgyPurp) | TECH | EXIT:DMEE_EXIT_SEPA_COUNTRIES |  |  | DMEE_EXIT_SEPA_COUNTRIES |  | DEFINED_IN_Z_SEPA_CT |
| N_3282960100 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp | ELEM | STRUCTURE |  |  |  | = 'X' | DEFINED_IN_Z_SEPA_CT |
| N_6497320730 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd | ELEM | EXIT:DMEE_EXIT_SEPA_COUNTRIES |  |  | DMEE_EXIT_SEPA_COUNTRIES |  | DEFINED_IN_Z_SEPA_CT |
| N_6181870380 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Prtry | ELEM | EXIT:DMEE_EXIT_SEPA_COUNTRIES |  |  | DMEE_EXIT_SEPA_COUNTRIES |  | DEFINED_IN_Z_SEPA_CT |
| N_5863659600 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf | ELEM | STRUCTURE |  |  |  | = SPACE | DEFINED_IN_Z_SEPA_CT |
| N_4702656500 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/(TECH:CtgyPurp) | TECH | EXIT:Z_DMEE_EXIT_SEPA_COUNTRIES |  |  | Z_DMEE_EXIT_SEPA_COUNTRIES |  | DEFINED_IN_Z_SEPA_CT |
| N_3484607870 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp | ELEM | EXIT:DMEE_EXIT_SEPA_COUNTRIES |  |  | DMEE_EXIT_SEPA_COUNTRIES | = 'X' | DEFINED_IN_Z_SEPA_CT |
| N_7593772580 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd | ELEM | CONST:SUPP | SUPP |  |  |  | DEFINED_IN_Z_SEPA_CT |
| N_2458843850 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:CASH | ATOM | CONST:CASH | CASH |  |  | FPAYP-GPA2R >= 999999 | DEFINED_IN_Z_SEPA_CT |
| N_2761956870 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:SUPP | ATOM | CONST:SUPP | SUPP |  |  |  | DEFINED_IN_Z_SEPA_CT |
| N_2559985060 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Prtry | ELEM | EXIT:DMEE_EXIT_SEPA_COUNTRIES |  |  | DMEE_EXIT_SEPA_COUNTRIES |  | DEFINED_IN_Z_SEPA_CT |

### DEV Z_PT_CGI_XML_CT_V9

| NODE_ID | path | type | map | const | field | exit | cond | flags |
|---|---|---|---|---|---|---|---|---|
| N_4597540400 | /Document/CstmrCdtTrfInitn/PmtInf | ELEM | STRUCTURE |  |  |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_7436948640 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf | ELEM | STRUCTURE |  |  |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_0386175880 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf | ELEM | STRUCTURE |  |  |  | FPAYHX-REF03+0(1) <> 'S' AND FPAYHX-XSCHK <> 'X' | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_7877104190 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp | ELEM | STRUCTURE |  |  |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_6988727120 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd | ELEM | FIELD:FPAYHX-REF06+079 |  | FPAYHX-REF06+079 |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_0081587289 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd/:atom:CODE2 | ATOM | FIELD:FPAYHX-CODE2 |  | FPAYHX-CODE2 |  | FPAYH-DORIGIN <> 'HR-PY' | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_0119167347 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd/:atom:HR_CODE | ATOM | FIELD:FPAYH-PURP_CODE |  | FPAYH-PURP_CODE |  | FPAYH-DORIGIN = 'HR-PY' | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_8323416300 | /Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Prtry | ELEM | STRUCTURE |  |  |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_2562711940 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf | ELEM | STRUCTURE |  |  |  | FPAYHX-REF03+0(1) = 'S' AND FPAYHX-XSCHK = SPACE | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_0584949040 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp | ELEM | STRUCTURE |  |  |  |  | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_7655237210 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd | ELEM | CONST:SUPP | SUPP |  |  |  | REDEF=X,REDEFINED_IN_Z_PT_CGI_XML_CT_V9 |
| N01623664445 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:CASH | ATOM | CONST:CASH | CASH |  |  | FPAYP-GPA2R >= 999999 | DEFINED_IN_Z_PT_CGI_XML_CT_V9 |
| N_0505655476 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:CODE2 | ATOM | FIELD:FPAYHX-CODE2 |  | FPAYHX-CODE2 |  | FPAYH-DORIGIN <> 'HR-PY' | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N_0928715854 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:HR_CODE | ATOM | FIELD:FPAYH-PURP_CODE |  | FPAYH-PURP_CODE |  | FPAYH-DORIGIN = 'HR-PY' | INHERITED_COPY_IN_Z_PT_CGI_XML_CT_V9 |
| N01240798012 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Cd/:atom:SUPP | ATOM | CONST:SUPP | SUPP |  |  |  | DEFINED_IN_Z_PT_CGI_XML_CT_V9 |
| N_5119155690 | /Document/CstmrCdtTrfInitn/PmtInf/PmtTpInf/CtgyPurp/Prtry | ELEM | STRUCTURE |  |  |  |  | REDEF=X,DEACT=X,DEACTIVATED_IN_Z_PT_CGI_XML_CT_V9 |

## Diferencas relevantes

- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/(TECH:-PmtTpInf)` PRD_NODE=N_1836464000 DEV_NODE= 
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/(TECH:PmtTpInf)` PRD_NODE=N_9897020250 DEV_NODE= 
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/CdtrAcct/(TECH:IbanForCashPayments)` PRD_NODE=N_5193282350 DEV_NODE= 
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/CdtrAgt/(TECH:BICForCashPayments)` PRD_NODE=N_2609636910 DEV_NODE= 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf` PRD_NODE=N_9359036710 DEV_NODE=N_0386175880 PARENT_ID: PRD='N_1139725890' DEV='N_7436948640', BROTHER_ID: PRD='N_6058284380' DEV='N_4239787460', FIRSTCHILD_ID: PRD='N_2528812510' DEV='N_0022976600', conditions_key: PRD='[{"ARG1_CONST": "", "ARG1_FLD": "", "ARG1_TAB": "", "ARG1_TYPE": "3", "ARG2_CONST": "'X'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "", "NODE_ID": "N_9359036710", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]' DEV='[{"ARG1_CONST": "", "ARG1_FLD": "REF03+0(1)", "ARG1_TAB": "FPAYHX", "ARG1_TYPE": "2", "ARG2_CONST": "'S'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "AND", "NODE_ID": "N_0386175880", "OPERATOR": "<>", "PAR_CLOSE": "", "PAR_OPEN": ""}, {"ARG1_CONST": "", "ARG1_FLD": "XSCHK", "ARG1_TAB": "FPAYHX", "ARG1_TYPE": "2", "ARG2_CONST": "'X'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "002", "LINK_OPERATOR": "", "NODE_ID": "N_0386175880", "OPERATOR": "<>", "PAR_CLOSE": "", "PAR_OPEN": ""}]'
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/(TECH:-SvcLvlCdtr)` PRD_NODE=N_1610764260 DEV_NODE= 
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/(TECH:CtgyPurp)` PRD_NODE=N_4731479220 DEV_NODE= 
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/(TECH:LclInstrm)` PRD_NODE=N_6104007050 DEV_NODE= 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp` PRD_NODE=N_3282960100 DEV_NODE=N_7877104190 PARENT_ID: PRD='N_9359036710' DEV='N_0386175880', FIRSTCHILD_ID: PRD='N_6497320730' DEV='N_6988727120', LEV: PRD='000' DEV='003', conditions_key: PRD='[{"ARG1_CONST": "", "ARG1_FLD": "", "ARG1_TAB": "", "ARG1_TYPE": "3", "ARG2_CONST": "'X'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "", "NODE_ID": "N_3282960100", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]' DEV='[]'
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd` PRD_NODE=N_6497320730 DEV_NODE=N_6988727120 PARENT_ID: PRD='N_3282960100' DEV='N_7877104190', BROTHER_ID: PRD='N_6181870380' DEV='N_8323416300', FIRSTCHILD_ID: PRD='' DEV='N_0081587289', LENGTH: PRD='0000' DEV='0004', LEV: PRD='000' DEV='003', MP_IF_TP: PRD='' DEV='1', MP_SELECTION: PRD='5' DEV='6', MP_SC_TAB: PRD='' DEV='FPAYHX'
- SO_DIREITA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd/:atom:CODE2` PRD_NODE= DEV_NODE=N_0081587289 
- SO_DIREITA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Cd/:atom:HR_CODE` PRD_NODE= DEV_NODE=N_0119167347 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/CtgyPurp/Prtry` PRD_NODE=N_6181870380 DEV_NODE=N_8323416300 PARENT_ID: PRD='N_3282960100' DEV='N_7877104190', LENGTH: PRD='0000' DEV='0035', LEV: PRD='000' DEV='003', MP_SELECTION: PRD='5' DEV='1', MP_EXIT_FUNC: PRD='DMEE_EXIT_SEPA_COUNTRIES' DEV=''
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/InstrPrty` PRD_NODE=N_2528812510 DEV_NODE=N_0022976600 PARENT_ID: PRD='N_9359036710' DEV='N_0386175880', BROTHER_ID: PRD='N_1610764260' DEV='N_6348225970', FIRSTCHILD_ID: PRD='' DEV='N_4201182960', LEV: PRD='000' DEV='003', MP_IF_TP: PRD='' DEV='1', MP_SELECTION: PRD='5' DEV='6', MP_SC_TAB: PRD='' DEV='FPAYHX', MP_SC_FLD: PRD='' DEV='DTURG'
- SO_DIREITA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/InstrPrty/:atom:HIGH` PRD_NODE= DEV_NODE=N_3780541170 
- SO_DIREITA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/InstrPrty/:atom:NORM` PRD_NODE= DEV_NODE=N_4201182960 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/LclInstrm` PRD_NODE=N_5778091230 DEV_NODE=N_1131109450 PARENT_ID: PRD='N_9359036710' DEV='N_0386175880', BROTHER_ID: PRD='N_4731479220' DEV='N_7877104190', FIRSTCHILD_ID: PRD='N_1317460400' DEV='N_4797314670', LEV: PRD='000' DEV='003', conditions_key: PRD='[{"ARG1_CONST": "", "ARG1_FLD": "", "ARG1_TAB": "", "ARG1_TYPE": "3", "ARG2_CONST": "'X'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "", "NODE_ID": "N_5778091230", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]' DEV='[]'
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/LclInstrm/Cd` PRD_NODE=N_1317460400 DEV_NODE=N_4797314670 PARENT_ID: PRD='N_5778091230' DEV='N_1131109450', BROTHER_ID: PRD='N_6402352590' DEV='N_4074350710', LENGTH: PRD='0000' DEV='0035', LEV: PRD='000' DEV='003', MP_IF_TP: PRD='' DEV='1', MP_SELECTION: PRD='5' DEV='1', MP_EXIT_FUNC: PRD='DMEE_EXIT_SEPA_COUNTRIES' DEV=''
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/LclInstrm/Prtry` PRD_NODE=N_6402352590 DEV_NODE=N_4074350710 PARENT_ID: PRD='N_5778091230' DEV='N_1131109450', LENGTH: PRD='0000' DEV='0035', LEV: PRD='000' DEV='003', MP_SELECTION: PRD='5' DEV='1', MP_EXIT_FUNC: PRD='DMEE_EXIT_SEPA_COUNTRIES' DEV=''
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl` PRD_NODE=N_6045176050 DEV_NODE=N_6348225970 PARENT_ID: PRD='N_9359036710' DEV='N_0386175880', BROTHER_ID: PRD='N_6104007050' DEV='N_1131109450', FIRSTCHILD_ID: PRD='N_6392975760' DEV='N_1955869580', LEV: PRD='000' DEV='003', conditions_key: PRD='[{"ARG1_CONST": "", "ARG1_FLD": "", "ARG1_TAB": "", "ARG1_TYPE": "3", "ARG2_CONST": "SPACE", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "AND", "NODE_ID": "N_6045176050", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}, {"ARG1_CONST": "", "ARG1_FLD": "", "ARG1_TAB": "", "ARG1_TYPE": "3", "ARG2_CONST": "SPACE", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "002", "LINK_OPERATOR": "", "NODE_ID": "N_6045176050", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]' DEV='[]'
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl/(TECH:-Cd)` PRD_NODE=N_6392975760 DEV_NODE= 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl/Cd` PRD_NODE=N_3661101640 DEV_NODE=N_1955869580 PARENT_ID: PRD='N_6045176050' DEV='N_6348225970', BROTHER_ID: PRD='N_8821950710' DEV='N_1978739260', FIRSTCHILD_ID: PRD='N_5801142520' DEV='N_0271455296', LENGTH: PRD='0000' DEV='0004', LEV: PRD='000' DEV='003', MP_IF_TP: PRD='' DEV='1', MP_SC_TAB: PRD='' DEV='FPAYHX', MP_SC_FLD: PRD='' DEV='CODE1'
- SO_ESQUERDA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl/Cd/:atom:Instruction` PRD_NODE=N_2692067310 DEV_NODE= 
- SO_DIREITA `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl/Cd/:atom:Non_SEPA` PRD_NODE= DEV_NODE=N_0271455296 
- DIFERENTE `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtTpInf/SvcLvl/Cd/:atom:SEPA` PRD_NODE=N_5801142520 DEV_NODE=N_0212834634 PARENT_ID: PRD='N_3661101640' DEV='N_1955869580', BROTHER_ID: PRD='N_2692067310' DEV='', conditions_key: PRD='[{"ARG1_CONST": "", "ARG1_FLD": "DTWS1", "ARG1_TAB": "FPAYH", "ARG1_TYPE": "2", "ARG2_CONST": "SPACE", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "", "NODE_ID": "N_5801142520", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]' DEV='[{"ARG1_CONST": "", "ARG1_FLD": "CODE1", "ARG1_TAB": "FPAYHX", "ARG1_TYPE": "2", "ARG2_CONST": "SPACE", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "002", "LINK_OPERATOR": "", "NODE_ID": "N_0212834634", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}, {"ARG1_CONST": "", "ARG1_FLD": "REF03+0(1)", "ARG1_TAB": "FPAYHX", "ARG1_TYPE": "2", "ARG2_CONST": "'S'", "ARG2_FLD": "", "ARG2_TAB": "", "ARG2_TYPE": "1", "CD_EXIT_FUNC": "", "COND_NUMBER": "001", "LINK_OPERATOR": "AND", "NODE_ID": "N_0212834634", "OPERATOR": "=", "PAR_CLOSE": "", "PAR_OPEN": ""}]'

## Conclusao

Situacao 2: existe diferenca concreta de DMEEX. O SUPP existe e esta ativo no DEV, mas o ramo DEV PmtInf/PmtTpInf so e avaliado quando FPAYHX-REF03+0(1) = 'S' e FPAYHX-XSCHK = SPACE; alem disso o Cd DEV tem propriedades tecnicas diferentes do PRD e atomos herdados adicionais CODE2/HR_CODE. Efeito esperado: se essa condicao de ramo nao for verdadeira no pagamento, CtgyPurp/Cd=SUPP nao e produzido apesar de o no existir.
