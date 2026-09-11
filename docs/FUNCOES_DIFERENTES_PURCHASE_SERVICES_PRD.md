# Funções diferentes por utilizador — Purchase & Services

Comparação realizada em 11/09/2026 entre as atribuições ativas em `AGR_USERS`,
no PRD (mandante 100), e o modelo completo do ficheiro
`S4H_Perfis de autorização.xlsx`.

## Critério correto de cruzamento

Para cada utilizador, o conjunto esperado é construído a partir da `Proposta Ativa`
e expandido recursivamente usando:

- `PFCG_CREATE`: catálogo das funções individuais criadas;
- `PFCG_COMPOSTA`: membros de cada função composta atribuída;
- `PFCG_AUTHORITY`: funções avulsas associadas às respetivas compostas;
- `EXCLUÇÃO`: padrões que ficam fora da análise e nunca são candidatos a remoção.

As regras da folha `EXCLUÇÃO` são `ZMM_APROVA_PEDC_COD_*`, `ZFIN_PS_BASIC`,
`ZFIN_PS_SPEC`, `Z_MY_HOME` e `SAP_*` (incluindo todas as funções técnicas/standard SAP).

## Resultado

| Utilizador | Funções esperadas | Em falta | Adicionais candidatas | Protegidas por EXCLUÇÃO |
|---|---:|---:|---:|:---|
| `S170` | 48 | 0 | 0 | `Z_MY_HOME` (1) |
| `S270` | 48 | 0 | 1 | `Z_MY_HOME` (1) |
| `S419` | 48 | 0 | 2 | `SAP_*` (2), `Z_MY_HOME` (1) |
| `S75` | 50 | 0 | 0 | `ZFIN_PS_SPEC`, `ZMM_APROVA_*`, `Z_MY_HOME` (7) |
| `S80000148` | 48 | 0 | 0 | `ZMM_APROVA_*`, `Z_MY_HOME` (5) |
| `S965` | 48 | 0 | 0 | `SAP_*` (2), `Z_MY_HOME` (1) |
| **Total** | **290** | **0** | **3** | **20 ocorrências protegidas** |

`S80001870` não integra o resultado dos utilizadores ativos porque a sua validade no mestre `USR02`
terminou em 29/07/2026 (conta inativa por desativação/offboarding).

## Candidatas por utilizador

### S170 — Monica Rodrigues
Nenhuma função candidata a remoção (100% conforme).

### S270 — Cidália Oliveira (1 adicional)
- `Z_COSTCENTER_CREATE` (Transações: `KS01`, `KS02`, `KS03`, `KS04` — Criação de Centros de Custo)

### S419 — Conceição Cunha (2 adicionais)
- `ZORG_CENTROS_2XXX` (Nível Organizacional de Centros 2XXX)
- `ZORG_CENTROS_SALSA` (Nível Organizacional de Centros Salsa)
*(As funções `SAP_BR_TRD_CLS_SPECIALIST` e `SAP_FND_BCR_MANAGER_T` foram protegidas pela regra `SAP_*`)*

### S75 — Carla Costa
Nenhuma função candidata a remoção (100% conforme).

### S80000148 — Catarina Faia
Nenhuma função candidata a remoção (100% conforme).

### S965 — Dulce Guimarães
Nenhuma função candidata a remoção (100% conforme — `SAP_BR_TRD_CLS_SPECIALIST` e `SAP_FND_BCR_MANAGER_T` protegidas por `SAP_*`).

## Segurança

Estas 3 atribuições restantes estão fora do conjunto derivado para o respetivo utilizador,
mas exigem aprovação funcional antes de qualquer remoção no CUA. Nenhuma alteração
foi efetuada no CUA durante esta análise.
