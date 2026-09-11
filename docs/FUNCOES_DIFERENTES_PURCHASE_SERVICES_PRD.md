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

## Limpeza de Atribuições Expiradas no CUA (Concluída em 11/09/2026)

Em 11/09/2026, foi realizada a limpeza integral das funções legadas e expiradas (fora do modelo e fora das regras da folha `EXCLUÇÃO`) no SAP CUA (`SPA` / `SU01`), abrangendo todos os utilizadores do departamento **Purchase & Services**:

- **Total de funções eliminadas no CUA**: **49 funções**.
- **Registo na folha `CUA_REMOVE`**: Todas as 49 funções foram adicionadas com status `CONCLUÍDO` (IDs `506` a `554`, linhas `507` a `555`) no ficheiro Excel oficial (`S4H_Perfis de autorização.xlsx`).

### Resumo da Limpeza por Utilizador:
- **`S965` (Dulce Guimarães)**: 10 funções eliminadas (IDs `506` a `515`)
- **`S75` (Carla Costa)**: 2 funções eliminadas (IDs `516` a `517`)
- **`S270` (Cidália Oliveira)**: 5 funções eliminadas (IDs `518` a `522`)
- **`S419` (Conceição Cunha)**: 14 funções eliminadas (IDs `523` a `536`)
- **`S80000148` (Catarina Faia)**: 18 funções eliminadas (IDs `537` a `554`)
- **`S170` (Mónica Rodrigues)**: 0 funções expiradas (100% conforme)
- **`S80001870` (Pedro Matos)**: 0 funções (conta inativa em PRD)

### Proteção Rigorosa Aplicada:
- **`EXCLUÇÃO`**: As funções `ZFIN_PS_BASIC`, `ZFIN_PS_SPEC` e `ZMM_APROVA_PEDC_COD_*` foram rigorosamente preservadas em todos os utilizadores.
- **`Proposta Ativa`**: O nível organizacional `ZORG_TODAS_EMPRESAS` (válido até 99991231) foi integralmente preservado para todos os utilizadores da proposta.
- **Atribuições Ativas Legadas**: As 3 atribuições ativas fora do modelo (`Z_COSTCENTER_CREATE` em `S270`, `ZORG_CENTROS_2XXX` e `ZORG_CENTROS_SALSA` em `S419`) foram mantidas e aguardam aprovação funcional.

