# Layout padrão — análise SU53

Formato para cada análise de erros de autorização (SU53) por departamento ou utilizador.

## Cabeçalho

```
Departamento: <nome> | Ambiente: <PRD|QAS> | Data: <dd.mm.aaaa>
Utilizadores analisados: <n> | Com erros: <n> | Sem erros: <n> | Falhas RFC: <n>
```

## Por utilizador (só com erros)

```
<USER> – <Nome> (<n> registos hoje)
  Erros:
    - <OBJETO> <campos/valores> × <n>
    - <OBJETO> <campos/valores> × <n>
  Função sugerida: <Z_ROLE> (<Folha>, linha <n>) | Descrição: <texto>
```

Regras:
- Erros por utilizador, sem comentários nem explicações.
- Uma linha por combinação objeto + campos/valores, com contagem.
- `S_TCODE`: indicar `TCD=<transação>`.
- `S_APPL_LOG`, `S_ADMI_FCD` e similares: indicar campos completos.
- Função sugerida: a Função Z_ da folha `Proposta` (ou `Single Roles (S4H)` / `PFCG_CREATE`) que contém a transação ou objeto. Se não houver, escrever `Sem função no catálogo`.
- Dados sempre do dia da análise, salvo indicação em contrário.
