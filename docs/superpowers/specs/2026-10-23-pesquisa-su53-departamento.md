# Pesquisa SU53 por departamento

## Objetivo

Permitir pesquisar falhas SU53 por utilizador ou por todos os utilizadores de um departamento.

## Comportamento

- Manter argumentos CLI atuais como pesquisa direta por utilizador.
- Sem utilizador CLI, mostrar `[1] Utilizador` e `[2] Departamento`.
- Ler a lista de departamentos exclusivamente da sheet `CONTROLO`.
- Mostrar o número de utilizadores de cada departamento na sheet `Proposta Ativa`.
- Ler utilizador SAP e nome pela coluna `DEPARTAMENTO` da `Proposta Ativa`.
- Consultar todos os utilizadores sequencialmente no ambiente escolhido.
- Continuar quando uma consulta RFC falhar.
- Mostrar resumo com utilizadores sem erros, com erros e com falha RFC.
- Para utilizadores com erros, mostrar aplicação, objeto, campos e funções da sheet `Proposta`.
- O lote departamental é somente leitura.
- Após o resumo, permitir selecionar um utilizador com erros para abrir o fluxo individual existente.
