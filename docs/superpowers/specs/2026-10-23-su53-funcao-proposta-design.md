# Relatorio SU53 com Funcao da Proposta

## Intencao

Ao diagnosticar um erro SU53, mostrar a aplicacao que originou a falha e a funcao individual correspondente na sheet `Proposta`.

## Requisitos

- Usar `P_TCODE` quando preenchido; usar `NAME` como fallback.
- Procurar a aplicacao na sheet `Proposta` do Excel mestre.
- Mostrar transacao, objeto, atividade, horario, funcao, descricao e linhas do Excel.
- Considerar somente falhas de hoje na instancia da ocorrencia mais recente, igual ao contexto atual mostrado pela SU53.
- Agrupar por aplicacao e objeto, sem misturar a mesma autorizacao entre transacoes diferentes.
- Expor os mesmos dados no terminal e no JSON consumido pelo Cockpit/API.
- Quando nao houver mapeamento, informar `Nao encontrada na sheet Proposta`.
- Permanecer somente leitura: nao alterar SAP nem Excel.

## Abordagens

- Escolhida: enriquecer o diagnostico na camada de servico e reutilizar a pesquisa existente da sheet `Proposta`. O terminal e o Cockpit recebem os mesmos campos.
- Rejeitada: formatar somente o script interativo. Deixaria Cockpit/API sem a informacao.
- Rejeitada: criar um segundo parser da sheet `Proposta`. Duplicaria regra ja existente.

## Limites

- A funcao encontrada indica o mapeamento da transacao no projeto. Nao prova que essa funcao contem o objeto de autorizacao recusado no SAP.
- Nenhuma atribuicao ou correcao automatica faz parte desta mudanca.
