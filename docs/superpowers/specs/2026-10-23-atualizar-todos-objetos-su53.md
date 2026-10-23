# Atualizar todos os objetos SU53

## Objetivo

Permitir processar todos os objetos de autorização apresentados no menu SU53 sem selecionar cada item manualmente.

## Comportamento

- Mostrar `[T] Atualizar TODOS` apenas quando existirem vários objetos.
- Pedir uma única confirmação antes do lote.
- Executar os objetos sequencialmente pela rotina existente.
- Manter pré-validação, atualização e validação RFC independentes por objeto.
- Continuar o lote quando um objeto falhar.
- Mostrar resumo final com sucessos e falhas.
- Não fundir ocorrências repetidas, pois aplicações diferentes podem mapear funções diferentes.
- Não executar alterações SAP durante testes automatizados.
