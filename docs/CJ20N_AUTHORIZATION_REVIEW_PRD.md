# Revisão de autorização da CJ20N em PRD

Data da validação: 2026-09-14

## Escopo

Validação somente leitura da segregação para liberação de projetos pela CJ20N no cliente PRD 100.

## Constatações

- A CJ20N está presente, entre outras, nas funções `ZFIN_PS_MANAGER` e `ZFIN_PS_SPEC`.
- `ZFIN_PS_MANAGER` possui valores explícitos de `PS_ACTVT` nos objetos PS; `ZFIN_PS_SPEC` possui `PS_ACTVT = *` em vários desses objetos.
- O perfil de status de utilizador `ZPS00000` contém:
  - `E0001` / `INIT` / Inicial;
  - `E0002` / `ENCE` / Encerrado.
- Ambos os status estão sem chave de autorização (`BERSL`).
- Não há controlo de operações empresariais configurado para `ZPS00000` em `TJ31`.
- A documentação SAP informa que não existe uma verificação de autorização standard específica para definir status de sistema.

## Conclusão

Atualmente, `B_USERSTAT` não implementa a segregação da operação de liberação para o perfil `ZPS00000`. A presença isolada do objeto na função não impede outros utilizadores de definir o status de sistema `LIB/REL`.

Para implementar a segregação pelo standard, deve ser desenhado um fluxo de status de utilizador que:

1. bloqueie a operação empresarial de liberação no status inicial;
2. disponibilize um status de aprovação com chave de autorização;
3. conceda essa chave, por `B_USERSTAT`, apenas à função responsável pela liberação;
4. permita a operação de liberação somente após a transição autorizada.

Como alternativa, uma ampliação pode executar uma verificação explícita de objeto de autorização no momento da liberação.

Qualquer alteração deve ser criada e testada em ambiente não produtivo, seguida de teste positivo e negativo com utilizadores distintos.
