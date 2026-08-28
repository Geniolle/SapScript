# Regras críticas deste projeto

## Nunca eliminar sem pedido explícito
Nunca executar operações destrutivas em SAP (ex.: `PRGN_ACTIVITY_GROUP_DELETE`, apagar
funções/roles, apagar registos, remover transportes) por iniciativa própria — mesmo que
pareçam dados de teste óbvios ("pode apagar" na descrição, funções `Z_TESTE_*`, etc.) e
mesmo que uma eliminação semelhante tenha sido pedida antes na mesma conversa.

Cada eliminação exige um pedido explícito e específico do utilizador, feito naquele
momento. Nunca assumir consentimento antecipado nem estender uma autorização anterior
para um novo conjunto de objetos.

Isto aplica-se a qualquer ambiente (DEV incluído), mesmo em ambientes só usados para
testes.

## Cuidado com scripts monolíticos
Ao alterar templates HTML com um único bloco `<script>` grande, validar sempre a
sintaxe do JavaScript depois da edição. Um `if`, vírgula ou chaveta fora do lugar
pode quebrar o parse de toda a página e impedir que áreas independentes, como a
lista de tickets, carreguem.
