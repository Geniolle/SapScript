# Regras críticas deste projeto

## Comando de encerramento "Perfeito"
Quando o utilizador disser **"Perfeito"** como aprovação/encerramento do trabalho atual,
atualizar a memória/documentação relevante, validar as alterações, criar um commit
descritivo e publicar a branch atual no GitHub. Não incluir credenciais, `.env` ou
artefactos locais. Não acionar a regra quando a palavra aparecer apenas numa citação,
exemplo ou pergunta.

## Acesso SAP via RFC
Assumir que as credenciais RFC SAP estao disponiveis no `.env` deste repositorio e tentar a leitura via RFC primeiro.
Se a ligacao falhar, reportar o erro tecnico concreto em vez de pedir novamente "ligacao" ou "credenciais".
Nao expor valores sensiveis do `.env` nas respostas.

Ao abrir ou mostrar uma sessão SAP GUI, posicionar a janela na metade esquerda do ecrã,
salvo indicação diferente do utilizador.

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

`node --check` valida apenas sintaxe. Erros de *runtime* no nível de topo do script
(ex.: usar um `const` antes de o declarar — Temporal Dead Zone) também abortam a
página inteira e não são detetados pelo `--check`. Depois de editar, carregar a
página com a consola do browser aberta.

## Erros já resolvidos
Antes de investigar um problema do cockpit/scripts, consultar
[docs/ERROS_RESOLVIDOS.md](docs/ERROS_RESOLVIDOS.md) — registo de incidentes
ultrapassados com sintoma, causa raiz, correção e método de diagnóstico.
Adicionar uma entrada nova sempre que se resolver um bug não trivial.
