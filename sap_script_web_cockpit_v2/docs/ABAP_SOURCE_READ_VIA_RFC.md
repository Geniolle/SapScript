# Leitura de source ABAP via RFC

Este projeto consegue ler código ABAP remoto quando o ambiente SAP tem
credenciais RFC válidas e o runtime com `pyrfc` disponível.

## Função usada

O caminho que funcionou no QAD foi:

- `RPY_PROGRAM_READ`

Parâmetros principais usados:

- `PROGRAM_NAME`
- `LANGUAGE`
- `WITH_INCLUDELIST = 'X'`
- `ONLY_SOURCE = 'X'`
- `READ_LATEST_VERSION = 'X'`
- `WITH_LOWERCASE = 'X'`

## Resultado prático

Em alguns programas SAP, o `SOURCE` principal vem vazio e a lógica real está
distribuída por includes. Nesses casos:

- ler `INCLUDE_TAB`
- repetir `RPY_PROGRAM_READ` para cada include relevante
- procurar os pontos funcionais por palavras-chave

## O que aprendemos

Para o caso do `RFF110S`, o source principal veio vazio e os includes
principais foram:

- `RFF110S_DATA`
- `RFF110S_FORMS`
- `RFF110S_SELSCR_BOE`
- `RKASMAWF`
- `SCHEDMAN_EVENTS`

## Boas práticas

- Não expor credenciais RFC em logs ou em Markdown.
- Usar um ambiente de diagnóstico isolado.
- Guardar os snippets relevantes numa nota específica do programa analisado.
- Preferir `RPY_PROGRAM_READ` quando o objetivo for ler programa completo e
  includes.
- Se a leitura do source falhar, validar primeiro:
  - ligação RFC
  - permissões
  - nome do programa
  - idioma
  - se o source está em include e não no report principal

