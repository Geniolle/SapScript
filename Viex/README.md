# Validador VIES

Aplicação Windows profissional para validação de VAT (NIF) europeus através do serviço oficial VIES da União Europeia.

## Funcionalidades

### 1. Pesquisa Individual
- Valida um VAT por vez
- Suporta qualquer país europeu (ex: PT504772694, ESA28017895)
- Apresenta resultado detalhado com:
  - Validade do VAT
  - Código do país e número
  - Nome da empresa (quando disponível)
  - Morada (quando disponível)
  - Data da consulta no VIES

### 2. Pesquisa em Massa
- Processa múltiplos VATs num único Excel
- Atualização incremental: cada VAT é gravado após validação
- Barra de progresso em tempo real
- Tratamento de erros sem interrupção:
  - Falha técnica: regista como "ERRO"
  - VAT inválido: regista como "NÃO"
- Suporta países que não disponibilizam dados

## Requisitos

- Windows 7 ou superior
- Nenhuma instalação de Python ou dependências necessárias
- Acesso à internet (para consultar o serviço VIES)

## Instalação

1. Descarregue a pasta `ValidadorVIES` (resultado do build)
2. Copie para um local no seu computador (ex: `C:\Program Files\Validador VIES`)
3. Dê duplo clique em `ValidadorVIES.exe`

Não é necessário instalar mais nada.

## Utilização

### Pesquisa Individual
1. Selecione "Pesquisa Individual" no menu inicial
2. Informe o VAT (ex: PT504772694)
3. Clique em "Validar"
4. O resultado aparece na área de texto

### Pesquisa em Massa
1. Selecione "Pesquisa em Massa" no menu inicial
2. Clique em "Selecionar ficheiro"
3. Escolha um Excel (.xlsx) com uma sheet chamada "NIF"
4. A sheet deve ter colunas com os cabeçalhos:
   - `NIFS` (lista de VATs a validar)
   - `Válido no VIES` (será preenchido com SIM/NÃO/ERRO)
   - `Data consulta` (será preenchida com a data do VIES)
5. Clique em "Iniciar Validação"
6. Acompanhe o progresso na barra e no log
7. O Excel é atualizado automaticamente após cada VAT

## Estrutura do Excel

Exemplo de Excel válido:

| NIFS        | Válido no VIES | Data consulta |
|-------------|----------------|---------------|
| PT504772694 |                |               |
| ESA28017895 |                |               |
| FRXY123456  |                |               |

**Notas importantes:**
- A sheet deve chamar-se exatamente "NIF"
- Os cabeçalhos devem estar na linha 1
- Os dados começam na linha 2
- As colunas são identificadas pelo nome do cabeçalho, não pela posição
- Pode haver outras colunas além destas 3

## Tratamento de Dados

### Validação bem-sucedida
```
Válido no VIES: SIM
Data consulta: 2026-09-29+02:00
```

### Validação inválida
```
Válido no VIES: NÃO
Data consulta: 2026-09-29+02:00
```

Nota: Um VAT "NÃO" válido é diferente de um erro técnico.

### Erro técnico
```
Válido no VIES: ERRO
Data consulta: (vazio)
```

Exemplos de erros técnicos:
- Timeout do VIES (serviço indisponível)
- Erro de rede
- Resposta XML inválida
- SOAP Fault do servidor

## Segurança e Privacidade

- Nenhum dado é armazenado localmente além do que você escolher guardar no Excel
- Comunicação apenas com o serviço oficial VIES (https://ec.europa.eu/taxation_customs/vies/)
- Sem recolha de telemetria
- Sem acesso à rede para além das consultas VIES
- Sem modificação de ficheiros do sistema

## Tratamento de Erros

Se tiver problemas:

### "O ficheiro Excel não pode ser aberto"
- Feche o Excel se o tiver aberto
- Verifique se tem permissões de escrita na pasta
- Tente criar um novo Excel com os cabeçalhos necessários

### "A sheet 'NIF' não foi encontrada"
- Crie uma sheet chamada exatamente "NIF"
- Não use aspas nem espaços adicionais
- Verifique a capitalização (NIF com maiúsculas)

### "Colunas obrigatórias não encontradas"
- Verifique os nomes dos cabeçalhos:
  - `NIFS`
  - `Válido no VIES`
  - `Data consulta`
- Os nomes têm de ser exatos (capitalização, espaços, acentos)

### "Timeout do VIES"
- O serviço VIES pode estar indisponível ou lento
- Tente novamente mais tarde
- O Excel mantém os resultados já processados

## Reprocessamento

Se deseja revalidar VATs já processados:
1. Abra o Excel
2. Apague manualmente o conteúdo das colunas `Válido no VIES` e `Data consulta`
3. Guarde o Excel
4. Use o Validador VIES normalmente (só processa linhas vazias)

Ou use "Força Revalidação" (se implementada) para processar tudo novamente.

## Desenvolvimento

### Estrutura do projeto
```
C:\workspace\SapScript\Viex\
├── validador_vies_gui.py      (aplicação GUI principal)
├── vies_service.py             (lógica de validação VIES)
├── vies_excel.py               (leitura/escrita Excel)
├── requirements.txt            (dependências Python)
├── ValidadorVIES.spec          (configuração PyInstaller)
├── build.bat                   (script de build)
├── teste_viex.xlsx             (ficheiro de teste)
└── dist/
    └── ValidadorVIES/
        ├── ValidadorVIES.exe   (executável)
        └── _internal/          (dependências)
```

### Build do executável

```bash
cd C:\workspace\SapScript\Viex
build.bat
```

O executável será gerado em `dist\ValidadorVIES\ValidadorVIES.exe`

### Testes

```bash
python test_vies.py
```

## Limitações Conhecidas

- A aplicação funciona com um VAT de cada vez (sequencial) para respeitar o serviço VIES
- Alguns países (ex: Espanha) podem não disponibilizar Nome/Morada mesmo com VAT válido
- O serviço VIES pode estar indisponível em manutenção

## Suporte

Para problemas ou sugestões, contacte o administrador do sistema.

## Versão

Validador VIES v1.0 (2026)

## Compatibilidade

- Windows 7 ou superior
- Python 3.9+ (para desenvolvimento)
- Excel 2007+ (.xlsx)

## Aviso Legal

Este software utiliza o serviço oficial VIES da Comissão Europeia.
Para mais informações sobre o VIES: https://ec.europa.eu/taxation_customs/vies/

---

**Notas de Desenvolvimento:**
- PySimpleGUI: GUI leve e sem dependências web
- openpyxl: Manipulação segura de Excel
- requests: Cliente HTTP para SOAP/XML
- PyInstaller: Packaging para Windows
