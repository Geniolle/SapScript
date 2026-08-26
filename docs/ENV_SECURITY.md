# Segurança do ficheiro `.env`

O projeto mantém os valores reais apenas no `.env` local. O Git recebe somente o modelo sem segredos (`.env.example`) e o backup encriptado (`.env.enc`). A chave privada age nunca deve ser guardada no repositório.

## Configuração inicial no Windows

1. Instale as ferramentas em PowerShell:

   ```powershell
   winget install --exact --id Mozilla.SOPS
   winget install --exact --id FiloSottile.age
   ```

2. Feche e abra o PowerShell para atualizar o `PATH`.
3. Crie a pasta padrão e uma chave age, somente se ainda não possuir a chave usada pelo projeto:

   ```powershell
   New-Item -ItemType Directory -Force "$env:APPDATA\sops\age"
   age-keygen -o "$env:APPDATA\sops\age\keys.txt"
   ```

4. Guarde uma cópia da chave privada num gestor de passwords ou cofre corporativo aprovado. Sem essa chave, o `.env.enc` não poderá ser recuperado após uma formatação.
5. Copie `.env.example` para `.env`, preencha apenas o ficheiro local e gere o backup:

   ```powershell
   Copy-Item .env.example .env
   .\scripts\backup-env.ps1
   ```

6. Reveja e versione `.env.example`, `.env.enc`, `.sops.yaml`, os scripts e esta documentação. Nunca use `git add -f .env`.

## Depois de formatar a máquina

```powershell
git clone https://github.com/Geniolle/SapScript.git
cd SapScript

winget install --exact --id Mozilla.SOPS
winget install --exact --id FiloSottile.age
```

Restaure a chave privada original em `%APPDATA%\sops\age\keys.txt` a partir do cofre seguro. Não crie uma chave nova: uma chave diferente não consegue abrir o backup existente. Depois execute:

```powershell
.\scripts\restore-env.ps1
```

O fluxo é:

```text
.env.enc -> restore-env.ps1 -> .env
```

Se um `.env` local já existir, o script recusa substituí-lo. Confira o ficheiro e, somente quando a substituição for intencional, execute `restore-env.ps1 -Force`.

## O que pode ir para o GitHub

| Ficheiro | GitHub | Conteúdo |
|---|---:|---|
| `.env` | Nunca | Valores reais e segredos locais |
| `.env.example` | Sim | Nomes, comentários e defaults não sensíveis |
| `.env.enc` | Sim | Conteúdo cifrado pelo SOPS + age |
| `.sops.yaml` | Sim | Regra e chave pública age |
| `%APPDATA%\sops\age\keys.txt` | Nunca | Chave privada age |

## Operações habituais

Após alterar qualquer valor local:

```powershell
.\scripts\backup-env.ps1
git status
```

O script não imprime o conteúdo do `.env`. A escrita é feita primeiro num ficheiro temporário ignorado e só substitui `.env.enc` depois de o SOPS terminar com sucesso.

Para usar uma chave guardada noutro local seguro fora do projeto:

```powershell
$env:SOPS_AGE_KEY_FILE = "D:\cofre-seguro\keys.txt"
.\scripts\restore-env.ps1
```

## Incidente ou perda de chave

- Se um `.env` em texto aberto tiver sido publicado, remova-o do índice e rode todas as credenciais contidas nele. Apagar apenas o ficheiro no commit atual não elimina cópias históricas.
- Não reescreva histórico partilhado sem coordenar com os colaboradores.
- Se a chave privada age for perdida, o `.env.enc` não é recuperável. Restaure a chave a partir do cofre seguro.
