
# WPO Academy - SPFx Básico

##

# SharePoint

## O que é

O SharePoint é uma plataforma web que disponibiliza sites com listas, bibliotecas e aplicações.

- As **listas** funcionam de forma semelhante a um banco de dados, mas com uma interface mais visual e voltada para o usuário final.
- As **bibliotecas** são utilizadas para armazenar e gerenciar arquivos com recursos avançados de controle e colaboração.
- A parte de **site** permite a navegação entre páginas, conteúdos e recursos da organização.

---

## Site Collections

Uma **Site Collection** é a estrutura raiz que agrupa todos os elementos do SharePoint: sites, subsites, listas, bibliotecas e páginas. É o "pai" de todo o conteúdo relacionado.

---

## Ambiente de Desenvolvimento SharePoint (SPFx)

📚 Documentação oficial da Microsoft:  
https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment

### Instale os seguintes itens:

- **Node.js** (versão recomendada: LTS)
- **npm** (gerenciador de pacotes do Node.js)
- **Yeoman e generator para SharePoint**:

```powershell
npm install -g yo @microsoft/generator-sharepoint
```

- **Gulp** (utilizado para automação de tarefas):

```powershell
npm install -g gulp
```

- **Visual Studio Code** (IDE recomendada)
- **Opcional**: Instale o [NVM](https://github.com/nvm-sh/nvm) para facilitar a gestão de múltiplas versões do Node.js

---

### Criação do tenant

Acesse:  
🔗 https://cdx.transform.microsoft.com/

---

### Configuração via PowerShell

1. Abra o **CMD ou PowerShell como administrador**
2. Instale o módulo do SharePoint Online:

```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
```

3. Conecte-se ao seu tenant:

```powershell
Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
```

4. Instale o módulo PnP PowerShell:

```powershell
Install-Module -Name "PnP.PowerShell" -Scope CurrentUser
```

---

## WebPart

Uma **WebPart** é um componente modular reutilizável que pode ser adicionado a páginas do SharePoint para exibir informações ou interagir com conteúdos de forma personalizada.

A ideia é utilizar WebParts principalmente para funcionalidades frequentes, como:

- Operações **CRUD**
- **Controle de permissões**
- **Consumo de APIs**
- Exibição de listas personalizadas

Para iniciar um novo projeto e criar sua primeira WebPart:

```powershell
yo @microsoft/sharepoint
```

> A **Solution** é o pacote que encapsula todas as WebParts e recursos do projeto.

---

## Estrutura do Projeto

- `config/package-solution.json`: configurações da solução, versionamento, recursos e metadados do projeto
- `config/serve.json`: define qual página será aberta com `gulp serve`
- `package-lock.json`: lista exata das dependências instaladas (não editar manualmente)
- `package.json`: define as dependências e scripts do projeto (pode ser editado)
- `node_modules/`: pasta com as dependências instaladas (não deve ser enviada ao repositório)
- `src/`: diretório principal do desenvolvimento
  - `src/webparts/`: onde ficam as WebParts
  - `src/webparts/components/`: onde estão os componentes e a estrutura da WebPart

**Importante**: O CSS da WebPart deve ser usado via `import styles from './MinhaWebPart.module.scss'`.  
A aplicação de estilos deve ser feita com:

```tsx
className={styles.containerLG}
```

> Isso porque, no momento da compilação, o nome da classe será transformado.

---

## Debug local (sem usar o workbench online)

Use a seguinte URL para testar localmente:

```
?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

---

## Verificar pacotes instalados globalmente

```bash
npm -g ls --depth=0
```

---

## Fast Serve (opcional)

Você pode usar o **SPFx Fast Serve** como alternativa ao `gulp serve`, para acelerar a compilação e recarregamento do projeto:

🔗 https://github.com/s-KaiNet/spfx-fast-serve



############################################################
# spfx

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
