 # Microsoft Graph MCP Server

Este projeto é um servidor **MCP (Model Context Protocol)** que atua como uma ponte entre um assistente de IA e a **API do Microsoft Graph**. Ele expõe um conjunto de ferramentas que permitem à IA interagir com serviços da Microsoft, como SharePoint, OneDrive e Excel, de forma programática e segura.

## Funcionalidades

- **Autenticação Segura**: Utiliza o fluxo de autenticação de dispositivo OAuth 2.0 para conectar contas da Microsoft de forma segura, armazenando tokens localmente.
- **Manipulação de Arquivos**: Lista, baixa e faz upload de arquivos no SharePoint e OneDrive.
- **Interação com SharePoint**: Obtém informações sobre sites e bibliotecas de documentos (Drives).
- **Manipulação de Planilhas Excel**: Lista planilhas, lê e escreve em intervalos de células e adiciona linhas a tabelas formatadas.

---

## 📋 Pré-requisitos

1.  **Python 3.10+**
2.  **Poetry** para gerenciamento de dependências.
3.  **Registro de Aplicativo no Azure**: Você precisa de um aplicativo registrado no Microsoft Azure para obter as credenciais da API.

---

## ⚙️ Configuração e Instalação

### 1. Registro do Aplicativo no Azure AD

Para interagir com a API do Microsoft Graph, você precisa registrar um aplicativo no Azure Active Directory.

1.  Acesse o portal do Azure e navegue até **Azure Active Directory**.
2.  Vá para **Registros de aplicativo** e clique em **Novo registro**.
3.  Dê um nome ao seu aplicativo (ex: `My-MCP-Server`).
4.  Em **Tipos de conta com suporte**, selecione **Contas em qualquer diretório organizacional (Qualquer diretório do Azure AD – Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**.
5.  Vá para a guia **Autenticação**, clique em **Adicionar uma plataforma** e selecione **Aplicativos móveis e de desktop**.
6.  Marque a caixa para `https://login.microsoftonline.com/common/oauth2/nativeclient`.
7.  Ainda em **Autenticação**, role para baixo até **Configurações avançadas** e ative a opção **Permitir fluxos de cliente público**.
8.  Vá para **Permissões de API** e adicione as seguintes permissões delegadas:
    - `Files.ReadWrite.All`
    - `Sites.ReadWrite.All`
    - `User.Read`
    - `offline_access`
9.  Na página de **Visão geral** do seu aplicativo, copie o **ID do aplicativo (cliente)**. Este será o seu `GRAPH_CLIENT_ID`.

### 2. Configuração do Ambiente Local

Clone o repositório e configure o ambiente.

```bash
# 1. Clone o repositório
git clone <URL_DO_SEU_REPOSITORIO>
cd projetoCmp

# 2. Crie o arquivo .env
# Copie o conteúdo abaixo para um novo arquivo chamado .env na raiz do projeto
```

**`.env`**
```env
# Cole o ID do Aplicativo (cliente) que você copiou do portal do Azure
GRAPH_CLIENT_ID="seu-client-id-aqui"

# (Opcional) Se você estiver usando um tenant específico, descomente e preencha.
# Caso contrário, o padrão será 'common' (multilocatário).
# GRAPH_TENANT_ID="seu-tenant-id-aqui"
```

### 3. Instalação das Dependências

Use o Poetry para instalar todas as dependências listadas no `pyproject.toml`.

```bash
poetry install
```

---

## 🔑 Autenticação

Antes de iniciar o servidor, você precisa autenticar as contas da Microsoft que deseja usar. Execute o script de autenticação interativo:

```bash
poetry run python src/autentichate.py
```

Siga as instruções no terminal:
1.  O script perguntará se você deseja autenticar uma nova conta. Digite `y`.
2.  Ele fornecerá uma URL e um código de dispositivo.
3.  Abra a URL em um navegador, insira o código e faça login com a conta da Microsoft desejada.
4.  Após a autenticação bem-sucedida, o script salvará o token de acesso no arquivo `~/.microsoft_mcp_token_cache.json`.

Você pode repetir o processo para adicionar várias contas.

---

## 🚀 Executando o Servidor

Com as dependências instaladas e pelo menos uma conta autenticada, inicie o servidor MCP:

```bash
poetry run python -m src.microsoft_mcp.server
```

O servidor estará em execução e pronto para receber chamadas de um cliente MCP compatível.

---

## ⚡ Execução Direta do GitHub (Avançado)

Se você deseja executar o servidor sem clonar o repositório, pode usar uma ferramenta como o `pipx`. Isso é ideal para integrar o servidor a outros sistemas de forma rápida.

### 1. Pré-requisito: Instalar o `pipx`

Se você ainda não tem o `pipx`, instale-o com o pip:

```bash
pip install pipx
```

### 2. Executando o Servidor via `pipx`

Use o comando abaixo para instalar e executar o servidor diretamente do repositório do GitHub. Certifique-se de definir a variável de ambiente `MICROSOFT_MCP_CLIENT_ID` na mesma linha.

```bash
# Substitua <URL_DO_SEU_REPOSITORIO> pela URL do seu repo no GitHub
MICROSOFT_MCP_CLIENT_ID="seu-client-id-aqui" pipx run --spec git+<URL_DO_SEU_REPOSITORIO> microsoft-mcp
```

Isso fará com que o `pipx` baixe o código, instale as dependências em um ambiente virtual isolado e execute o ponto de entrada `microsoft-mcp` que definimos no `pyproject.toml`.
---

## 🛠️ Ferramentas Disponíveis (API)

Aqui está a lista de ferramentas que o servidor expõe.

### Autenticação

- **`list_accounts()`**: Lista todas as contas da Microsoft que já foram autenticadas.
- **`authenticate_account()`**: Inicia um novo fluxo de autenticação de dispositivo.
- **`complete_authentication(flow_cache: str)`**: Finaliza o processo de autenticação após o usuário inserir o código no navegador.

### SharePoint e OneDrive

- **`sharepoint_get_site(hostname: str, relative_path: str)`**: Obtém detalhes de um site do SharePoint.
- **`sharepoint_list_drives(site_id: str)`**: Lista as bibliotecas de documentos (Drives) de um site.
- **`sharepoint_list_files(drive_id: str, item_id: str | None = None)`**: Lista arquivos e pastas em um Drive ou pasta específica.
- **`sharepoint_download_file(drive_id: str, item_id: str)`**: Baixa o conteúdo de um arquivo (retorna em base64).
- **`sharepoint_upload_file(drive_id: str, parent_id: str, filename: str, content_b64: str)`**: Faz upload de um arquivo pequeno (< 4MB).

### Excel

- **`excel_list_worksheets(drive_id: str, item_id: str)`**: Lista todas as planilhas em um arquivo Excel.
- **`excel_list_tables(drive_id: str, item_id: str, worksheet_name: str)`**: Lista todas as tabelas formatadas em uma planilha.
- **`excel_read_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str)`**: Lê dados de um intervalo (ex: "A1:C5" ou "NomeDaTabela").
- **`excel_update_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str, values: list[list])`**: Atualiza um intervalo de células com novos valores.
- **`excel_add_table_row(drive_id: str, item_id: str, worksheet_name: str, table_name: str, values: list[list])`**: Adiciona uma ou mais linhas ao final de uma tabela.

> **Nota**: Todas as ferramentas aceitam um parâmetro opcional `account_id: str` para especificar qual conta autenticada usar. Se não for fornecido, a primeira conta da lista será usada como padrão.

---

## 🧑‍💻 Desenvolvimento

Para manter a qualidade do código, utilize as seguintes ferramentas:

```bash
# Formatar o código
poetry run black .
poetry run isort .

# Checagem de tipos estática
poetry run mypy .

# Rodar testes
poetry run pytest
```
