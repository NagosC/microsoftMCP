# Servidor MCP para Microsoft Graph

Este projeto é um servidor **MCP (Model Context Protocol)** que atua como uma ponte entre um assistente de IA e a **API do Microsoft Graph**. Ele expõe um conjunto de ferramentas que permitem à IA interagir com serviços da Microsoft, como SharePoint, OneDrive e Excel, de forma programática e segura.

## 🚀 Começando: A Maneira Mais Fácil (com Docker)

A forma mais simples e recomendada de executar o servidor é com Docker e Docker Compose.

### 1. Pré-requisitos

- **Docker** e **Docker Compose** instalados.
- Um **ID de Cliente (Client ID)** de um aplicativo registrado no Microsoft Azure. Se não tiver um, siga as instruções em **Apêndice A**.

### 2. Configuração

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/NagosC/microsoftMCP.git
    cd microsoftMCP
    ```
2.  **Crie um arquivo `.env`**:
    Copie o arquivo de exemplo `.env.example` para um novo arquivo chamado `.env`.
    ```bash
    cp .env.example .env
    ```
3.  **Preencha o arquivo `.env`**:
    Abra o arquivo `.env` e adicione seu `GRAPH_CLIENT_ID` e outras credenciais que desejar.

### 3. Executando o Servidor

Inicie o servidor com um único comando:

```bash
docker compose up --build -d
```

O servidor estará rodando em segundo plano e acessível em `http://localhost:8000`.

### 4. Autenticação da Conta Microsoft

Para autenticar uma nova conta Microsoft, execute o script de autenticação dentro do container:

```bash
docker compose exec server poetry run microsoft-mcp-auth
```

Siga as instruções no terminal para abrir a URL, inserir o código e autorizar o acesso.

---

## 🧑‍💻 Para Desenvolvedores: Configuração Local

Se você deseja modificar ou contribuir com o projeto, siga estes passos.

### 1. Pré-requisitos

- **Python 3.10+**
- **Poetry** (gerenciador de dependências)
- Um **ID de Cliente (Client ID)** do Azure (veja o **Apêndice A**).

### 2. Instalação

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/NagosC/microsoftMCP.git
    cd microsoftMCP
    ```
2.  **Instale as dependências:**
    ```bash
    poetry install
    ```
3.  **Crie um arquivo `.env`** na raiz do projeto (pode copiar do `.env.example`) e adicione seu Client ID:
    ```env
    # Cole o ID do Aplicativo (cliente) que você copiou do portal do Azure
    GRAPH_CLIENT_ID="seu-client-id-aqui"
    ```

### 3. Autenticação Local

Antes de iniciar o servidor, autentique sua conta Microsoft:

```bash
poetry run microsoft-mcp-auth
```

Siga as instruções no terminal para gerar o arquivo de token (`~/.microsoft_mcp_token_cache.json`).

### 4. Executando o Servidor

Inicie o servidor MCP em modo de desenvolvimento:

```bash
poetry run microsoft-mcp
```

### 5. Comandos Úteis

- **Formatar código**: `poetry run black . && poetry run isort .`
- **Checagem de tipos**: `poetry run mypy .`
- **Rodar testes**: `poetry run pytest`

---

## 🛠️ Ferramentas Disponíveis (API)

Aqui está a lista de ferramentas que o servidor expõe.

*Nota: Todas as ferramentas aceitam um `account_id` opcional. Se omitido, a primeira conta autenticada será usada.*

| Ferramenta | Descrição |
| --- | --- |
| **`list_accounts()`** | Lista todas as contas da Microsoft autenticadas. |
| **`authenticate_account()`** | Inicia um novo fluxo de autenticação de dispositivo. |
| **`complete_authentication(flow_cache)`** | Finaliza o processo de autenticação. |
| **`sharepoint_get_site(hostname, relative_path)`** | Obtém detalhes de um site do SharePoint. |
| **`sharepoint_list_drives(site_id)`** | Lista as bibliotecas de documentos (Drives) de um site. |
| **`sharepoint_list_files(drive_id, item_id)`** | Lista arquivos e pastas em um Drive ou pasta. |
| **`sharepoint_download_file(drive_id, item_id)`** | Baixa o conteúdo de um arquivo (retorna em base64). |
| **`sharepoint_upload_file(drive_id, parent_id, filename, content_b64)`** | Faz upload de um arquivo pequeno (< 4MB). |
| **`excel_list_worksheets(drive_id, item_id)`** | Lista todas as planilhas em um arquivo Excel. |
| **`excel_list_tables(drive_id, item_id, worksheet_name)`** | Lista todas as tabelas formatadas em uma planilha. |
| **`excel_read_range(drive_id, item_id, worksheet_name, range_address)`** | Lê dados de um intervalo (ex: "A1:C5"). |
| **`excel_update_range(drive_id, item_id, worksheet_name, range_address, values)`** | Atualiza um intervalo de células. |
| **`excel_add_table_row(drive_id, item_id, worksheet_name, table_name, values)`** | Adiciona uma ou mais linhas a uma tabela. |

---

## Apêndice A: Registro de Aplicativo no Azure AD

Para obter um **ID de Cliente (Client ID)**, você precisa registrar um aplicativo no Azure Active Directory.

1.  Acesse o **portal do Azure** e navegue até **Azure Active Directory**.
2.  Vá para **Registros de aplicativo** e clique em **Novo registro**.
3.  Dê um nome ao seu aplicativo (ex: `My-MCP-Server`).
4.  Em **Tipos de conta com suporte**, selecione **Contas em qualquer diretório organizacional... e contas pessoais da Microsoft...**.
5.  Vá para a guia **Autenticação**, clique em **Adicionar uma plataforma** e selecione **Aplicativos móveis e de desktop**.
6.  Marque a caixa de seleção para `https://login.microsoftonline.com/common/oauth2/nativeclient`.
7.  Ainda em **Autenticação**, ative a opção **Permitir fluxos de cliente público**.
8.  Vá para **Permissões de API** e adicione as seguintes permissões **delegadas**:
    - `Files.ReadWrite.All`
    - `Sites.ReadWrite.All`
    - `User.Read`
    - `offline_access`
9.  Na página de **Visão geral** do seu aplicativo, copie o **ID do aplicativo (cliente)**. Este é o seu `GRAPH_CLIENT_ID`.