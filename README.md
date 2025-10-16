# Servidor MCP para Microsoft Graph

Este projeto é um servidor **MCP (Model Context Protocol)** que atua como uma ponte entre um assistente de IA e a **API do Microsoft Graph**. Ele expõe um conjunto de ferramentas que permitem à IA interagir com serviços da Microsoft, como SharePoint, OneDrive e Excel, de forma programática e segura.

## 🚀 Começando: A Maneira Mais Fácil

A forma mais simples de usar o servidor é com `uvx`, que o executa em um ambiente isolado diretamente do GitHub.

### Pré-requisitos

- Você precisa de um **ID de Cliente (Client ID)** de um aplicativo registrado no Microsoft Azure. Se não tiver um, siga as instruções em **Apêndice A**.

### Execução

Execute o comando abaixo no seu terminal, substituindo `"seu-client-id-aqui"` pelo seu ID de Cliente.

```bash
uvx --from https://github.com/NagosC/microsoftMCP.git microsoft-mcp
```

- **Para Ambientes de IA (como o Gemini):**
  Você pode configurar a ferramenta para ser executada com o `CLIENT_ID` como uma variável de ambiente.

  ```json
  {
      "microsoft": {
          "command": "uvx",
          "args": [
              "--from",
              "https://github.com/NagosC/microsoftMCP.git",
              "microsoft-mcp"
          ],
          "env": {
              "MICROSOFT_MCP_CLIENT_ID": "seu-client-id-aqui"
          }
      }
  }
  ```

### Autenticação da Conta Microsoft

Após iniciar o servidor, você precisa autorizar o acesso à sua conta Microsoft.

1.  **Inicie a autenticação**:
    ```bash
    authenticate_account()
    ```
2.  **Código de Dispositivo**: O sistema fornecerá uma URL e um código.
    - Abra a URL no seu navegador.
    - Insira o código fornecido.
    - Faça login com sua conta da Microsoft e aprove o acesso.
3.  **Complete a autenticação**:
    ```bash
    complete_authentication(flow_cache="...")
    ```
    Use o `flow_cache` retornado pelo passo anterior.

Com a conta autenticada, você já pode usar todas as ferramentas disponíveis.

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
3.  **Crie um arquivo `.env`** na raiz do projeto e adicione seu Client ID:
    ```env
    # Cole o ID do Aplicativo (cliente) que você copiou do portal do Azure
    GRAPH_CLIENT_ID="seu-client-id-aqui"
    ```

### 3. Autenticação Local

Antes de iniciar o servidor, autentique sua conta Microsoft:

```bash
poetry run python src/autentichate.py
```

Siga as instruções no terminal para gerar o arquivo de token (`~/.microsoft_mcp_token_cache.json`).

### 4. Executando o Servidor

Inicie o servidor MCP em modo de desenvolvimento:

```bash
poetry run python -m src.microsoft_mcp.server
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