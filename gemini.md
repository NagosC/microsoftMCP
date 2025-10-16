
````
# GEMINI.md

Este arquivo fornece orientação para assistentes de IA (como Gemini) ao trabalhar com o código neste repositório.

## Visão Geral do Projeto

Este é um **Servidor MCP (Model Context Protocol) em Python** que permite que assistentes de IA interajam com arquivos e listas do **Microsoft SharePoint** usando a **Microsoft Graph API**. Ele serve como uma ponte para manipulação programática de dados e documentos no SharePoint.

## Comandos de Desenvolvimento (Python/fastMCP)

Assumindo que você usa `Poetry` para gerenciamento de dependências, ou um ambiente virtual padrão.

### Configuração e Desenvolvimento
```bash
poetry install        # Instala as dependências do `pyproject.toml`
python main.py        # Inicia o servidor fastMCP em modo de desenvolvimento
````

### Testes

Bash

```
pytest                # Roda todos os testes
pytest tests/tools/test_sharepoint.py  # Roda testes de um arquivo específico
pytest -k "test_read_list_data"        # Roda um teste específico
```

### Linting e Formatação

Bash

```
black .               # Formata o código Python (estilo Black)
isort .               # Organiza imports
mypy .                # Checagem de tipos estática
```

## Arquitetura

### Componentes Principais

**Backend:** O servidor utiliza o framework **fastMCP** (baseado em FastAPI) para hospedar os _tools_ do MCP. A lógica de negócios primária para a interação com o SharePoint reside em um módulo de serviço dedicado.

**Camada de Serviço do Graph:** Um módulo Python que encapsula a lógica de autenticação e todas as chamadas HTTP para a **Microsoft Graph API**.

**Interfaces Chave:**

- **`GraphInterface`** (ex: `internal/sharepoint/graph_service.py`) - API unificada para operações no SharePoint/Graph.
    
- **`Tool` Interface** (implementada em `internal/tools/`) - Implementações dos _tools_ do MCP usando a `GraphInterface`.
    

**Ponto de Entrada:**

- `main.py` - Ponto de entrada principal do servidor, onde o `fastMCP` é inicializado e os _tools_ são registrados.
    

### Sistema de Tools (Ferramentas)

As ferramentas (tools) do MCP são implementadas em `internal/tools/` para interagir com o SharePoint:

- `sharepoint_get_site_info` - Obtém metadados de um site/drive específico.
    
- `sharepoint_list_files` - Lista arquivos em um drive ou pasta do SharePoint.
    
- `sharepoint_read_document` - Lê o conteúdo de um documento (ex: TXT, JSON, ou usar a API de visualização do Graph).
    
- `sharepoint_read_list_data` - Lê itens de uma Lista do SharePoint (equivalente a uma tabela).
    
- `sharepoint_search` - Realiza buscas em documentos e listas do SharePoint.
    

### Autenticação e Configuração

A autenticação é feita via **OAuth 2.0 (Client Credentials ou On-Behalf-Of)** para acessar a Microsoft Graph API.

- As credenciais (Client ID, Tenant ID, Client Secret) **NUNCA** devem ser _hard-coded_ e devem ser lidas de **variáveis de ambiente** ou de um serviço de _secrets_.
    

## Estrutura de Arquivos

```
main.py                   # Ponto de entrada do servidor fastMCP
pyproject.toml            # Gerenciamento de dependências (Poetry)
internal/
  sharepoint/
    graph_service.py      # Camada de abstração do Microsoft Graph
    auth_manager.py       # Lógica de autenticação e tokens
  server/
    server_setup.py       # Configuração e inicialização do fastMCP
  tools/
    __init__.py
    sharepoint_list_tools.py # Implementação dos tools de Listas
    sharepoint_file_tools.py # Implementação dos tools de Arquivos
tests/                    # Testes de unidade e integração
```

## Configuração

Variáveis de ambiente (necessárias para o Graph API):

|Variável|Descrição|
|---|---|
|`GRAPH_TENANT_ID`|ID do seu Tenant (Azure AD)|
|`GRAPH_CLIENT_ID`|ID do Aplicativo (Cliente) registrado no Azure AD|
|`GRAPH_CLIENT_SECRET`|Segredo do Cliente (Chave do Aplicativo)|
|`GRAPH_SITE_ID`|ID do site do SharePoint alvo (para acesso direto)|
|`MCP_HOST`|Host onde o servidor MCP irá rodar (padrão: `127.0.0.1`)|
|`MCP_PORT`|Porta onde o servidor MCP irá rodar (padrão: `8000`)|

## Dependências

**Python:** Requer Python 3.10 ou superior.

**Pacotes Chave:**

- **`fastmcp`**: Framework para construir o servidor MCP.
    
- **`httpx` ou `requests`**: Cliente HTTP para interagir com o Microsoft Graph.
    
- **`azure-identity`**: Para gerenciamento de tokens de autenticação Azure/Graph.
    
- **`pydantic`**: Para validação e serialização de dados (usado pelo `fastMCP`).