import base64
import datetime as dt
import json
import os
import pathlib as pl
from typing import Any

from fastmcp import FastMCP, auth as mcp_auth

from . import graph, auth

tenant_id = os.getenv("GRAPH_TENANT_ID", "common")

oauth_config = mcp_auth.OAuth2(
    authorization_url=f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize",
    token_url=f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
    scopes=auth.SCOPES,
)

mcp = FastMCP("microsoft-mcp", auth=oauth_config)


@mcp.tool
def set_client_id(client_id: str) -> dict[str, str]:
    """
    Sets and caches the Microsoft Application (Client) ID for this server.

    This ID will be stored locally and used for all subsequent authentication requests,
    removing the need to set the MICROSOFT_MCP_CLIENT_ID environment variable.

    Args:
        client_id: The Application (Client) ID from your Azure App Registration.
    """
    config_dir = auth.CONFIG_DIR
    config_dir.mkdir(parents=True, exist_ok=True)
    config_file = config_dir / "config.json"

    config = {"client_id": client_id}
    config_file.write_text(json.dumps(config, indent=2))

    # Attempt to get the app to validate the new client_id
    auth.get_app()

    return {"status": "success", "message": f"Client ID successfully set and cached at {config_file}"}

@mcp.tool
def list_accounts() -> list[dict[str, str]]:
    """List all signed-in Microsoft accounts"""
    return [
        {"username": acc.username, "account_id": acc.account_id}
        for acc in auth.list_accounts()
    ]


@mcp.tool
def authenticate_account() -> dict[str, str]:
    """Authenticate a new Microsoft account using device flow authentication

    Returns authentication instructions and device code for the user to complete authentication.
    The user must visit the URL and enter the code to authenticate their Microsoft account.
    """
    app = auth.get_app()
    flow = app.initiate_device_flow(scopes=auth.SCOPES)

    if "user_code" not in flow:
        error_msg = flow.get("error_description", "Unknown error")
        raise Exception(f"Failed to get device code: {error_msg}")

    verification_url = flow.get(
        "verification_uri",
        flow.get("verification_url", "https://microsoft.com/devicelogin"),
    )

    return {
        "status": "authentication_required",
        "instructions": "To authenticate a new Microsoft account:",
        "step1": f"Visit: {verification_url}",
        "step2": f"Enter code: {flow['user_code']}",
        "step3": "Sign in with the Microsoft account you want to add",
        "step4": "After authenticating, use the 'complete_authentication' tool to finish the process",
        "device_code": flow["user_code"],
        "verification_url": verification_url,
        "expires_in_seconds": flow.get("expires_in", 900),
        "_flow_cache": json.dumps(flow),
    }


@mcp.tool
def complete_authentication(flow_cache: str) -> dict[str, str]:
    """Complete the authentication process after the user has entered the device code

    Args:
        flow_cache: The flow data returned from authenticate_account (the _flow_cache field)

    Returns:
        Account information if authentication was successful
    """
    try:
        flow = json.loads(flow_cache)
    except (ValueError, SyntaxError):
        raise ValueError("Invalid flow cache data")

    app = auth.get_app()
    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        error_msg = result.get("error_description", result["error"])
        if "authorization_pending" in error_msg:
            return {
                "status": "pending",
                "message": "Authentication is still pending. The user needs to complete the authentication process.",
                "instructions": "Please ensure you've visited the URL and entered the code, then try again.",
            }
        raise Exception(f"Authentication failed: {error_msg}")

    # Save the token cache
    cache = app.token_cache
    if isinstance(cache, auth.msal.SerializableTokenCache) and cache.has_state_changed:
        auth._write_cache(cache.serialize())

    # Get the newly added account
    accounts = app.get_accounts()
    if accounts:
        # Find the account that matches the token we just got
        for account in accounts:
            if (
                account.get("username", "").lower()
                == result.get("id_token_claims", {})
                .get("preferred_username", "")
                .lower()
            ):
                return {
                    "status": "success",
                    "username": account["username"],
                    "account_id": account["home_account_id"],
                    "message": f"Successfully authenticated {account['username']}",
                }
        # If exact match not found, return the last account
        account = accounts[-1]
        return {
            "status": "success",
            "username": account["username"],
            "account_id": account["home_account_id"],
            "message": f"Successfully authenticated {account['username']}",
        }

    return {
        "status": "error",
        "message": "Authentication succeeded but no account was found",
    }


# --- SharePoint/Drive Tools ---


@mcp.tool
def sharepoint_get_site(
    hostname: str, relative_path: str, account_id: str | None = None
) -> dict[str, Any] | None:
    """Gets a SharePoint site by its hostname and relative server path.

    Args:
        hostname: The hostname of the SharePoint site (e.g., 'contoso.sharepoint.com').
        relative_path: The relative path to the site (e.g., '/sites/MySite').
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A dictionary containing the site's properties, or None if not found.
    """
    return graph.get_site(
        hostname=hostname, relative_path=relative_path, account_id=account_id
    )


@mcp.tool
def sharepoint_list_drives(
    site_id: str, account_id: str | None = None
) -> list[dict[str, Any]]:
    """Lists all Document Libraries (Drives) in a given SharePoint site.

    Args:
        site_id: The ID of the SharePoint site.
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A list of drives, where each drive is a dictionary of its properties.
    """
    return graph.get_drives(site_id=site_id, account_id=account_id)


@mcp.tool
def sharepoint_list_files(
    drive_id: str, item_id: str | None = None, account_id: str | None = None
) -> list[dict[str, Any]]:
    """Lists files and folders in a SharePoint drive or a specific folder.

    Args:
        drive_id: The ID of the drive.
        item_id: The ID of the folder to list items from. If None, lists from the root.
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A list of items, where each item is a dictionary of its properties.
    """
    items = graph.list_drive_items(
        drive_id=drive_id, item_id=item_id, account_id=account_id
    )
    return [
        {
            "id": item["id"],
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size"),
            "created_at": item.get("createdDateTime"),
            "last_modified_at": item.get("lastModifiedDateTime"),
        }
        for item in items
    ]


@mcp.tool
def sharepoint_download_file(
    drive_id: str, item_id: str, account_id: str | None = None
) -> str:
    """Downloads a file from a SharePoint drive and returns its content as a base64 encoded string.

    Args:
        drive_id: The ID of the drive containing the file.
        item_id: The ID of the file to download.
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        The content of the file, base64 encoded.
    """
    content = graph.download_file(
        drive_id=drive_id, item_id=item_id, account_id=account_id
    )
    return base64.b64encode(content).decode("utf-8") if content else ""


@mcp.tool
def sharepoint_upload_file(
    drive_id: str, parent_id: str, filename: str, content_b64: str, account_id: str | None = None
) -> dict[str, Any] | None:
    """Uploads a file to a SharePoint drive. The file content must be base64 encoded.
    For small files only (< 4MB).
    """
    data = base64.b64decode(content_b64)
    return graph.upload_small_file(drive_id, parent_id, filename, data, account_id)


# --- Excel Tools ---


@mcp.tool
def excel_list_worksheets(
    drive_id: str, item_id: str, account_id: str | None = None
) -> list[dict[str, Any]]:
    """Lists all worksheets in an Excel file located on SharePoint/OneDrive.

    Args:
        drive_id: The ID of the drive containing the Excel file.
        item_id: The ID of the Excel file item.
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A list of worksheets, each with its name and visibility status.
    """
    worksheets = graph.get_excel_worksheets(
        drive_id=drive_id, item_id=item_id, account_id=account_id
    )
    return [
        {"name": ws["name"], "visibility": ws["visibility"]} for ws in worksheets
    ]


@mcp.tool
def excel_read_range(
    drive_id: str,
    item_id: str,
    worksheet_name: str,
    range_address: str,
    account_id: str | None = None,
) -> dict[str, Any] | None:
    """Reads data from a specified range in an Excel worksheet (e.g., 'A1:C3' or 'MyTable').

    Args:
        drive_id: The ID of the drive containing the Excel file.
        item_id: The ID of the Excel file item.
        worksheet_name: The name of the worksheet.
        range_address: The range to read from (e.g., 'A1:C3', 'Sheet1!A1:C3', or a table name).
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A dictionary containing the values, text, and formulas from the specified range.
    """
    return graph.get_excel_range(
        drive_id, item_id, worksheet_name, range_address, account_id
    )


@mcp.tool
def excel_update_range(
    drive_id: str, item_id: str, worksheet_name: str, range_address: str, values: list[list[Any]], account_id: str | None = None
) -> dict[str, Any] | None:
    """Updates a range in an Excel worksheet with the given values. The 'values' should be a 2D list."""
    return graph.update_excel_range(drive_id, item_id, worksheet_name, range_address, values, account_id)


@mcp.tool
def excel_list_tables(
    drive_id: str, item_id: str, worksheet_name: str, account_id: str | None = None
) -> list[dict[str, Any]]:
    """Lists all tables in a specific Excel worksheet.

    Args:
        drive_id: The ID of the drive containing the Excel file.
        item_id: The ID of the Excel file item.
        worksheet_name: The name of the worksheet to list tables from.
        account_id: The ID of the account to use. Uses the default account if not provided.

    Returns:
        A list of tables, where each table is a dictionary of its properties.
    """
    return graph.get_excel_tables(drive_id, item_id, worksheet_name, account_id)


@mcp.tool
def excel_add_table_row(
    drive_id: str,
    item_id: str,
    worksheet_name: str,
    table_name: str,
    values: list[list[Any]],
    account_id: str | None = None,
) -> dict[str, Any] | None:
    """Adds one or more rows to the end of a specified table in an Excel worksheet.

    The 'values' argument must be a 2D list (a list of lists), where each inner list represents a row to be added.
    """
    return graph.add_excel_table_row(drive_id, item_id, worksheet_name, table_name, values, account_id)
