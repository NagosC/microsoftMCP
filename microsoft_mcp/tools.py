import base64
import datetime as dt
import json
import os
import pathlib as pl
from typing import Any

from fastmcp import FastMCP
from . import graph, auth

mcp = FastMCP("microsoft-mcp")

FOLDERS = {
    k.casefold(): v
    for k, v in {
        "inbox": "inbox",
        "sent": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "junk": "junkemail",
        "archive": "archive",
    }.items()
}


@mcp.tool
def list_accounts() -> list[dict[str, str]]:
    """List all signed-in Microsoft accounts"""
    return [
        {"username": acc.username, "account_id": acc.account_id}
        for acc in auth.list_accounts()
    ]


@mcp.tool
def authenticate_account() -> dict[str, str]:
    """Authenticate a new Microsoft account using device flow.

    This tool initiates a device flow authentication process. It will print a URL
    and a code to the console. The user must open the URL in a browser, enter
    the code, and sign in to their Microsoft account. The tool will wait for
    the authentication to complete and then return the account information.
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

    print(f"To authenticate, visit: {verification_url} and enter code: {flow['user_code']}")

    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        error_msg = result.get("error_description", result["error"])
        raise Exception(f"Authentication failed: {error_msg}")

    cache = app.token_cache
    if isinstance(cache, auth.msal.SerializableTokenCache) and cache.has_state_changed:
        auth._write_cache(cache.serialize())

    accounts = app.get_accounts()
    if accounts:
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


@mcp.tool
async def sharepoint_get_site(
    hostname: str, relative_path: str, account_id: str | None = None, timeout: float = 30.0
) -> dict[str, Any] | None:
    return await graph.get_site(
        hostname=hostname, relative_path=relative_path, account_id=account_id, timeout=timeout
    )


@mcp.tool
async def sharepoint_get_site_by_url(
    url: str | None = None, account_id: str | None = None, timeout: float = 30.0
) -> dict[str, Any] | None:
    from urllib.parse import urlparse

    if not url:
        url = os.getenv("SHAREPOINT_SITE_URL")

    if not url:
        raise ValueError("SharePoint URL not provided. Pass it as an argument or set the SHAREPOINT_SITE_URL environment variable.")

    parsed_url = urlparse(url)
    hostname = parsed_url.hostname
    relative_path = parsed_url.path

    if not hostname or not relative_path:
        raise ValueError(f"Invalid SharePoint URL provided: {url}. Could not parse hostname or path.")

    return await graph.get_site(
        hostname=hostname, relative_path=relative_path, account_id=account_id, timeout=timeout
    )


@mcp.tool
async def sharepoint_list_drives(
    site_id: str, account_id: str | None = None, timeout: float = 30.0
) -> list[dict[str, Any]]:
    return await graph.get_drives(site_id=site_id, account_id=account_id, timeout=timeout)


@mcp.tool
async def sharepoint_list_files(
    drive_id: str, item_id: str | None = None, account_id: str | None = None, timeout: float = 30.0
) -> list[dict[str, Any]]:
    items = await graph.list_drive_items(
        drive_id=drive_id, item_id=item_id, account_id=account_id, timeout=timeout
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
async def sharepoint_download_file(
    drive_id: str, item_id: str, account_id: str | None = None, timeout: float = 60.0
) -> str:
    content = await graph.download_file(
        drive_id=drive_id, item_id=item_id, account_id=account_id, timeout=timeout
    )
    return base64.b64encode(content).decode("utf-8") if content else ""


@mcp.tool
async def sharepoint_upload_file(
    drive_id: str, parent_id: str, filename: str, content_b64: str, account_id: str | None = None, timeout: float = 30.0
) -> dict[str, Any] | None:
    data = base64.b64decode(content_b64)
    return await graph.upload_small_file(drive_id, parent_id, filename, data, account_id, timeout=timeout)


@mcp.tool
async def excel_list_worksheets(
    drive_id: str, item_id: str, account_id: str | None = None, timeout: float = 30.0
) -> list[dict[str, Any]]:
    worksheets = await graph.get_excel_worksheets(
        drive_id=drive_id, item_id=item_id, account_id=account_id, timeout=timeout
    )
    return [
        {"name": ws["name"], "visibility": ws["visibility"]} for ws in worksheets
    ]


@mcp.tool
async def excel_read_range(
    drive_id: str,
    item_id: str,
    worksheet_name: str,
    range_address: str,
    account_id: str | None = None,
    timeout: float = 30.0,
) -> dict[str, Any] | None:
    return await graph.get_excel_range(
        drive_id, item_id, worksheet_name, range_address, account_id, timeout=timeout
    )


@mcp.tool
async def excel_update_range(
    drive_id: str, item_id: str, worksheet_name: str, range_address: str, values: list[list[Any]], account_id: str | None = None, timeout: float = 30.0
) -> dict[str, Any] | None:
    return await graph.update_excel_range(drive_id, item_id, worksheet_name, range_address, values, account_id, timeout=timeout)


@mcp.tool
async def excel_list_tables(
    drive_id: str, item_id: str, worksheet_name: str, account_id: str | None = None, timeout: float = 30.0
) -> list[dict[str, Any]]:
    return await graph.get_excel_tables(drive_id, item_id, worksheet_name, account_id, timeout=timeout)


@mcp.tool
async def excel_add_table_row(
    drive_id: str,
    item_id: str,
    worksheet_name: str,
    table_name: str,
    values: list[list[Any]],
    account_id: str | None = None,
    timeout: float = 30.0,
) -> dict[str, Any] | None:
    return await graph.add_excel_table_row(drive_id, item_id, worksheet_name, table_name, values, account_id, timeout=timeout)