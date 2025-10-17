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
        "expires_in": str(flow.get("expires_in", 900)),
        "_flow_cache": str(flow),
    }


@mcp.tool
def complete_authentication(flow_cache: str) -> dict[str, str]:
    """Complete the authentication process after the user has entered the device code

    Args:
        flow_cache: The flow data returned from authenticate_account (the _flow_cache field)

    Returns:
        Account information if authentication was successful
    """
    import ast

    try:
        flow = ast.literal_eval(flow_cache)
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