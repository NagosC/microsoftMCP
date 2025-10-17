import httpx
import asyncio
import logging
from typing import Any
from .auth import get_token

BASE_URL = "https://graph.microsoft.com/v1.0"
UPLOAD_CHUNK_SIZE = 15 * 320 * 1024

logger = logging.getLogger(__name__)


async def request(
    method: str,
    path: str,
    account_id: str | None = None,
    params: dict[str, Any] | None = None,
    json: dict[str, Any] | None = None,
    data: bytes | None = None,
    timeout: float = 30.0,
    max_retries: int = 3,
) -> dict[str, Any] | None:
    """
    Makes a request to the Microsoft Graph API with authentication and retry logic.
    """
    async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
        headers = {
            "Authorization": f"Bearer {get_token(account_id)}",
        }

        if method.upper() == "GET":
            current_params = params or {}
            if "$search" in current_params or "body" in current_params.get("$select", ""):
                headers["Prefer"] = 'outlook.body-content-type="text"'
        else:
            headers["Content-Type"] = (
                "application/json" if json is not None else "application/octet-stream"
            )

        if params:
            filter_str = params.get("$filter", "")
            if "$search" in params or "contains(" in filter_str or "/any(" in filter_str:
                headers["ConsistencyLevel"] = "eventual"
                params.setdefault("$count", "true")

        for attempt in range(max_retries + 1):
            try:
                response = await client.request(
                    method=method,
                    url=f"{BASE_URL}{path}",
                    headers=headers,
                    params=params,
                    json=json,
                    content=data,
                )

                if response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", "5"))
                    if attempt < max_retries:
                        logger.warning(
                            f"Rate limited. Retrying after {retry_after} seconds."
                        )
                        await asyncio.sleep(min(retry_after, 60))
                        continue

                response.raise_for_status()

                if response.status_code == 204 or not response.content:
                    return None
                return response.json()

            except (httpx.HTTPStatusError, httpx.TransportError) as e:
                is_server_error = (
                    isinstance(e, httpx.HTTPStatusError) and e.response.status_code >= 500
                )
                is_transport_error = isinstance(e, httpx.TransportError)

                if (is_server_error or is_transport_error) and attempt < max_retries:
                    wait_time = (2**attempt) * 1
                    logger.warning(
                        f"Request failed (attempt {attempt + 1}/{max_retries + 1}). Retrying in {wait_time}s. Error: {e}"
                    )
                    await asyncio.sleep(wait_time)
                    continue
                raise

    return None


async def _paginated_request(path: str, account_id: str | None, timeout: float = 30.0) -> list[dict[str, Any]]:
    """
    Handles paginated requests to the Graph API.
    """
    results = []
    next_url = path

    while next_url:
        path_to_request = next_url if next_url.startswith(BASE_URL) else f"{BASE_URL}{next_url}"
        
        response = await request(
            "GET",
            path_to_request.replace(BASE_URL, ""),
            account_id=account_id,
            timeout=timeout
        )

        if response and "value" in response:
            results.extend(response["value"])
            next_url = response.get("@odata.nextLink")
        else:
            break

    return results


async def get_site(hostname: str, relative_path: str, account_id: str | None = None, timeout: float = 30.0) -> dict[str, Any] | None:
    path = f"/sites/{hostname}:/{relative_path}"
    return await request("GET", path, account_id=account_id, timeout=timeout)


async def get_drives(site_id: str, account_id: str | None = None, timeout: float = 30.0) -> list[dict[str, Any]]:
    path = f"/sites/{site_id}/drives"
    return await _paginated_request(path, account_id, timeout=timeout)


async def list_drive_items(drive_id: str, item_id: str | None = None, account_id: str | None = None, timeout: float = 30.0) -> list[dict[str, Any]]:
    if item_id:
        path = f"/drives/{drive_id}/items/{item_id}/children"
    else:
        path = f"/drives/{drive_id}/root/children"
    return await _paginated_request(path, account_id, timeout=timeout)


async def get_excel_worksheets(drive_id: str, item_id: str, account_id: str | None = None, timeout: float = 30.0) -> list[dict[str, Any]]:
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets"
    return await _paginated_request(path, account_id, timeout=timeout)


async def get_excel_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str, account_id: str | None = None, timeout: float = 30.0) -> dict[str, Any] | None:
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
    return await request("GET", path, account_id=account_id, timeout=timeout)


async def update_excel_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str, values: list[list[Any]], account_id: str | None = None, timeout: float = 30.0) -> dict[str, Any] | None:
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
    json_data = {"values": values}
    return await request("PATCH", path, account_id=account_id, json=json_data, timeout=timeout)


async def add_excel_table_row(drive_id: str, item_id: str, worksheet_name: str, table_name: str, values: list[list[Any]], account_id: str | None = None, timeout: float = 30.0) -> dict[str, Any] | None:
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/tables/{table_name}/rows/add"
    json_data = {"values": values}
    return await request("POST", path, account_id=account_id, json=json_data, timeout=timeout)


async def download_file(drive_id: str, item_id: str, account_id: str | None = None, timeout: float = 60.0) -> bytes | None:
    path = f"/drives/{drive_id}/items/{item_id}/content"
    async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
        headers = {"Authorization": f"Bearer {get_token(account_id)}"}
        response = await client.get(f"{BASE_URL}{path}", headers=headers)
        response.raise_for_status()
        return response.content


async def upload_small_file(drive_id: str, parent_id: str, filename: str, data: bytes, account_id: str | None = None, timeout: float = 30.0) -> dict[str, Any] | None:
    if len(data) > 4 * 1024 * 1024:
        raise ValueError("File is larger than 4MB. Use upload_large_file instead.")

    path = f"/drives/{drive_id}/items/{parent_id}:/{filename}:/content"
    return await request("PUT", path, account_id=account_id, data=data, timeout=timeout)