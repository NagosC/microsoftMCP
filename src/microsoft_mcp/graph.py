import httpx
import time
import logging
from typing import Any
from .auth import get_token, Account

BASE_URL = "https://graph.microsoft.com/v1.0"
# 15 x 320 KiB = 4,915,200 bytes
UPLOAD_CHUNK_SIZE = 15 * 320 * 1024

logger = logging.getLogger(__name__)


def request(
    method: str,
    path: str,
    account_id: str | None = None,
    params: dict[str, Any] | None = None,
    json: dict[str, Any] | None = None,
    data: bytes | None = None,
    max_retries: int = 3,
) -> dict[str, Any] | None:
    """
    Makes a request to the Microsoft Graph API with authentication and retry logic.
    """
    with httpx.Client(timeout=30.0, follow_redirects=True) as client:
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
                response = client.request(
                    method=method,
                    url=f"{BASE_URL}{path}",
                    headers=headers,
                    params=params,
                    json=json,
                    content=data,
                )

                # Handle rate limiting
                if response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After", "5"))
                    if attempt < max_retries:
                        logger.warning(
                            f"Rate limited. Retrying after {retry_after} seconds."
                        )
                        time.sleep(min(retry_after, 60))
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
                    wait_time = (2**attempt) * 1  # Exponential backoff
                    logger.warning(
                        f"Request failed (attempt {attempt + 1}/{max_retries + 1}). Retrying in {wait_time}s. Error: {e}"
                    )
                    time.sleep(wait_time)
                    continue
                raise  # Re-raise the exception if max retries are exceeded

    return None


def _paginated_request(path: str, account_id: str | None) -> list[dict[str, Any]]:
    """
    Handles paginated requests to the Graph API.
    """
    results = []
    next_url = path

    while next_url:
        # For subsequent requests, the URL is absolute and includes the base URL
        path_to_request = next_url if next_url.startswith(BASE_URL) else f"{BASE_URL}{next_url}"
        
        response = request(
            "GET",
            path_to_request.replace(BASE_URL, ""), # request function adds BASE_URL
            account_id=account_id
        )

        if response and "value" in response:
            results.extend(response["value"])
            next_url = response.get("@odata.nextLink")
        else:
            break

    return results


# --- SharePoint/Drive Functions ---

def get_site(hostname: str, relative_path: str, account_id: str | None = None) -> dict[str, Any] | None:
    """
    Gets a SharePoint site by its hostname and relative server path.
    e.g., get_site("contoso.sharepoint.com", "/sites/MySite")
    """
    path = f"/sites/{hostname}:{relative_path}"
    return request("GET", path, account_id=account_id)


def get_drives(site_id: str, account_id: str | None = None) -> list[dict[str, Any]]:
    """
    Lists all Document Libraries (Drives) in a given SharePoint site.
    """
    path = f"/sites/{site_id}/drives"
    return _paginated_request(path, account_id)


def list_drive_items(drive_id: str, item_id: str | None = None, account_id: str | None = None) -> list[dict[str, Any]]:
    """
    Lists items in a drive's root or in a specific folder (item_id).
    """
    if item_id:
        path = f"/drives/{drive_id}/items/{item_id}/children"
    else:
        path = f"/drives/{drive_id}/root/children"
    return _paginated_request(path, account_id)


# --- Excel Functions ---

def get_excel_worksheets(drive_id: str, item_id: str, account_id: str | None = None) -> list[dict[str, Any]]:
    """
    Lists all worksheets in an Excel file.
    """
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets"
    return _paginated_request(path, account_id)


def get_excel_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str, account_id: str | None = None) -> dict[str, Any] | None:
    """
    Gets data from a specified range in an Excel worksheet.
    e.g., get_excel_range(..., "Sheet1", "A1:C3")
    """
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
    return request("GET", path, account_id=account_id)


def update_excel_range(drive_id: str, item_id: str, worksheet_name: str, range_address: str, values: list[list[Any]], account_id: str | None = None) -> dict[str, Any] | None:
    """
    Updates a range in an Excel worksheet with the given values.
    """
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
    json_data = {"values": values}
    return request("PATCH", path, account_id=account_id, json=json_data)


# --- File Download/Upload Functions ---

def download_file(drive_id: str, item_id: str, account_id: str | None = None) -> bytes | None:
    """
    Downloads the content of a file from a drive.
    """
    path = f"/drives/{drive_id}/items/{item_id}/content"
    # This endpoint returns raw content, not JSON, so we need a custom request.
    with httpx.Client(timeout=60.0, follow_redirects=True) as client:
        headers = {"Authorization": f"Bearer {get_token(account_id)}"}
        response = client.get(f"{BASE_URL}{path}", headers=headers)
        response.raise_for_status()
        return response.content


def upload_small_file(drive_id: str, parent_id: str, filename: str, data: bytes, account_id: str | None = None) -> dict[str, Any] | None:
    """
    Uploads a small file (under 4MB) to a specific folder in a drive.
    Use 'root' for parent_id to upload to the root folder.
    """
    if len(data) > 4 * 1024 * 1024:
        raise ValueError("File is larger than 4MB. Use upload_large_file instead.")

    path = f"/drives/{drive_id}/items/{parent_id}:/{filename}:/content"
    return request("PUT", path, account_id=account_id, data=data)
