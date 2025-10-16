import os
import msal
import pathlib as pl
from typing import NamedTuple
from dotenv import load_dotenv

load_dotenv()

CONFIG_DIR = pl.Path.home() / ".microsoft-mcp"
TOKEN_CACHE_FILE = CONFIG_DIR / "token_cache.json"
SCOPES = ["https://graph.microsoft.com/.default"]


class Account(NamedTuple):
    username: str
    account_id: str


def _read_cache() -> str | None:
    """Reads the MSAL token cache from the filesystem."""
    try:
        return TOKEN_CACHE_FILE.read_text()
    except FileNotFoundError:
        return None


def _write_cache(content: str) -> None:
    """Writes the MSAL token cache to the filesystem."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_FILE.write_text(content)


def get_client_id() -> str:
    """
    Retrieves the Client ID from environment variables or a cached config file.
    Priority:
    1. GRAPH_CLIENT_ID environment variable.
    2. 'client_id' field in ~/.microsoft-mcp/config.json.
    """
    client_id = os.getenv("GRAPH_CLIENT_ID")
    if client_id:
        return client_id

    config_file = CONFIG_DIR / "config.json"
    if config_file.exists():
        config = json.loads(config_file.read_text())
        client_id = config.get("client_id")
        if client_id:
            return client_id

    raise ValueError("GRAPH_CLIENT_ID is not set as an environment variable or configured via the 'set_client_id' tool.")


def get_app() -> msal.PublicClientApplication:
    client_id = get_client_id()

    tenant_id = os.getenv("GRAPH_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    cache = msal.SerializableTokenCache()
    cache_content = _read_cache()
    if cache_content:
        cache.deserialize(cache_content)

    app = msal.PublicClientApplication(
        client_id, authority=authority, token_cache=cache
    )

    return app


def get_token(account_id: str | None = None) -> str:
    app = get_app()

    accounts = app.get_accounts()
    account = None

    if account_id:
        account = next(
            (a for a in accounts if a["home_account_id"] == account_id), None
        )
    elif accounts:
        account = accounts[0]

    result = app.acquire_token_silent(SCOPES, account=account)

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception(
                f"Failed to get device code: {flow.get('error_description', 'Unknown error')}"
            )
        verification_uri = flow.get(
            "verification_uri",
            flow.get("verification_url", "https://microsoft.com/devicelogin"),
        )
        print(
            f"\nTo authenticate:\n1. Visit {verification_uri}\n2. Enter code: {flow['user_code']}"
        )
        result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        raise Exception(
            f"Auth failed: {result.get('error_description', result['error'])}"
        )

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    return result["access_token"]


def list_accounts() -> list[Account]:
    app = get_app()
    return [
        Account(username=a["username"], account_id=a["home_account_id"])
        for a in app.get_accounts()
    ]


def authenticate_new_account() -> Account | None:
    """Authenticate a new account interactively"""
    app = get_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(
            f"Failed to get device code: {flow.get('error_description', 'Unknown error')}"
        )

    print("\nTo authenticate:")
    print(
        f"1. Visit: {flow.get('verification_uri', flow.get('verification_url', 'https://microsoft.com/devicelogin'))}"
    )
    print(f"2. Enter code: {flow['user_code']}")
    print("3. Sign in with your Microsoft account")
    print("\nWaiting for authentication...")

    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        raise Exception(
            f"Auth failed: {result.get('error_description', result['error'])}"
        )

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

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
                return Account(
                    username=account["username"], account_id=account["home_account_id"]
                )
        # If exact match not found, return the last account
        account = accounts[-1]
        return Account(
            username=account["username"], account_id=account["home_account_id"]
        )

    return None