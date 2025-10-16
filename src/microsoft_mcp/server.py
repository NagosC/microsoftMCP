import os
import sys
from . import auth
from .tools import mcp # noqa


def main() -> None:
    try:
        auth.get_client_id()
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("\nHint: Set the MICROSOFT_MCP_CLIENT_ID environment variable, or use the 'set_client_id' tool.", file=sys.stderr)
        sys.exit(1)

    mcp.run()


if __name__ == "__main__":
    main()