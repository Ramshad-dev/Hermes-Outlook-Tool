from __future__ import annotations

import argparse
import json
from typing import Any

from .auth import GraphAuthenticator
from .config import ConfigurationError, Settings, load_settings
from .errors import OutlookError
from .graph import GraphClient


def _print(value: Any) -> None:
    print(json.dumps(value, indent=2, ensure_ascii=False, default=str))


def _settings() -> Settings:
    return load_settings()


def _safe_settings(settings: Settings) -> dict[str, Any]:
    return {
        "tenant_id": settings.tenant_id,
        "redirect_uri": settings.redirect_uri,
        "graph_base_url": settings.graph_base_url,
        "scopes": list(settings.scopes),
        "token_cache": str(settings.token_cache_path),
        "mailbox_user": settings.mailbox_user,
        "default_page_size": settings.default_page_size,
        "max_page_size": settings.max_page_size,
    }


def _handle(args: argparse.Namespace) -> int:
    subcommand = getattr(args, "outlook_command", None)
    if not subcommand:
        print("Usage: hermes outlook <login|status|logout|test|config>")
        return 0
    try:
        settings = _settings()
        if subcommand == "config":
            _print(_safe_settings(settings))
            return 0
        authenticator = GraphAuthenticator(settings)
        if subcommand == "login":
            _print(authenticator.login())
            return 0
        if subcommand == "status":
            _print(authenticator.status())
            return 0
        if subcommand == "logout":
            _print(authenticator.logout())
            return 0
        if subcommand == "test":
            with GraphClient(settings, authenticator) as graph:
                folder = graph.get(
                    "/me/mailFolders/inbox",
                    params={
                        "$select": (
                            "id,displayName,totalItemCount,unreadItemCount"
                        )
                    },
                )
            _print(
                {
                    "ok": True,
                    "account": authenticator.status().get("account"),
                    "inbox": {
                        "folder_id": folder.get("id"),
                        "display_name": folder.get("displayName"),
                        "total_item_count": folder.get("totalItemCount"),
                        "unread_item_count": folder.get("unreadItemCount"),
                    },
                }
            )
            return 0
        print(f"Unknown subcommand: {subcommand}")
        return 2
    except ConfigurationError as exc:
        print(f"Configuration error: {exc}")
        return 1
    except OutlookError as exc:
        print(f"{exc.code}: {exc.message}")
        return 1


def _setup(subparser: argparse.ArgumentParser) -> None:
    subs = subparser.add_subparsers(dest="outlook_command")
    subs.add_parser("login", help="Open a browser and sign in to Outlook")
    subs.add_parser("status", help="Show the current Outlook session")
    subs.add_parser("logout", help="Remove the cached Outlook session")
    subs.add_parser("test", help="Verify Microsoft Graph mailbox access")
    subs.add_parser("config", help="Show non-secret Outlook configuration")
    subparser.set_defaults(func=_handle)


def register_cli(ctx: Any) -> None:
    ctx.register_cli_command(
        name="outlook",
        help="Manage the Outlook mail integration",
        setup_fn=_setup,
        handler_fn=_handle,
        description="Outlook mail sign-in, status, and connectivity checks.",
    )
