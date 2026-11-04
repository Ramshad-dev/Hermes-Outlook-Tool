from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pytest

from hermes_outlook.auth import GraphAuthenticator
from hermes_outlook.config import Settings
from hermes_outlook.errors import (
    AuthenticationRequiredError,
    InvalidArgumentError,
    OutlookError,
    ResourceNotFoundError,
)
from hermes_outlook.graph import GraphClient, validate_next_link


class DummyAuth:
    def __init__(self) -> None:
        self.calls = 0

    def get_access_token(self, **kwargs: object) -> str:
        self.calls += 1
        if kwargs.get("force_refresh"):
            return "refreshed-token"
        return "initial-token"


def test_foreign_next_link_is_rejected(settings: Settings) -> None:
    with pytest.raises(InvalidArgumentError):
        validate_next_link("https://evil.example/v1.0/me/messages", settings)
    with pytest.raises(InvalidArgumentError):
        validate_next_link("https://graph.microsoft.com/v1.0/users", settings)


def test_graph_retries_throttling(settings: Settings, client_factory: Any) -> None:
    calls: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(request)
        if len(calls) == 1:
            return httpx.Response(429, headers={"retry-after": "0"})
        return httpx.Response(200, json={"ok": True})

    client = client_factory(settings, httpx.MockTransport(handler))
    assert client.get("/me/messages") == {"ok": True}
    assert len(calls) == 2


def test_graph_refreshes_once_on_401(settings: Settings) -> None:
    tokens: list[str | None] = []
    auth = DummyAuth()

    def handler(request: httpx.Request) -> httpx.Response:
        tokens.append(request.headers.get("authorization"))
        if len(tokens) == 1:
            return httpx.Response(401, json={"error": {"code": "InvalidAuthenticationToken"}})
        return httpx.Response(200, json={"ok": True})

    client = GraphClient(
        settings,
        auth,
        client=httpx.Client(transport=httpx.MockTransport(handler)),
        sleep=lambda _: None,
    )
    assert client.get("/me/messages") == {"ok": True}
    assert tokens == ["Bearer initial-token", "Bearer refreshed-token"]


def test_graph_maps_not_found(settings: Settings, client_factory: Any) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            404,
            json={"error": {"code": "ErrorItemNotFound", "message": "missing"}},
        )

    client = client_factory(settings, httpx.MockTransport(handler))
    with pytest.raises(ResourceNotFoundError):
        client.get("/me/messages/id")


def test_graph_does_not_follow_redirects(
    settings: Settings, client_factory: Any
) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return httpx.Response(302, headers={"location": "https://evil.example/"})

    client = client_factory(settings, httpx.MockTransport(handler))
    with pytest.raises(OutlookError):
        client.get("/me/messages")
    assert len(seen) == 1


class FakeApp:
    def __init__(self) -> None:
        self.removed: list[dict[str, Any]] = []
        self.silent_result: dict[str, Any] | None = {"access_token": "silent-token"}
        self.interactive_result: dict[str, Any] = {"access_token": "browser-token"}
        self.accounts = [
            {"username": "user@example.com", "name": "User", "tenant_id": "tenant"}
        ]

    def get_accounts(self, username: str | None = None) -> list[dict[str, Any]]:
        return list(self.accounts)

    def acquire_token_silent(self, scopes, account, **kwargs):
        return self.silent_result

    def acquire_token_interactive(self, scopes, **kwargs):
        return self.interactive_result

    def remove_account(self, account: dict[str, Any]) -> None:
        self.removed.append(account)


def test_authenticator_uses_silent_cache(settings: Settings) -> None:
    app = FakeApp()
    auth = GraphAuthenticator(settings, cache=object(), app=app)
    assert auth.get_access_token() == "silent-token"


def test_authenticator_requires_login_without_cache(settings: Settings) -> None:
    app = FakeApp()
    app.accounts = []
    auth = GraphAuthenticator(settings, cache=object(), app=app)
    with pytest.raises(AuthenticationRequiredError):
        auth.get_access_token()


def test_authenticator_reports_interactive_failure(settings: Settings) -> None:
    app = FakeApp()
    app.accounts = []
    app.interactive_result = {"error": "access_denied", "error_description": "denied"}
    auth = GraphAuthenticator(settings, cache=object(), app=app)
    with pytest.raises(OutlookError, match="access_denied"):
        auth.login()


def test_authenticator_strips_reserved_scopes(settings: Settings) -> None:
    auth = GraphAuthenticator(settings, cache=object(), app=FakeApp())
    assert "offline_access" not in auth.scopes
    assert "Mail.ReadWrite" in auth.scopes


def test_logout_removes_cache(settings: Settings, tmp_path: Path) -> None:
    cache_path = tmp_path / "token_cache.bin"
    cache_path.write_text("ciphertext", encoding="utf-8")
    app = FakeApp()
    auth = GraphAuthenticator(settings, cache=object(), app=app)
    result = auth.logout()
    assert result["removed_accounts"] == 1
    assert not cache_path.exists()
