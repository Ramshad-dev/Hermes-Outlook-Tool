from __future__ import annotations

import json
from typing import Any

import pytest

from hermes_outlook import tools
from hermes_outlook.errors import AuthenticationRequiredError
from hermes_outlook.plugin import register
from hermes_outlook.schemas import TOOL_SCHEMAS
from hermes_outlook.tools import TOOL_HANDLERS


class FakeService:
    def __init__(self) -> None:
        self.last: Any = None

    def list_emails(self, request: Any) -> dict[str, Any]:
        self.last = request
        return {"emails": [], "count": 0, "next_link": None}

    def get_email(self, message_id: str) -> dict[str, Any]:
        self.last = message_id
        return {"message_id": message_id}


def test_tool_success_envelope(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeService()
    monkeypatch.setattr(tools, "_service", fake)
    payload = json.loads(tools.list_emails({"limit": 5}))
    assert payload["success"] is True
    assert payload["data"]["count"] == 0
    assert fake.last.limit == 5


def test_tool_validation_error(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(tools, "_service", FakeService())
    payload = json.loads(tools.list_emails({"limit": 0}))
    assert payload["success"] is False
    assert payload["error"]["code"] == "INVALID_ARGUMENT"
    assert payload["error"]["details"][0]["field"] == "limit"


def test_tool_missing_required_argument(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(tools, "_service", FakeService())
    payload = json.loads(tools.get_email({}))
    assert payload["success"] is False
    assert payload["error"]["code"] == "INVALID_ARGUMENT"


def test_tool_authentication_error(monkeypatch: pytest.MonkeyPatch) -> None:
    class FailingService(FakeService):
        def list_emails(self, request: Any) -> dict[str, Any]:
            raise AuthenticationRequiredError("log in")

    monkeypatch.setattr(tools, "_service", FailingService())
    payload = json.loads(tools.list_emails({}))
    assert payload["success"] is False
    assert payload["error"]["code"] == "AUTH_REQUIRED"
    assert payload["error"]["retryable"] is False


def test_tool_registry_is_complete() -> None:
    assert set(TOOL_SCHEMAS) == set(TOOL_HANDLERS)
    assert not any("send" in name or "delete" in name for name in TOOL_SCHEMAS)
    for name, schema in TOOL_SCHEMAS.items():
        assert schema["name"] == name
        assert schema["parameters"]["type"] == "object"
        assert schema["parameters"]["additionalProperties"] is False


def test_plugin_registers_tools_and_cli() -> None:
    class FakeCtx:
        def __init__(self) -> None:
            self.tools: list[dict[str, Any]] = []
            self.cli: list[dict[str, Any]] = []

        def register_tool(self, **kwargs: Any) -> None:
            self.tools.append(kwargs)

        def register_cli_command(self, **kwargs: Any) -> None:
            self.cli.append(kwargs)

    ctx = FakeCtx()
    register(ctx)
    assert len(ctx.tools) == len(TOOL_SCHEMAS) == 13
    assert all(tool["toolset"] == "outlook" for tool in ctx.tools)
    assert all(callable(tool["check_fn"]) for tool in ctx.tools)
    assert ctx.cli[0]["name"] == "outlook"


def test_manifest_has_no_send_or_delete_capability() -> None:
    manifest = (
        __import__("pathlib").Path(__file__).resolve().parents[1] / "plugin.yaml"
    ).read_text(encoding="utf-8")
    assert "outlook_send" not in manifest
    assert "outlook_delete" not in manifest
    assert "outlook_create_draft" in manifest
