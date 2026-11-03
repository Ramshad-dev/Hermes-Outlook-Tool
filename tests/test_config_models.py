from __future__ import annotations

from datetime import datetime, timezone

import pytest
from pydantic import ValidationError

from hermes_outlook.config import ConfigurationError, Settings, load_settings
from hermes_outlook.html_text import html_to_text
from hermes_outlook.models import (
    DraftInput,
    ListEmailsInput,
    SearchEmailsInput,
    UpdateDraftInput,
    UpdateEmailInput,
)


def test_missing_client_id(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("GRAPH_CLIENT_ID", raising=False)
    with pytest.raises(ConfigurationError):
        load_settings()


def test_send_scope_is_rejected(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    monkeypatch.setenv("GRAPH_CLIENT_ID", "client")
    monkeypatch.setenv("GRAPH_SCOPES", "Mail.ReadWrite,offline_access,Mail.Send")
    monkeypatch.setenv("OUTLOOK_TOKEN_CACHE", str(tmp_path / "cache.bin"))
    with pytest.raises(ConfigurationError, match="send-capable"):
        load_settings()


def test_insecure_redirect_is_rejected(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    monkeypatch.setenv("GRAPH_CLIENT_ID", "client")
    monkeypatch.setenv("GRAPH_REDIRECT_URI", "http://example.com/callback")
    monkeypatch.setenv("OUTLOOK_TOKEN_CACHE", str(tmp_path / "cache.bin"))
    with pytest.raises(ConfigurationError):
        load_settings()


def test_default_page_size_must_fit_max(settings: Settings) -> None:
    with pytest.raises(ValidationError):
        Settings(
            client_id="client",
            token_cache_path=settings.token_cache_path,
            default_page_size=50,
            max_page_size=10,
        )


def test_html_is_reduced_to_safe_text() -> None:
    value = (
        "<html><head><style>.x{}</style><script>alert('x')</script></head>"
        "<body><p>Hello <b>world</b></p><br><p>Second line</p></body></html>"
    )
    text = html_to_text(value)
    assert "alert" not in text
    assert "Hello world" in text
    assert "Second line" in text


def test_invalid_email_address_is_rejected() -> None:
    with pytest.raises(ValidationError):
        DraftInput(subject="x", to=["not an address"])


def test_draft_rejects_two_body_formats() -> None:
    with pytest.raises(ValidationError):
        DraftInput(subject="x", body_text="a", body_html="<p>a</p>")


def test_update_requires_a_field() -> None:
    with pytest.raises(ValidationError):
        UpdateDraftInput(message_id="id")
    with pytest.raises(ValidationError):
        UpdateEmailInput(message_id="id")


def test_search_requires_a_criterion() -> None:
    with pytest.raises(ValidationError):
        SearchEmailsInput()


def test_date_range_is_ordered() -> None:
    with pytest.raises(ValidationError):
        ListEmailsInput(
            received_after=datetime(2026, 2, 1, tzinfo=timezone.utc),
            received_before=datetime(2026, 1, 1, tzinfo=timezone.utc),
        )
