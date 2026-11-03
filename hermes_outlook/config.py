from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import urlparse

from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

DEFAULT_SCOPES = ("Mail.ReadWrite", "offline_access")
ALLOWED_SCOPES = {
    "Mail.Read",
    "Mail.ReadWrite",
    "offline_access",
    "openid",
    "profile",
    "User.Read",
}
FORBIDDEN_SCOPES = {"Mail.Send", "Mail.ReadWrite.All", "Mail.Read.All"}
ALLOWED_GRAPH_HOSTS = {
    "graph.microsoft.com",
    "graph.microsoft.us",
    "dod-graph.microsoft.us",
    "microsoftgraph.chinacloudapi.cn",
}


class ConfigurationError(Exception):
    pass


class Settings(BaseModel):
    client_id: str
    tenant_id: str = "organizations"
    redirect_uri: str = "http://localhost:8400"
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    token_cache_path: Path
    scopes: tuple[str, ...] = DEFAULT_SCOPES
    default_page_size: int = Field(default=20, ge=1, le=100)
    max_page_size: int = Field(default=100, ge=1, le=100)
    request_timeout: float = Field(default=30.0, gt=0, le=120)
    max_retries: int = Field(default=3, ge=0, le=5)
    max_body_chars: int = Field(default=20000, ge=1000, le=200000)
    max_preview_chars: int = Field(default=500, ge=100, le=2000)
    mailbox_user: str | None = None

    @field_validator("client_id")
    @classmethod
    def _validate_client_id(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("GRAPH_CLIENT_ID must not be empty")
        if any(character.isspace() for character in value):
            raise ValueError("GRAPH_CLIENT_ID must not contain whitespace")
        return value

    @field_validator("tenant_id")
    @classmethod
    def _validate_tenant_id(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("GRAPH_TENANT_ID must not be empty")
        return value

    @field_validator("redirect_uri")
    @classmethod
    def _validate_redirect_uri(cls, value: str) -> str:
        value = value.strip()
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            raise ValueError("GRAPH_REDIRECT_URI must be an http or https URL")
        if parsed.scheme == "http" and parsed.hostname not in {"localhost", "127.0.0.1", "::1"}:
            raise ValueError("GRAPH_REDIRECT_URI must use https unless it targets localhost")
        if parsed.query or parsed.fragment:
            raise ValueError("GRAPH_REDIRECT_URI must not include a query or fragment")
        return value

    @field_validator("graph_base_url")
    @classmethod
    def _validate_graph_base_url(cls, value: str) -> str:
        value = value.strip().rstrip("/")
        parsed = urlparse(value)
        if parsed.scheme != "https":
            raise ValueError("GRAPH_BASE_URL must use https")
        if parsed.hostname not in ALLOWED_GRAPH_HOSTS:
            raise ValueError("GRAPH_BASE_URL must target a supported Microsoft Graph host")
        if parsed.query or parsed.fragment:
            raise ValueError("GRAPH_BASE_URL must not include a query or fragment")
        return value

    @field_validator("scopes")
    @classmethod
    def _validate_scopes(cls, value: tuple[str, ...]) -> tuple[str, ...]:
        scopes = tuple(dict.fromkeys(scope.strip() for scope in value if scope.strip()))
        if not scopes:
            raise ValueError("GRAPH_SCOPES must contain at least Mail.ReadWrite")
        forbidden = sorted(set(scopes) & FORBIDDEN_SCOPES)
        if forbidden:
            raise ValueError(f"Refusing over-broad or send-capable scopes: {', '.join(forbidden)}")
        unsupported = sorted(set(scopes) - ALLOWED_SCOPES)
        if unsupported:
            raise ValueError(f"Unsupported scopes: {', '.join(unsupported)}")
        if "Mail.ReadWrite" not in scopes:
            raise ValueError("Mail.ReadWrite is required for drafts and organization")
        return scopes

    @model_validator(mode="after")
    def _validate_page_sizes(self) -> Settings:
        if self.default_page_size > self.max_page_size:
            raise ValueError("OUTLOOK_DEFAULT_PAGE_SIZE must not exceed OUTLOOK_MAX_PAGE_SIZE")
        return self

    @property
    def redirect_port(self) -> int:
        parsed = urlparse(self.redirect_uri)
        if parsed.port:
            return parsed.port
        return 443 if parsed.scheme == "https" else 80


def _read_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or not raw.strip():
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise ConfigurationError(f"{name} must be an integer") from exc


def _read_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None or not raw.strip():
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise ConfigurationError(f"{name} must be a number") from exc


def _read_scopes() -> tuple[str, ...]:
    raw = os.environ.get("GRAPH_SCOPES", "").strip()
    if not raw:
        return DEFAULT_SCOPES
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def _default_cache_path() -> Path:
    home = Path(os.environ.get("HERMES_HOME", str(Path.home() / ".hermes"))).expanduser()
    return home / "outlook" / "token_cache.bin"


def is_configured() -> bool:
    return bool(os.environ.get("GRAPH_CLIENT_ID", "").strip())


def load_settings() -> Settings:
    client_id = os.environ.get("GRAPH_CLIENT_ID", "").strip()
    if not client_id:
        raise ConfigurationError(
            "GRAPH_CLIENT_ID is not set. Register a public client in Microsoft Entra "
            "and add GRAPH_CLIENT_ID to your Hermes environment."
        )
    cache_value = os.environ.get("OUTLOOK_TOKEN_CACHE", "").strip()
    mailbox_user = os.environ.get("OUTLOOK_MAILBOX_USER", "").strip() or None
    try:
        return Settings(
            client_id=client_id,
            tenant_id=os.environ.get("GRAPH_TENANT_ID", "organizations"),
            redirect_uri=os.environ.get("GRAPH_REDIRECT_URI", "http://localhost:8400"),
            graph_base_url=os.environ.get(
                "GRAPH_BASE_URL", "https://graph.microsoft.com/v1.0"
            ),
            token_cache_path=Path(cache_value).expanduser()
            if cache_value
            else _default_cache_path(),
            scopes=_read_scopes(),
            default_page_size=_read_int("OUTLOOK_DEFAULT_PAGE_SIZE", 20),
            max_page_size=_read_int("OUTLOOK_MAX_PAGE_SIZE", 100),
            request_timeout=_read_float("OUTLOOK_REQUEST_TIMEOUT", 30.0),
            max_retries=_read_int("OUTLOOK_MAX_RETRIES", 3),
            max_body_chars=_read_int("OUTLOOK_MAX_BODY_CHARS", 20000),
            max_preview_chars=_read_int("OUTLOOK_MAX_PREVIEW_CHARS", 500),
            mailbox_user=mailbox_user,
        )
    except ValidationError as exc:
        messages = "; ".join(error["msg"] for error in exc.errors())
        raise ConfigurationError(messages) from exc
