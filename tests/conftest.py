from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pytest

from hermes_outlook.config import Settings
from hermes_outlook.graph import GraphClient


class DummyAuth:
    def __init__(self) -> None:
        self.calls = 0

    def get_access_token(self, **kwargs: object) -> str:
        self.calls += 1
        if kwargs.get("force_refresh"):
            return "refreshed-token"
        return "initial-token"


@pytest.fixture
def settings(tmp_path: Path) -> Settings:
    return Settings(
        client_id="test-client-id",
        tenant_id="contoso.onmicrosoft.com",
        redirect_uri="http://localhost:8400",
        token_cache_path=tmp_path / "token_cache.bin",
    )


@pytest.fixture
def client_factory() -> Any:
    def factory(settings: Settings, handler: httpx.MockTransport) -> GraphClient:
        http_client = httpx.Client(transport=handler)
        return GraphClient(
            settings,
            DummyAuth(),
            client=http_client,
            sleep=lambda _: None,
        )

    return factory
