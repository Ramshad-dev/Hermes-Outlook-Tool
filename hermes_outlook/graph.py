from __future__ import annotations

import random
import time
from email.utils import parsedate_to_datetime
from typing import Any, Callable
from urllib.parse import quote, urlsplit

import httpx

from .config import Settings
from .errors import (
    AuthenticationRequiredError,
    GraphResponseError,
    GraphTimeoutError,
    GraphUnavailableError,
    InvalidArgumentError,
    OutlookError,
    PermissionDeniedError,
    ResourceNotFoundError,
)
_RETRYABLE_STATUS = {429, 500, 502, 503, 504}
_MAX_RETRY_DELAY = 30.0


def quote_segment(value: str) -> str:
    return quote(value, safe="")


def validate_next_link(next_link: str, settings: Settings) -> str:
    parsed = urlsplit(next_link)
    base = urlsplit(settings.graph_base_url)
    if parsed.scheme != base.scheme or parsed.netloc.lower() != base.netloc.lower():
        raise InvalidArgumentError("next_link must target the configured Microsoft Graph host")
    if parsed.username or parsed.password:
        raise InvalidArgumentError("next_link must not contain credentials")
    if not parsed.path.startswith(f"{base.path}/me/"):
        raise InvalidArgumentError("next_link must stay within the signed-in user's mailbox")
    return next_link


def _retry_delay(response: httpx.Response, attempt: int) -> float:
    retry_after = response.headers.get("retry-after")
    if retry_after:
        try:
            return min(_MAX_RETRY_DELAY, max(0.0, float(retry_after)))
        except ValueError:
            try:
                target = parsedate_to_datetime(retry_after).timestamp()
                return min(_MAX_RETRY_DELAY, max(0.0, target - time.time()))
            except (TypeError, ValueError, OverflowError):
                pass
    return min(_MAX_RETRY_DELAY, (2**attempt) * 0.5 + random.uniform(0.0, 0.25))


def _graph_error(response: httpx.Response) -> OutlookError:
    code = None
    message = None
    request_id = None
    try:
        payload = response.json()
        error = payload.get("error") if isinstance(payload, dict) else None
        if isinstance(error, dict):
            code = error.get("code")
            message = error.get("message")
            inner = error.get("innerError")
            if isinstance(inner, dict):
                request_id = inner.get("request-id") or inner.get("date")
    except ValueError:
        payload = None
    status = response.status_code
    detail = f"Microsoft Graph returned HTTP {status}"
    if code:
        detail += f" ({code})"
    if message:
        detail += f": {message}"
    if request_id:
        detail += f" [request-id: {request_id}]"
    if status == 401:
        return AuthenticationRequiredError(
            "Microsoft Graph rejected the access token. Run `hermes outlook login` and retry.",
            code="AUTH_REQUIRED",
        )
    if status == 403:
        return PermissionDeniedError(detail)
    if status in {404, 410}:
        return ResourceNotFoundError(detail)
    if status == 400 or status == 422:
        if code == "ErrorInvalidIdMalformed":
            detail += (
                " Pass the exact message_id or folder_id returned by the Outlook list "
                "tools; webLink URLs, conversation ids, internetMessageIds, and email "
                "addresses are not valid ids."
            )
        elif (
            code == "ErrorInvalidOperation"
            and message
            and "conversationid" in message.lower()
        ):
            detail += " Use outlook_get_conversation for conversation ids."
        return InvalidArgumentError(detail)
    if status == 429:
        return GraphUnavailableError(detail, retryable=True)
    if status >= 500:
        return GraphUnavailableError(detail, retryable=True)
    return OutlookError(detail, code="GRAPH_ERROR")


class GraphClient:
    def __init__(
        self,
        settings: Settings,
        auth: Any,
        *,
        client: httpx.Client | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self.settings = settings
        self.auth = auth
        self._sleep = sleep
        self._client = client or httpx.Client(
            timeout=httpx.Timeout(settings.request_timeout),
            follow_redirects=False,
        )

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> GraphClient:
        return self

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        self.close()

    def _url(self, path_or_url: str) -> str:
        if path_or_url.startswith("http://") or path_or_url.startswith("https://"):
            return validate_next_link(path_or_url, self.settings)
        if not path_or_url.startswith("/"):
            raise InvalidArgumentError("Graph paths must start with /")
        return f"{self.settings.graph_base_url}{path_or_url}"

    def request(
        self,
        method: str,
        path_or_url: str,
        *,
        params: dict[str, Any] | None = None,
        json: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = self._url(path_or_url)
        token = self.auth.get_access_token()
        refreshed = False
        last_error: OutlookError | None = None
        for attempt in range(self.settings.max_retries + 1):
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
            if json is not None:
                headers["Content-Type"] = "application/json"
            try:
                response = self._client.request(
                    method,
                    url,
                    params=params,
                    json=json,
                    headers=headers,
                )
            except httpx.TimeoutException as exc:
                last_error = GraphTimeoutError(
                    "Microsoft Graph timed out; retry the request.",
                    retryable=True,
                )
                if attempt < self.settings.max_retries:
                    self._sleep(_retry_delay(httpx.Response(503), attempt))
                    continue
                raise last_error from exc
            except httpx.HTTPError as exc:
                last_error = GraphUnavailableError(
                    f"Could not reach Microsoft Graph: {exc}",
                    retryable=True,
                )
                if attempt < self.settings.max_retries:
                    self._sleep(_retry_delay(httpx.Response(503), attempt))
                    continue
                raise last_error from exc
            if response.status_code == 401 and not refreshed:
                token = self.auth.get_access_token(force_refresh=True)
                refreshed = True
                continue
            if response.status_code in _RETRYABLE_STATUS and attempt < self.settings.max_retries:
                last_error = _graph_error(response)
                self._sleep(_retry_delay(response, attempt))
                continue
            if 300 <= response.status_code < 400:
                raise GraphResponseError("Microsoft Graph returned an unexpected redirect")
            if response.status_code >= 400:
                raise _graph_error(response)
            if response.status_code == 204 or not response.content:
                return {}
            try:
                payload = response.json()
            except ValueError as exc:
                raise GraphResponseError("Microsoft Graph returned invalid JSON") from exc
            if not isinstance(payload, dict):
                raise GraphResponseError("Microsoft Graph returned an unexpected response shape")
            return payload
        raise last_error or GraphUnavailableError("Microsoft Graph request failed", retryable=True)

    def get(self, path_or_url: str, *, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.request("GET", path_or_url, params=params)

    def post(
        self,
        path_or_url: str,
        *,
        json: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        return self.request("POST", path_or_url, json=json, params=params)

    def patch(self, path_or_url: str, *, json: dict[str, Any]) -> dict[str, Any]:
        return self.request("PATCH", path_or_url, json=json)

    def delete(self, path_or_url: str) -> dict[str, Any]:
        return self.request("DELETE", path_or_url)
