from __future__ import annotations

from typing import Any


class OutlookError(Exception):
    code = "OUTLOOK_ERROR"

    def __init__(
        self,
        message: str,
        *,
        code: str | None = None,
        retryable: bool = False,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        if code:
            self.code = code
        self.retryable = retryable
        self.details = details or {}


class ConfigurationToolError(OutlookError):
    code = "CONFIGURATION_ERROR"


class AuthenticationRequiredError(OutlookError):
    code = "AUTH_REQUIRED"


class PermissionDeniedError(OutlookError):
    code = "PERMISSION_DENIED"


class ResourceNotFoundError(OutlookError):
    code = "NOT_FOUND"


class InvalidArgumentError(OutlookError):
    code = "INVALID_ARGUMENT"


class GraphUnavailableError(OutlookError):
    code = "GRAPH_UNAVAILABLE"


class GraphTimeoutError(OutlookError):
    code = "GRAPH_TIMEOUT"


class GraphResponseError(OutlookError):
    code = "GRAPH_RESPONSE_ERROR"


class TokenCacheError(OutlookError):
    code = "TOKEN_CACHE_ERROR"


class NotADraftError(OutlookError):
    code = "NOT_A_DRAFT"
