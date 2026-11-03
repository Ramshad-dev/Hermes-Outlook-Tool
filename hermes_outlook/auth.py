from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import msal
from msal_extensions import PersistedTokenCache, build_encrypted_persistence

from .config import Settings
from .errors import AuthenticationRequiredError, OutlookError, TokenCacheError

logger = logging.getLogger(__name__)


def _authority(tenant_id: str) -> str:
    return f"https://login.microsoftonline.com/{tenant_id}"


def _safe_result_error(result: dict[str, Any] | None) -> str:
    if not result:
        return "Microsoft sign-in did not return a result"
    error = result.get("error") or "unknown_error"
    description = result.get("error_description") or "no description"
    return f"Microsoft sign-in failed ({error}): {description}"


class GraphAuthenticator:
    def __init__(
        self,
        settings: Settings,
        *,
        cache: Any | None = None,
        app: Any | None = None,
    ) -> None:
        self.settings = settings
        self.cache_path = settings.token_cache_path
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise TokenCacheError(
                f"Cannot create token cache directory {self.cache_path.parent}: {exc}"
            ) from exc
        if cache is None:
            try:
                persistence = build_encrypted_persistence(str(self.cache_path))
            except Exception as exc:
                raise TokenCacheError(
                    "Secure token storage is unavailable. The Outlook plugin requires "
                    "Windows DPAPI, macOS Keychain, or Linux libsecret."
                ) from exc
            cache = PersistedTokenCache(persistence)
        self.cache = cache
        try:
            self.app = app or msal.PublicClientApplication(
                client_id=settings.client_id,
                authority=_authority(settings.tenant_id),
                token_cache=self.cache,
            )
        except Exception as exc:
            raise OutlookError(f"Cannot initialize Microsoft authentication: {exc}") from exc

    @property
    def scopes(self) -> list[str]:
        return list(self.settings.scopes)

    def _accounts(self) -> list[dict[str, Any]]:
        try:
            if self.settings.mailbox_user:
                return list(self.app.get_accounts(username=self.settings.mailbox_user))
            return list(self.app.get_accounts())
        except Exception as exc:
            raise OutlookError(f"Cannot read the Outlook account cache: {exc}") from exc

    def _select_account(self) -> dict[str, Any] | None:
        accounts = self._accounts()
        return accounts[0] if accounts else None

    def get_access_token(
        self,
        *,
        interactive: bool = False,
        force_refresh: bool = False,
        force_interactive: bool = False,
    ) -> str:
        account = self._select_account()
        if account and not force_interactive:
            result = self.app.acquire_token_silent(
                self.scopes,
                account=account,
                force_refresh=force_refresh,
            )
            if result and result.get("access_token"):
                return str(result["access_token"])
        if not interactive:
            raise AuthenticationRequiredError(
                "No valid Outlook session is available. Run `hermes outlook login` "
                "to sign in, then retry."
            )
        try:
            result = self.app.acquire_token_interactive(
                self.scopes,
                prompt="select_account" if force_interactive else None,
                login_hint=self.settings.mailbox_user,
                port=self.settings.redirect_port,
                timeout=300,
            )
        except Exception as exc:
            raise OutlookError(f"Interactive Microsoft sign-in failed: {exc}") from exc
        if not result or not result.get("access_token"):
            raise OutlookError(_safe_result_error(result))
        return str(result["access_token"])

    def login(self) -> dict[str, Any]:
        self.get_access_token(interactive=True, force_interactive=True)
        accounts = self._accounts()
        if not accounts:
            return {"authenticated": True, "account": None}
        return {"authenticated": True, "account": self._account_summary(accounts[0])}

    def status(self) -> dict[str, Any]:
        accounts = self._accounts()
        if not accounts:
            return {"authenticated": False, "account": None, "cache": str(self.cache_path)}
        account = accounts[0]
        result = self.app.acquire_token_silent(self.scopes, account=account)
        return {
            "authenticated": bool(result and result.get("access_token")),
            "account": self._account_summary(account),
            "cache": str(self.cache_path),
        }

    def logout(self) -> dict[str, Any]:
        accounts = self._accounts()
        for account in accounts:
            try:
                self.app.remove_account(account)
            except Exception:
                logger.debug("Unable to remove one cached account", exc_info=True)
        for path in (self.cache_path, Path(f"{self.cache_path}.lockfile")):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                logger.debug("Unable to remove cache file %s", path, exc_info=True)
        return {"authenticated": False, "removed_accounts": len(accounts)}

    @staticmethod
    def _account_summary(account: dict[str, Any]) -> dict[str, Any]:
        return {
            "username": account.get("username"),
            "name": account.get("name"),
            "tenant_id": account.get("tenant_id") or account.get("realm"),
            "environment": account.get("environment"),
        }
