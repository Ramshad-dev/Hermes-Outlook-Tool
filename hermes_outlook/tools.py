from __future__ import annotations

import json
import logging
import threading
from typing import Any, Callable

from pydantic import BaseModel, ValidationError

from .config import ConfigurationError, load_settings
from .errors import InvalidArgumentError, OutlookError
from .models import (
    ArchiveEmailInput,
    AttachmentsInput,
    ConversationInput,
    CopyEmailInput,
    CreateFolderInput,
    DraftInput,
    GetEmailInput,
    ListEmailsInput,
    ListFoldersInput,
    MoveEmailInput,
    SearchEmailsInput,
    UpdateDraftInput,
    UpdateEmailInput,
)
from .services import OutlookService

logger = logging.getLogger(__name__)

_service: OutlookService | None = None
_service_lock = threading.Lock()


def get_service() -> OutlookService:
    global _service
    with _service_lock:
        if _service is None:
            try:
                settings = load_settings()
            except ConfigurationError as exc:
                raise InvalidArgumentError(
                    str(exc), code="CONFIGURATION_ERROR"
                ) from exc
            _service = OutlookService(settings)
        return _service


def reset_service() -> None:
    global _service
    with _service_lock:
        if _service is not None:
            _service.close()
            _service = None


def _success(data: dict[str, Any]) -> str:
    return json.dumps(
        {"success": True, "data": data, "error": None},
        ensure_ascii=False,
        default=str,
    )


def _failure(error: OutlookError) -> str:
    return json.dumps(
        {
            "success": False,
            "data": None,
            "error": {
                "code": error.code,
                "message": error.message,
                "retryable": error.retryable,
                "details": error.details,
            },
        },
        ensure_ascii=False,
        default=str,
    )


def _validation_failure(exc: ValidationError) -> str:
    details = [
        {
            "field": ".".join(str(part) for part in error.get("loc", ())),
            "message": error.get("msg", "invalid value"),
        }
        for error in exc.errors()
    ]
    error = InvalidArgumentError("Invalid arguments for the Outlook tool.", details=details)
    return _failure(error)


def _execute(
    model: type[BaseModel],
    args: dict[str, Any] | None,
    action: Callable[[OutlookService, Any], dict[str, Any]],
) -> str:
    try:
        request = model.model_validate(args or {})
        result = action(get_service(), request)
        return _success(result)
    except ValidationError as exc:
        return _validation_failure(exc)
    except OutlookError as exc:
        return _failure(exc)
    except Exception:
        logger.exception("Unexpected Outlook plugin error")
        return _failure(
            OutlookError(
                "Unexpected Outlook plugin error. Check the Hermes logs for details.",
                code="INTERNAL_ERROR",
            )
        )


def list_emails(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(ListEmailsInput, args, lambda service, request: service.list_emails(request))


def search_emails(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        SearchEmailsInput, args, lambda service, request: service.search_emails(request)
    )


def get_email(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        GetEmailInput,
        args,
        lambda service, request: service.get_email(request.message_id),
    )


def get_conversation(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        ConversationInput,
        args,
        lambda service, request: service.get_conversation(request),
    )


def list_folders(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        ListFoldersInput,
        args,
        lambda service, request: service.list_folders(request),
    )


def list_attachments(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        AttachmentsInput,
        args,
        lambda service, request: service.list_attachments(request),
    )


def create_draft(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        DraftInput, args, lambda service, request: service.create_draft(request)
    )


def update_draft(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        UpdateDraftInput, args, lambda service, request: service.update_draft(request)
    )


def move_email(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        MoveEmailInput, args, lambda service, request: service.move_email(request)
    )


def copy_email(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        CopyEmailInput, args, lambda service, request: service.copy_email(request)
    )


def archive_email(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        ArchiveEmailInput, args, lambda service, request: service.archive_email(request)
    )


def update_email(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        UpdateEmailInput, args, lambda service, request: service.update_email(request)
    )


def create_folder(args: dict[str, Any] | None = None, **kwargs: Any) -> str:
    return _execute(
        CreateFolderInput, args, lambda service, request: service.create_folder(request)
    )


TOOL_HANDLERS: dict[str, Callable[..., str]] = {
    "outlook_list_emails": list_emails,
    "outlook_search_emails": search_emails,
    "outlook_get_email": get_email,
    "outlook_get_conversation": get_conversation,
    "outlook_list_folders": list_folders,
    "outlook_list_attachments": list_attachments,
    "outlook_create_draft": create_draft,
    "outlook_update_draft": update_draft,
    "outlook_move_email": move_email,
    "outlook_copy_email": copy_email,
    "outlook_archive_email": archive_email,
    "outlook_update_email": update_email,
    "outlook_create_folder": create_folder,
}
