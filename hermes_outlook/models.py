from __future__ import annotations

from datetime import datetime, timezone
from email.utils import parseaddr
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


def _normalize_address(value: str) -> str:
    value = value.strip()
    if not value:
        raise ValueError("email address must not be empty")
    name, address = parseaddr(value)
    if not address or "@" not in address or any(character.isspace() for character in address):
        raise ValueError(f"invalid email address: {value}")
    if name:
        return f"{name} <{address}>"
    return address


def _normalize_addresses(values: list[str]) -> list[str]:
    return list(dict.fromkeys(_normalize_address(value) for value in values))


def _normalize_datetime(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    if value.tzinfo is None:
        return value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


class InputModel(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class PaginatedInput(InputModel):
    limit: int = Field(default=20, ge=1, le=100)
    next_link: str | None = Field(default=None, max_length=4096)


class DateRangeInput(PaginatedInput):
    received_after: datetime | None = None
    received_before: datetime | None = None

    @field_validator("received_after", "received_before")
    @classmethod
    def _validate_datetime(cls, value: datetime | None) -> datetime | None:
        return _normalize_datetime(value)

    @model_validator(mode="after")
    def _validate_range(self) -> DateRangeInput:
        if (
            self.received_after
            and self.received_before
            and self.received_after > self.received_before
        ):
            raise ValueError("received_after must not be later than received_before")
        return self


class ListEmailsInput(DateRangeInput):
    folder: str = Field(default="inbox", min_length=1, max_length=255)
    unread_only: bool = False


class SearchEmailsInput(DateRangeInput):
    query: str | None = Field(default=None, max_length=500)
    sender: str | None = Field(default=None, max_length=320)
    subject: str | None = Field(default=None, max_length=500)
    folder: str | None = Field(default=None, min_length=1, max_length=255)

    @field_validator("query", "sender", "subject", "folder")
    @classmethod
    def _blank_to_none(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        return value or None

    @model_validator(mode="after")
    def _validate_search(self) -> SearchEmailsInput:
        if not any((self.query, self.sender, self.subject)):
            raise ValueError("provide at least one of query, sender, or subject")
        return self


class GetEmailInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)


class ConversationInput(PaginatedInput):
    conversation_id: str = Field(min_length=1, max_length=2048)
    include_body: bool = False


class ListFoldersInput(InputModel):
    include_children: bool = True


class AttachmentsInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)


class DraftInput(InputModel):
    subject: str | None = Field(default=None, max_length=1000)
    body_text: str | None = Field(default=None, max_length=200000)
    body_html: str | None = Field(default=None, max_length=200000)
    to: list[str] = Field(default_factory=list, max_length=100)
    cc: list[str] = Field(default_factory=list, max_length=100)
    bcc: list[str] = Field(default_factory=list, max_length=100)
    importance: Literal["low", "normal", "high"] = "normal"
    categories: list[str] = Field(default_factory=list, max_length=50)

    @field_validator("to", "cc", "bcc")
    @classmethod
    def _validate_addresses(cls, values: list[str]) -> list[str]:
        return _normalize_addresses(values)

    @field_validator("categories")
    @classmethod
    def _validate_categories(cls, values: list[str]) -> list[str]:
        cleaned = [value.strip() for value in values if value.strip()]
        if any(len(value) > 255 for value in cleaned):
            raise ValueError("category names must be 255 characters or fewer")
        return list(dict.fromkeys(cleaned))

    @model_validator(mode="after")
    def _validate_draft(self) -> DraftInput:
        if self.body_text is not None and self.body_html is not None:
            raise ValueError("provide body_text or body_html, not both")
        if not any((self.subject, self.body_text, self.body_html)):
            raise ValueError("a draft needs a subject or a body")
        return self


class UpdateDraftInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)
    subject: str | None = Field(default=None, max_length=1000)
    body_text: str | None = Field(default=None, max_length=200000)
    body_html: str | None = Field(default=None, max_length=200000)
    to: list[str] | None = Field(default=None, max_length=100)
    cc: list[str] | None = Field(default=None, max_length=100)
    bcc: list[str] | None = Field(default=None, max_length=100)
    importance: Literal["low", "normal", "high"] | None = None
    categories: list[str] | None = Field(default=None, max_length=50)

    @field_validator("to", "cc", "bcc")
    @classmethod
    def _validate_addresses(cls, values: list[str] | None) -> list[str] | None:
        return None if values is None else _normalize_addresses(values)

    @field_validator("categories")
    @classmethod
    def _validate_categories(cls, values: list[str] | None) -> list[str] | None:
        if values is None:
            return None
        cleaned = [value.strip() for value in values if value.strip()]
        if any(len(value) > 255 for value in cleaned):
            raise ValueError("category names must be 255 characters or fewer")
        return list(dict.fromkeys(cleaned))

    @model_validator(mode="after")
    def _validate_update(self) -> UpdateDraftInput:
        if self.body_text is not None and self.body_html is not None:
            raise ValueError("provide body_text or body_html, not both")
        if not self.model_fields_set - {"message_id"}:
            raise ValueError("provide at least one draft field to update")
        return self


class MoveEmailInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)
    destination_folder: str = Field(min_length=1, max_length=255)


class CopyEmailInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)
    destination_folder: str = Field(min_length=1, max_length=255)


class ArchiveEmailInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)


class UpdateEmailInput(InputModel):
    message_id: str = Field(min_length=1, max_length=2048)
    is_read: bool | None = None
    flag_status: Literal["notFlagged", "flagged", "complete"] | None = None
    categories: list[str] | None = Field(default=None, max_length=50)

    @field_validator("categories")
    @classmethod
    def _validate_categories(cls, values: list[str] | None) -> list[str] | None:
        if values is None:
            return None
        cleaned = [value.strip() for value in values if value.strip()]
        if any(len(value) > 255 for value in cleaned):
            raise ValueError("category names must be 255 characters or fewer")
        return list(dict.fromkeys(cleaned))

    @model_validator(mode="after")
    def _validate_update(self) -> UpdateEmailInput:
        if not self.model_fields_set - {"message_id"}:
            raise ValueError("provide is_read, flag_status, or categories")
        return self


class CreateFolderInput(InputModel):
    display_name: str = Field(min_length=1, max_length=255)
    parent_folder: str | None = Field(default=None, min_length=1, max_length=255)

    @field_validator("display_name")
    @classmethod
    def _validate_name(cls, value: str) -> str:
        value = value.strip()
        if any(ord(character) < 32 for character in value):
            raise ValueError("folder name must not contain control characters")
        if value in {".", ".."}:
            raise ValueError("invalid folder name")
        return value
