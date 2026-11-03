from __future__ import annotations

from email.utils import parseaddr
from typing import Any

from ..config import Settings
from ..errors import GraphResponseError, InvalidArgumentError, NotADraftError
from ..graph import GraphClient, quote_segment
from ..models import DraftInput, UpdateDraftInput
from .base import message_summary
from .message import MessageService


def _recipients(values: list[str]) -> list[dict[str, Any]]:
    recipients = []
    for value in values:
        name, address = parseaddr(value)
        email_address: dict[str, str] = {"address": address}
        if name:
            email_address["name"] = name
        recipients.append({"emailAddress": email_address})
    return recipients


class DraftService:
    def __init__(
        self, graph: GraphClient, settings: Settings, messages: MessageService
    ) -> None:
        self.graph = graph
        self.settings = settings
        self.messages = messages

    def create(self, request: DraftInput) -> dict[str, Any]:
        if request.body_html is not None:
            body = {"contentType": "HTML", "content": request.body_html}
        else:
            body = {"contentType": "Text", "content": request.body_text or ""}
        payload: dict[str, Any] = {
            "subject": request.subject or "",
            "body": body,
            "toRecipients": _recipients(request.to),
            "ccRecipients": _recipients(request.cc),
            "bccRecipients": _recipients(request.bcc),
            "importance": request.importance,
        }
        if request.categories:
            payload["categories"] = request.categories
        created = self.graph.post("/me/messages", json=payload)
        if not created.get("id"):
            raise GraphResponseError("Microsoft Graph did not return a draft id")
        return {
            "message_id": created.get("id"),
            "subject": request.subject or "",
            "is_draft": True,
            "to": request.to,
            "cc": request.cc,
            "bcc": request.bcc,
            "importance": request.importance,
            "categories": request.categories,
            "web_link": created.get("webLink"),
        }

    def update(self, request: UpdateDraftInput) -> dict[str, Any]:
        current = self.messages.get_raw(
            request.message_id, select="id,isDraft,subject"
        )
        if not current.get("isDraft"):
            raise NotADraftError(
                "Only draft messages can be edited with outlook_update_draft."
            )
        patch: dict[str, Any] = {}
        fields = request.model_fields_set
        if "subject" in fields:
            patch["subject"] = request.subject or ""
        if "body_html" in fields or "body_text" in fields:
            if request.body_html is not None:
                patch["body"] = {"contentType": "HTML", "content": request.body_html}
            else:
                patch["body"] = {
                    "contentType": "Text",
                    "content": request.body_text or "",
                }
        if "to" in fields:
            patch["toRecipients"] = _recipients(request.to or [])
        if "cc" in fields:
            patch["ccRecipients"] = _recipients(request.cc or [])
        if "bcc" in fields:
            patch["bccRecipients"] = _recipients(request.bcc or [])
        if "importance" in fields and request.importance is not None:
            patch["importance"] = request.importance
        if "categories" in fields:
            patch["categories"] = request.categories or []
        if not patch:
            raise InvalidArgumentError(
                "Provide at least one non-empty draft field to update."
            )
        updated = self.graph.patch(
            f"/me/messages/{quote_segment(request.message_id)}", json=patch
        )
        return {
            "updated_fields": sorted(patch),
            "draft": message_summary(updated, self.settings),
        }
