from __future__ import annotations

from typing import Any

from ..config import Settings
from ..graph import GraphClient, quote_segment
from ..models import ListEmailsInput
from .base import (
    MESSAGE_DETAIL_SELECT,
    MESSAGE_SUMMARY_SELECT,
    message_detail,
    message_summary,
    next_link_from,
)
from .folder import FolderService

_MAX_ATTACHMENTS = 200


class MessageService:
    def __init__(
        self, graph: GraphClient, settings: Settings, folders: FolderService
    ) -> None:
        self.graph = graph
        self.settings = settings
        self.folders = folders

    def list_emails(self, request: ListEmailsInput) -> dict[str, Any]:
        if request.next_link:
            payload = self.graph.get(request.next_link)
        else:
            folder = self.folders.resolve(request.folder)
            limit = min(request.limit, self.settings.max_page_size)
            params: dict[str, Any] = {
                "$top": limit,
                "$select": MESSAGE_SUMMARY_SELECT,
                "$orderby": "receivedDateTime desc",
            }
            filters = []
            if request.unread_only:
                filters.append("isRead eq false")
            if request.received_after:
                filters.append(
                    f"receivedDateTime ge {request.received_after.strftime('%Y-%m-%dT%H:%M:%SZ')}"
                )
            if request.received_before:
                filters.append(
                    f"receivedDateTime le {request.received_before.strftime('%Y-%m-%dT%H:%M:%SZ')}"
                )
            if filters:
                params["$filter"] = " and ".join(filters)
            path = f"/me/mailFolders/{quote_segment(folder.folder_id)}/messages"
            payload = self.graph.get(path, params=params)
        values = payload.get("value")
        if not isinstance(values, list):
            values = []
        messages = [
            message_summary(value, self.settings)
            for value in values
            if isinstance(value, dict)
        ]
        return {
            "emails": messages,
            "count": len(messages),
            "next_link": next_link_from(payload, self.settings),
        }

    def get_raw(
        self,
        message_id: str,
        *,
        select: str = MESSAGE_DETAIL_SELECT,
    ) -> dict[str, Any]:
        path = f"/me/messages/{quote_segment(message_id)}"
        return self.graph.get(path, params={"$select": select})

    def get_email(self, message_id: str) -> dict[str, Any]:
        value = self.get_raw(message_id)
        attachments: list[dict[str, Any]] = []
        if value.get("hasAttachments"):
            attachments = self.attachments_metadata(message_id)
        return message_detail(value, self.settings, attachments=attachments)

    def attachments_metadata(self, message_id: str) -> list[dict[str, Any]]:
        path = f"/me/messages/{quote_segment(message_id)}/attachments"
        results: list[dict[str, Any]] = []
        current: str | None = path
        first = True
        while current and len(results) < _MAX_ATTACHMENTS:
            payload = self.graph.get(
                current,
                params={
                    "$top": 100,
                    "$select": "id,name,contentType,size,isInline",
                }
                if first
                else None,
            )
            first = False
            values = payload.get("value")
            if not isinstance(values, list):
                break
            for value in values:
                if not isinstance(value, dict) or len(results) >= _MAX_ATTACHMENTS:
                    continue
                odata_type = str(value.get("@odata.type") or "")
                results.append(
                    {
                        "attachment_id": value.get("id"),
                        "name": value.get("name"),
                        "content_type": value.get("contentType"),
                        "size": value.get("size"),
                        "is_inline": value.get("isInline"),
                        "attachment_type": odata_type.rsplit(".", 1)[-1] or None,
                    }
                )
            current = next_link_from(payload, self.settings)
        return results
