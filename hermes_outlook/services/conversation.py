from __future__ import annotations

from typing import Any

from ..config import Settings
from ..graph import GraphClient
from ..models import ConversationInput
from .base import (
    MESSAGE_DETAIL_SELECT,
    MESSAGE_SUMMARY_SELECT,
    message_detail,
    message_summary,
    next_link_from,
    odata_quote,
)


class ConversationService:
    def __init__(self, graph: GraphClient, settings: Settings) -> None:
        self.graph = graph
        self.settings = settings

    def get_conversation(self, request: ConversationInput) -> dict[str, Any]:
        if request.next_link:
            payload = self.graph.get(request.next_link)
        else:
            limit = min(request.limit, self.settings.max_page_size)
            payload = self.graph.get(
                "/me/messages",
                params={
                    "$top": limit,
                    "$select": MESSAGE_DETAIL_SELECT
                    if request.include_body
                    else MESSAGE_SUMMARY_SELECT,
                    "$filter": (
                        f"conversationId eq '{odata_quote(request.conversation_id)}'"
                    ),
                },
            )
        values = payload.get("value")
        if not isinstance(values, list):
            values = []
        if request.include_body:
            emails = [
                message_detail(value, self.settings, attachments_omitted=True)
                for value in values
                if isinstance(value, dict)
            ]
        else:
            emails = [
                message_summary(value, self.settings)
                for value in values
                if isinstance(value, dict)
            ]
        emails.sort(key=lambda email: email.get("received_at") or "")
        return {
            "conversation_id": request.conversation_id,
            "emails": emails,
            "count": len(emails),
            "next_link": next_link_from(payload, self.settings),
        }
