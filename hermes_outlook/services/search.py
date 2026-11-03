from __future__ import annotations

from typing import Any

from ..config import Settings
from ..graph import GraphClient, quote_segment
from ..models import SearchEmailsInput
from .base import (
    MESSAGE_SUMMARY_SELECT,
    build_search,
    message_summary,
    next_link_from,
)
from .folder import FolderService


class SearchService:
    def __init__(
        self, graph: GraphClient, settings: Settings, folders: FolderService
    ) -> None:
        self.graph = graph
        self.settings = settings
        self.folders = folders

    def search(self, request: SearchEmailsInput) -> dict[str, Any]:
        if request.next_link:
            payload = self.graph.get(request.next_link)
        else:
            limit = min(request.limit, self.settings.max_page_size)
            params: dict[str, Any] = {
                "$top": limit,
                "$select": MESSAGE_SUMMARY_SELECT,
                "$search": build_search(
                    request.query,
                    request.sender,
                    request.subject,
                    request.received_after,
                    request.received_before,
                ),
            }
            if request.folder:
                folder = self.folders.resolve(request.folder)
                path = f"/me/mailFolders/{quote_segment(folder.folder_id)}/messages"
            else:
                path = "/me/messages"
            payload = self.graph.get(path, params=params)
        values = payload.get("value")
        if not isinstance(values, list):
            values = []
        emails = [
            message_summary(value, self.settings)
            for value in values
            if isinstance(value, dict)
        ]
        return {
            "emails": emails,
            "count": len(emails),
            "next_link": next_link_from(payload, self.settings),
        }
