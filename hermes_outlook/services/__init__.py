from __future__ import annotations

from typing import Any

from ..auth import GraphAuthenticator
from ..config import Settings
from ..graph import GraphClient
from ..models import (
    ArchiveEmailInput,
    AttachmentsInput,
    ConversationInput,
    CopyEmailInput,
    CreateFolderInput,
    DraftInput,
    ListEmailsInput,
    ListFoldersInput,
    MoveEmailInput,
    SearchEmailsInput,
    UpdateDraftInput,
    UpdateEmailInput,
)
from .conversation import ConversationService
from .draft import DraftService
from .folder import FolderService
from .message import MessageService
from .organize import OrganizeService
from .search import SearchService


class OutlookService:
    def __init__(
        self,
        settings: Settings,
        *,
        auth: Any | None = None,
        graph: GraphClient | None = None,
    ) -> None:
        self.settings = settings
        self.auth = auth or GraphAuthenticator(settings)
        self.graph = graph or GraphClient(settings, self.auth)
        self.folders = FolderService(self.graph, settings)
        self.messages = MessageService(self.graph, settings, self.folders)
        self.search_service = SearchService(self.graph, settings, self.folders)
        self.conversations = ConversationService(self.graph, settings)
        self.drafts = DraftService(self.graph, settings, self.messages)
        self.organizer = OrganizeService(self.graph, settings, self.folders)

    def close(self) -> None:
        self.graph.close()

    def list_emails(self, request: ListEmailsInput) -> dict[str, Any]:
        return self.messages.list_emails(request)

    def search_emails(self, request: SearchEmailsInput) -> dict[str, Any]:
        return self.search_service.search(request)

    def get_email(self, message_id: str) -> dict[str, Any]:
        return self.messages.get_email(message_id)

    def get_conversation(self, request: ConversationInput) -> dict[str, Any]:
        return self.conversations.get_conversation(request)

    def list_folders(self, request: ListFoldersInput) -> dict[str, Any]:
        return self.folders.list_folders(request.include_children)

    def list_attachments(self, request: AttachmentsInput) -> dict[str, Any]:
        attachments = self.messages.attachments_metadata(request.message_id)
        return {"attachments": attachments, "count": len(attachments)}

    def create_draft(self, request: DraftInput) -> dict[str, Any]:
        return self.drafts.create(request)

    def update_draft(self, request: UpdateDraftInput) -> dict[str, Any]:
        return self.drafts.update(request)

    def move_email(self, request: MoveEmailInput) -> dict[str, Any]:
        return self.organizer.move(request)

    def copy_email(self, request: CopyEmailInput) -> dict[str, Any]:
        return self.organizer.copy(request)

    def archive_email(self, request: ArchiveEmailInput) -> dict[str, Any]:
        return self.organizer.archive(request)

    def update_email(self, request: UpdateEmailInput) -> dict[str, Any]:
        return self.organizer.update(request)

    def create_folder(self, request: CreateFolderInput) -> dict[str, Any]:
        return self.organizer.create_folder(request)


__all__ = [
    "OutlookService",
    "ConversationService",
    "DraftService",
    "FolderService",
    "MessageService",
    "OrganizeService",
    "SearchService",
]
