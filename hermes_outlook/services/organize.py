from __future__ import annotations

from typing import Any

from ..config import Settings
from ..errors import InvalidArgumentError
from ..graph import GraphClient, quote_segment
from ..models import (
    ArchiveEmailInput,
    CopyEmailInput,
    CreateFolderInput,
    MoveEmailInput,
    UpdateEmailInput,
)
from .base import message_summary
from .folder import FolderService


class OrganizeService:
    def __init__(
        self, graph: GraphClient, settings: Settings, folders: FolderService
    ) -> None:
        self.graph = graph
        self.settings = settings
        self.folders = folders

    def _move(self, message_id: str, destination: str, action: str) -> dict[str, Any]:
        folder = self.folders.resolve(destination)
        result = self.graph.post(
            f"/me/messages/{quote_segment(message_id)}/{action}",
            json={"destinationId": folder.folder_id},
        )
        return {
            "action": action,
            "message": message_summary(result, self.settings),
            "destination_folder": folder.display_name or folder.folder_id,
            "well_known_folder": folder.well_known,
        }

    def move(self, request: MoveEmailInput) -> dict[str, Any]:
        return self._move(
            request.message_id, request.destination_folder, "move"
        )

    def copy(self, request: CopyEmailInput) -> dict[str, Any]:
        return self._move(
            request.message_id, request.destination_folder, "copy"
        )

    def archive(self, request: ArchiveEmailInput) -> dict[str, Any]:
        return self._move(request.message_id, "archive", "move")

    def update(self, request: UpdateEmailInput) -> dict[str, Any]:
        patch: dict[str, Any] = {}
        fields = request.model_fields_set
        if "is_read" in fields and request.is_read is not None:
            patch["isRead"] = request.is_read
        if "flag_status" in fields and request.flag_status is not None:
            patch["flag"] = {"flagStatus": request.flag_status}
        if "categories" in fields:
            patch["categories"] = request.categories or []
        if not patch:
            raise InvalidArgumentError(
                "Provide is_read, flag_status, or categories to update."
            )
        updated = self.graph.patch(
            f"/me/messages/{quote_segment(request.message_id)}", json=patch
        )
        return {
            "updated_fields": sorted(patch),
            "message": message_summary(updated, self.settings),
        }

    def create_folder(self, request: CreateFolderInput) -> dict[str, Any]:
        return self.folders.create(request.display_name, request.parent_folder)
