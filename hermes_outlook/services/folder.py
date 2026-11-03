from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from ..config import Settings
from ..errors import InvalidArgumentError, ResourceNotFoundError
from ..graph import GraphClient, quote_segment
from .base import FOLDER_SELECT, folder_summary, next_link_from

_MAX_FOLDERS = 500
_MAX_DEPTH = 10
_ID_PATTERN = re.compile(r"^[A-Za-z0-9_\-=+/]{40,}$")

_ALIASES = {
    "archive": "archive",
    "clutter": "clutter",
    "conflicts": "conflicts",
    "conversation history": "conversationhistory",
    "conversationhistory": "conversationhistory",
    "deleted items": "deleteditems",
    "deleteditems": "deleteditems",
    "drafts": "drafts",
    "inbox": "inbox",
    "junk email": "junkemail",
    "junkemail": "junkemail",
    "msgfolderroot": "msgfolderroot",
    "outbox": "outbox",
    "recoverable items deletions": "recoverableitemsdeletions",
    "scheduled": "scheduled",
    "search folders": "searchfolders",
    "sent items": "sentitems",
    "sentitems": "sentitems",
    "server failures": "serverfailures",
    "sync issues": "syncissues",
    "temporary": "temporary",
}


@dataclass(frozen=True)
class FolderRef:
    folder_id: str
    display_name: str | None
    well_known: bool


def _looks_like_id(value: str) -> bool:
    return bool(_ID_PATTERN.fullmatch(value))


def _alias(value: str) -> str | None:
    key = re.sub(r"[\s_-]+", " ", value.strip().casefold())
    return _ALIASES.get(key)


class FolderService:
    def __init__(self, graph: GraphClient, settings: Settings) -> None:
        self.graph = graph
        self.settings = settings
        self._all_folders: list[dict[str, Any]] | None = None

    def _fetch_collection(
        self, path: str, *, limit: int = _MAX_FOLDERS
    ) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        current: str | None = path
        first = True
        while current and len(results) < limit:
            payload = self.graph.get(
                current,
                params={"$top": 100, "$select": FOLDER_SELECT} if first else None,
            )
            first = False
            values = payload.get("value")
            if isinstance(values, list):
                for value in values:
                    if isinstance(value, dict) and len(results) < limit:
                        results.append(value)
            current = next_link_from(payload, self.settings)
        return results

    def _load_all(self) -> list[dict[str, Any]]:
        if self._all_folders is not None:
            return self._all_folders
        top_level = self._fetch_collection("/me/mailFolders")
        all_folders = list(top_level)
        seen = {
            value.get("id")
            for value in all_folders
            if isinstance(value.get("id"), str)
        }
        queue = [
            (value.get("id"), 1)
            for value in top_level
            if value.get("id") and (value.get("childFolderCount") or 0) > 0
        ]
        while queue and len(all_folders) < _MAX_FOLDERS:
            parent_id, depth = queue.pop(0)
            if depth > _MAX_DEPTH:
                continue
            children = self._fetch_collection(
                f"/me/mailFolders/{quote_segment(str(parent_id))}/childFolders",
                limit=_MAX_FOLDERS - len(all_folders),
            )
            for child in children:
                child_id = child.get("id")
                if not child_id or child_id in seen:
                    continue
                seen.add(child_id)
                all_folders.append(child)
                if (child.get("childFolderCount") or 0) > 0:
                    queue.append((child_id, depth + 1))
        self._all_folders = all_folders
        return all_folders

    def list_folders(self, include_children: bool = True) -> dict[str, Any]:
        if include_children:
            folders = self._load_all()
        else:
            folders = self._fetch_collection("/me/mailFolders")
        children: dict[str | None, list[dict[str, Any]]] = defaultdict(list)
        ids = {
            folder.get("id")
            for folder in folders
            if isinstance(folder.get("id"), str)
        }
        for folder in folders:
            parent = folder.get("parentFolderId")
            if not include_children or parent not in ids:
                parent = None
            children[parent].append(folder)

        count = 0

        def build(parent: str | None, depth: int) -> list[dict[str, Any]]:
            nonlocal count
            if depth > _MAX_DEPTH:
                return []
            nodes = []
            for folder in children.get(parent, []):
                count += 1
                node = folder_summary(folder)
                node["children"] = (
                    build(str(folder.get("id")), depth + 1)
                    if include_children and folder.get("id")
                    else []
                )
                nodes.append(node)
            return nodes

        tree = build(None, 0)
        return {"folders": tree, "count": count}

    def resolve(self, value: str) -> FolderRef:
        raw = value.strip()
        if _looks_like_id(raw):
            return FolderRef(folder_id=raw, display_name=None, well_known=False)
        alias = _alias(raw)
        if alias:
            return FolderRef(folder_id=alias, display_name=raw, well_known=True)
        matches = [
            folder
            for folder in self._load_all()
            if str(folder.get("displayName") or "").casefold() == raw.casefold()
        ]
        if len(matches) == 1:
            folder = matches[0]
            return FolderRef(
                folder_id=str(folder.get("id")),
                display_name=folder.get("displayName"),
                well_known=False,
            )
        if len(matches) > 1:
            candidates = [
                {"folder_id": folder.get("id"), "display_name": folder.get("displayName")}
                for folder in matches
            ]
            raise InvalidArgumentError(
                f"Folder name {raw!r} is ambiguous; use a folder_id instead.",
                details={"candidates": candidates},
            )
        raise ResourceNotFoundError(f"Outlook folder {raw!r} was not found.")

    def create(self, display_name: str, parent_folder: str | None = None) -> dict[str, Any]:
        payload = {"displayName": display_name}
        if parent_folder:
            parent = self.resolve(parent_folder)
            path = f"/me/mailFolders/{quote_segment(parent.folder_id)}/childFolders"
        else:
            path = "/me/mailFolders"
        created = self.graph.post(path, json=payload)
        self._all_folders = None
        return folder_summary(created)
