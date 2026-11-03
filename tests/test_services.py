from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import httpx
import pytest

from hermes_outlook.config import Settings
from hermes_outlook.errors import NotADraftError, ResourceNotFoundError
from hermes_outlook.models import (
    ArchiveEmailInput,
    ConversationInput,
    CreateFolderInput,
    DraftInput,
    ListEmailsInput,
    MoveEmailInput,
    SearchEmailsInput,
    UpdateDraftInput,
    UpdateEmailInput,
)
from hermes_outlook.services import OutlookService
from hermes_outlook.services.folder import FolderService

MESSAGE = {
    "id": "msg1",
    "conversationId": "conv1",
    "subject": "Quarterly report",
    "from": {"emailAddress": {"name": "Boss", "address": "boss@example.com"}},
    "sender": {"emailAddress": {"name": "Boss", "address": "boss@example.com"}},
    "receivedDateTime": "2026-01-02T03:04:05Z",
    "importance": "normal",
    "isRead": False,
    "isDraft": False,
    "hasAttachments": True,
    "bodyPreview": "Please review",
    "parentFolderId": "inbox-id",
    "webLink": "https://outlook.office.com/mail/id/msg1",
    "categories": [],
    "flag": {"flagStatus": "notFlagged"},
}


def service_with(settings: Settings, handler: Any, client_factory: Any) -> OutlookService:
    graph = client_factory(settings, httpx.MockTransport(handler))
    return OutlookService(settings, auth=object(), graph=graph)


def test_folder_tree_and_resolution(settings: Settings, client_factory: Any) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/me/mailFolders":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "inbox-id",
                            "displayName": "Inbox",
                            "parentFolderId": None,
                            "childFolderCount": 1,
                            "totalItemCount": 10,
                            "unreadItemCount": 2,
                        },
                        {
                            "id": "archive-id",
                            "displayName": "Archive",
                            "parentFolderId": None,
                            "childFolderCount": 0,
                            "totalItemCount": 3,
                            "unreadItemCount": 0,
                        },
                    ]
                },
            )
        if request.url.path == "/v1.0/me/mailFolders/inbox-id/childFolders":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "projects-id",
                            "displayName": "Projects",
                            "parentFolderId": "inbox-id",
                            "childFolderCount": 0,
                            "totalItemCount": 1,
                            "unreadItemCount": 1,
                        }
                    ]
                },
            )
        return httpx.Response(404, json={"error": {"code": "missing"}})

    folders = FolderService(
        client_factory(settings, httpx.MockTransport(handler)), settings
    )
    tree = folders.list_folders(True)
    assert tree["count"] == 3
    inbox = next(item for item in tree["folders"] if item["folder_id"] == "inbox-id")
    assert inbox["children"][0]["display_name"] == "Projects"
    assert folders.resolve("Sent Items").folder_id == "sentitems"
    assert folders.resolve("Projects").folder_id == "projects-id"
    with pytest.raises(ResourceNotFoundError):
        folders.resolve("Missing")


def test_list_emails_builds_filters(settings: Settings, client_factory: Any) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return httpx.Response(
            200,
            json={
                "value": [MESSAGE],
                "@odata.nextLink": (
                    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox-id/messages"
                    "?$skiptoken=abc"
                ),
            },
        )

    service = service_with(settings, handler, client_factory)
    result = service.list_emails(
        ListEmailsInput(
            folder="Inbox",
            unread_only=True,
            received_after=datetime(2026, 1, 1, tzinfo=timezone.utc),
        )
    )
    assert result["count"] == 1
    assert result["emails"][0]["message_id"] == "msg1"
    assert result["next_link"].endswith("$skiptoken=abc")
    params = seen[0].url.params
    assert params["$filter"] == "isRead eq false and receivedDateTime ge 2026-01-01T00:00:00Z"
    assert params["$orderby"] == "receivedDateTime desc"


def test_get_email_sanitizes_html_and_lists_attachments(
    settings: Settings, client_factory: Any
) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/me/messages/msg1":
            payload = dict(MESSAGE)
            payload.update(
                {
                    "toRecipients": [
                        {"emailAddress": {"name": "Me", "address": "me@example.com"}}
                    ],
                    "body": {
                        "contentType": "html",
                        "content": "<p>Hello</p><script>steal()</script><p>Review</p>",
                    },
                }
            )
            return httpx.Response(200, json=payload)
        if request.url.path == "/v1.0/me/messages/msg1/attachments":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "att1",
                            "name": "report.pdf",
                            "contentType": "application/pdf",
                            "size": 1234,
                            "isInline": False,
                        }
                    ]
                },
            )
        return httpx.Response(404, json={"error": {"code": "missing"}})

    service = service_with(settings, handler, client_factory)
    email = service.get_email("msg1")
    assert "steal" not in email["body_text"]
    assert "Hello" in email["body_text"]
    assert email["attachments"][0]["name"] == "report.pdf"
    assert email["attachments"][0]["attachment_type"] == "fileAttachment"
    assert email["content_trust"] == "untrusted_external"


def test_search_uses_kql_not_filter(settings: Settings, client_factory: Any) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return httpx.Response(200, json={"value": [MESSAGE]})

    service = service_with(settings, handler, client_factory)
    result = service.search_emails(
        SearchEmailsInput(
            query="quarterly report",
            sender="boss@example.com",
            received_after=datetime(2026, 1, 1, tzinfo=timezone.utc),
            received_before=datetime(2026, 2, 1, tzinfo=timezone.utc),
        )
    )
    assert result["count"] == 1
    params = seen[0].url.params
    assert params["$search"] == (
        '"quarterly report" AND "from:boss@example.com" AND '
        '"received>=2026-01-01" AND "received<=2026-02-01"'
    )
    assert params.get("$filter") is None


def test_conversation_escapes_odata_quotes(
    settings: Settings, client_factory: Any
) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return httpx.Response(200, json={"value": [MESSAGE]})

    service = service_with(settings, handler, client_factory)
    service.get_conversation(ConversationInput(conversation_id="conv'1"))
    assert seen[0].url.params["$filter"] == "conversationId eq 'conv''1'"
    assert seen[0].url.params["$orderby"] == "receivedDateTime asc"


def test_draft_create_and_update(settings: Settings, client_factory: Any) -> None:
    bodies: list[dict[str, Any]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        if request.method == "POST" and request.url.path == "/v1.0/me/messages":
            bodies.append(__import__("json").loads(request.content))
            return httpx.Response(
                201,
                json={"id": "draft1", "webLink": "https://outlook.office.com/draft1"},
            )
        if request.method == "GET" and request.url.path == "/v1.0/me/messages/draft1":
            return httpx.Response(200, json={"id": "draft1", "isDraft": True})
        if request.method == "PATCH" and request.url.path == "/v1.0/me/messages/draft1":
            bodies.append(__import__("json").loads(request.content))
            payload = dict(MESSAGE)
            payload.update({"id": "draft1", "isDraft": True})
            return httpx.Response(200, json=payload)
        return httpx.Response(404, json={"error": {"code": "missing"}})

    service = service_with(settings, handler, client_factory)
    created = service.create_draft(
        DraftInput(
            subject="Hello",
            body_text="Body",
            to=["person@example.com"],
            cc=["copy@example.com"],
        )
    )
    assert created["message_id"] == "draft1"
    assert bodies[0]["body"] == {"contentType": "Text", "content": "Body"}
    assert bodies[0]["toRecipients"][0]["emailAddress"]["address"] == "person@example.com"

    updated = service.update_draft(
        UpdateDraftInput(message_id="draft1", subject="New subject", to=[])
    )
    assert updated["updated_fields"] == ["subject", "toRecipients"]
    assert bodies[1]["toRecipients"] == []


def test_update_rejects_non_draft(settings: Settings, client_factory: Any) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={"id": "msg1", "isDraft": False})

    service = service_with(settings, handler, client_factory)
    with pytest.raises(NotADraftError):
        service.update_draft(UpdateDraftInput(message_id="msg1", subject="x"))


def test_move_archive_and_update_state(
    settings: Settings, client_factory: Any
) -> None:
    posts: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/me/mailFolders":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "projects-id",
                            "displayName": "Projects",
                            "parentFolderId": None,
                            "childFolderCount": 0,
                            "totalItemCount": 0,
                            "unreadItemCount": 0,
                        }
                    ]
                },
            )
        if request.method == "POST" and request.url.path == "/v1.0/me/messages/msg1/move":
            posts.append(__import__("json").loads(request.content))
            return httpx.Response(200, json=MESSAGE)
        if request.method == "PATCH" and request.url.path == "/v1.0/me/messages/msg1":
            patches.append(__import__("json").loads(request.content))
            payload = dict(MESSAGE)
            payload.update({"isRead": True, "flag": {"flagStatus": "flagged"}})
            return httpx.Response(200, json=payload)
        return httpx.Response(404, json={"error": {"code": "missing"}})

    service = service_with(settings, handler, client_factory)
    moved = service.move_email(
        MoveEmailInput(message_id="msg1", destination_folder="Projects")
    )
    assert moved["message"]["message_id"] == "msg1"
    assert posts[0] == {"destinationId": "projects-id"}
    archived = service.archive_email(ArchiveEmailInput(message_id="msg1"))
    assert archived["well_known_folder"] is True
    assert posts[1] == {"destinationId": "archive"}
    updated = service.update_email(
        UpdateEmailInput(message_id="msg1", is_read=True, flag_status="flagged")
    )
    assert updated["updated_fields"] == ["flag", "isRead"]
    assert patches[0] == {"isRead": True, "flag": {"flagStatus": "flagged"}}


def test_create_folder(settings: Settings, client_factory: Any) -> None:
    seen: list[dict[str, Any]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(__import__("json").loads(request.content))
        return httpx.Response(
            201,
            json={
                "id": "new-id",
                "displayName": "New folder",
                "parentFolderId": None,
                "childFolderCount": 0,
                "totalItemCount": 0,
                "unreadItemCount": 0,
            },
        )

    service = service_with(settings, handler, client_factory)
    folder = service.create_folder(CreateFolderInput(display_name="New folder"))
    assert folder["folder_id"] == "new-id"
    assert seen[0] == {"displayName": "New folder"}
