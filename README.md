# Hermes Outlook Tool

A standalone [Hermes Agent](https://github.com/NousResearch/hermes-agent) plugin for reading, searching, drafting, and organizing Outlook mail through Microsoft Graph.

The plugin **never sends or deletes mail**. It uses delegated browser sign-in (authorization code + PKCE) and stores the MSAL token cache with Windows DPAPI.

## Capabilities

| Tool | Purpose |
| --- | --- |
| `outlook_list_emails` | List messages in a folder with optional unread/date filters |
| `outlook_search_emails` | Search with Microsoft Graph KQL (`from:`, dates, free text) |
| `outlook_get_email` | Read one message, including safe text body and attachment metadata |
| `outlook_get_conversation` | Read a conversation thread |
| `outlook_list_folders` | List mail folders as a tree |
| `outlook_list_attachments` | List attachment metadata without downloading content |
| `outlook_create_draft` | Create a draft (no attachments) |
| `outlook_update_draft` | Update fields on an existing draft |
| `outlook_move_email` | Move a message to a folder |
| `outlook_copy_email` | Copy a message to a folder |
| `outlook_archive_email` | Move a message to the Archive folder |
| `outlook_update_email` | Mark read/unread or set flag status |
| `outlook_create_folder` | Create a mail folder |

Deliberately unsupported: sending, deleting, attachment downloads, and draft attachments.

## Requirements

- Windows (DPAPI token-cache encryption)
- Hermes Agent 0.21 or newer
- Python 3.10+
- A Microsoft Entra app registration with delegated Microsoft Graph permission `Mail.ReadWrite` and `offline_access`

## Entra app registration

1. In Microsoft Entra admin center, create an **App registration**.
2. Add a **Mobile and desktop applications** platform with redirect URI `http://localhost:8400`.
3. Under **API permissions**, add delegated Microsoft Graph permissions:
   - `Mail.ReadWrite`
   - `offline_access`
4. Grant admin consent if your tenant requires it.
5. Copy the **Application (client) ID** and **Directory (tenant) ID**.

The plugin refuses `Mail.Send`, `Mail.ReadWrite.All`, `Mail.Read.All`, non-localhost redirect URIs, and non-Graph base URLs.

## Configuration

Set these environment variables in the Hermes environment. Find the file with `hermes config env-path`:

```dotenv
GRAPH_CLIENT_ID=00000000-0000-0000-0000-000000000000
GRAPH_TENANT_ID=organizations
GRAPH_REDIRECT_URI=http://localhost:8400
# OUTLOOK_TOKEN_CACHE=C:\Users\you\.hermes\outlook\token_cache.bin
```

- `GRAPH_CLIENT_ID` is required.
- `GRAPH_TENANT_ID` defaults to `organizations` and accepts a tenant ID, domain, `common`, `organizations`, or `consumers`.
- `GRAPH_REDIRECT_URI` defaults to `http://localhost:8400` and must be an HTTP localhost URI.
- `OUTLOOK_TOKEN_CACHE` overrides the default cache location under `$HERMES_HOME/outlook/`.

Optional limits: `GRAPH_SCOPES`, `GRAPH_BASE_URL`, `OUTLOOK_MAILBOX_USER`, `OUTLOOK_PAGE_SIZE`, `OUTLOOK_MAX_PAGE_SIZE`, and retry/body-size settings. See `hermes_outlook/config.py`.

## Install for local development

Hermes loads directory plugins from `$HERMES_HOME/plugins/<name>` (or one category level below it). Find the plugin directory from the active Hermes home, then link this checkout into it:

```powershell
$hermesHome = Split-Path (hermes config path)
$plugins = Join-Path $hermesHome "plugins"
New-Item -ItemType Directory -Force $plugins | Out-Null
New-Item -ItemType Junction -Path (Join-Path $plugins "outlook") -Target (Get-Location)
hermes plugins list
hermes plugins enable outlook --no-allow-tool-override
```

Validate the plugin with the real Hermes runtime contracts:

```powershell
hermes plugins doctor . --ci
hermes plugins validate . --install-deps
```

## Sign in and test

```powershell
hermes outlook config
hermes outlook login
hermes outlook status
hermes outlook test
```

`login` opens a browser for delegated PKCE sign-in. `test` calls Graph for the Inbox folder metadata to confirm mailbox access.

To sign out and remove the cached account:

```powershell
hermes outlook logout
```

## Security model

- Delegated OAuth only; no client secrets.
- MSAL cache is encrypted with Windows DPAPI and fails closed if secure storage is unavailable.
- Graph requests are restricted to the configured Graph host.
- `@odata.nextLink` values must stay on the configured Graph host and under `/me/`.
- Graph redirects are rejected; 401 triggers one token refresh; 429/5xx are retried with backoff.
- HTML bodies are reduced to text, with `<script>` and `<style>` removed.
- Message content is labeled `content_trust: "untrusted_external"` and includes a security notice for the model.
- No send, delete, or destructive move tools are registered.

## Development

```powershell
uv venv .venv
uv pip install --python .venv\Scripts\python.exe -e ".[dev]"
.venv\Scripts\python.exe -m pytest
```

The test suite uses `httpx.MockTransport`; it does not call Microsoft Graph.

## Limitations

- Mailbox access is delegated to the signed-in user; shared mailboxes require a separate design.
- Search uses `$search` KQL, so date filters are date-level.
- Draft attachments are not supported.
- Attachment content is not downloaded; only metadata is listed.
- Folder names that are duplicated require an explicit folder ID.
