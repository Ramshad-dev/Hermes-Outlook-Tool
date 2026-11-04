# Installation Log — Hermes Outlook Tool

Steps performed to install and enable this plugin on the local machine.

## Environment

- OS: Windows (DPAPI required for the token cache)
- Hermes Agent: v0.22.5 (satisfies `requires_hermes: ">=0.21"`)
- Hermes config: `C:\Users\user\AppData\Local\hermes\config.yaml`
- Hermes home: `C:\Users\user\AppData\Local\hermes`
- Plugin checkout: `C:\vajra\hermes-outlook-tool`

## 1. Link the plugin into the Hermes plugins directory

Hermes loads directory plugins from `$HERMES_HOME/plugins/<name>`, so the checkout
is exposed there with a directory junction:

```powershell
$hermesHome = Split-Path (hermes config path)   # C:\Users\user\AppData\Local\hermes
$plugins = Join-Path $hermesHome "plugins"
New-Item -ItemType Directory -Force $plugins | Out-Null
New-Item -ItemType Junction -Path (Join-Path $plugins "outlook") -Target (Get-Location)
```

Result: `C:\Users\user\AppData\Local\hermes\outlook` → `C:\vajra\hermes-outlook-tool`
(actual junction path: `C:\Users\user\AppData\Local\hermes\plugins\outlook`).

## 2. Enable the plugin

```powershell
hermes plugins list
hermes plugins enable outlook --no-allow-tool-override
```

`--no-allow-tool-override` keeps the Outlook tools from shadowing Hermes built-ins.
Output: `Plugin 'outlook' is already enabled.` — 13 tools, 0 hooks registered.

## 3. Validate and install dependencies

```powershell
hermes plugins doctor . --ci
hermes plugins validate . --install-deps
```

Both passed:

- manifest parses (`outlook 0.1.0`, standalone, kind `standalone`)
- import and `register()` succeed in isolation
- 13 declared tools match registrations, no built-in collisions, security scan safe
- Python dependencies from `pyproject.toml` (`httpx`, `msal`, `msal-extensions`, `pydantic`) installed
- Validation passed

## 4. Configure Microsoft Graph credentials

The plugin reads config from the Hermes environment file
(`hermes config env-path` → `C:\Users\user\AppData\Local\hermes\.env`).
Appended:

```dotenv
# Hermes Outlook Tool (Microsoft Graph)
GRAPH_CLIENT_ID=7fb28125-4a7b-4332-a083-b664303cd2c1
GRAPH_TENANT_ID=consumers
GRAPH_REDIRECT_URI=http://localhost:8400
```

These values match the repo's local `.env`. The Entra app registration must have a
**Mobile and desktop applications** redirect URI `http://localhost:8400` and delegated
Graph permissions `Mail.ReadWrite` + `offline_access`.

> Note: appending to this file requires a direct write
> (`[System.IO.File]::AppendAllText`); `Add-Content` fails with
> "Stream was not readable" in this shell.

## 5. Verify

```powershell
hermes outlook status
```

Output (before sign-in):

```json
{
  "authenticated": false,
  "account": null,
  "cache": "C:\\Users\\user\\AppData\\Local\\hermes\\outlook\\token_cache.bin"
}
```

`GRAPH_CLIENT_ID` is picked up and the DPAPI token-cache path is resolved.

## 6. Remaining manual step

Interactive browser sign-in (not run during install):

```powershell
hermes outlook login     # browser PKCE sign-in
hermes outlook test      # verifies Inbox access via Graph
hermes outlook logout    # optional: clear cached account
```

## Uninstall

```powershell
hermes plugins disable outlook
Remove-Item (Join-Path $env:LOCALAPPDATA "hermes\plugins\outlook")   # removes the junction only
```
