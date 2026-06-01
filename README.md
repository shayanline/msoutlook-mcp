# msoutlook-mcp

MCP server for Microsoft Outlook Web — no app registration required.

Uses your existing Outlook Web session (the same way [msteams-mcp](https://github.com/m0nkmaster/msteams-mcp) uses the Teams web session). Opens a browser once for login, then caches tokens and refreshes them automatically.

## How it works

Microsoft Outlook Web (OWA) uses MSAL to store OAuth tokens in `localStorage`. This server:

1. Opens a browser to `outlook.office.com` via Playwright
2. Extracts the MSAL token cache from `localStorage` — using OWA's own first-party client ID (`9199bf20-a13f-4107-85dc-02114787ef48`)
3. Caches the access token, refresh token, and session state in `~/.msoutlook-mcp-server/` (AES-256-GCM encrypted)
4. Refreshes tokens automatically using the refresh token (HTTP, no browser) or headless browser as fallback

No Azure app registration. No admin consent. No client secrets. Your access is limited to what your account can already do.

## Quick start

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "msoutlook-mcp@latest"]
    }
  }
}
```

Then run `outlook_login` from your MCP client to open the browser and authenticate.

## Tools

### Auth

| Tool | Description |
|------|-------------|
| `outlook_login` | Open browser to sign in to Outlook Web |
| `outlook_status` | Check authentication status and token validity |
| `outlook_logout` | Clear saved session and tokens |

### Email

| Tool | Description |
|------|-------------|
| `outlook_list_emails` | List emails from any folder (Inbox by default) |
| `outlook_get_email` | Read full email content by ID |
| `outlook_get_unread` | Get unread emails from Inbox |
| `outlook_search_emails` | Search emails by keyword |
| `outlook_send_email` | Send an email |
| `outlook_create_draft` | Create a draft without sending |
| `outlook_send_draft` | Send a previously created draft |
| `outlook_reply` | Reply to an email (or reply all) |
| `outlook_forward` | Forward an email |
| `outlook_mark_read` | Mark email as read or unread |
| `outlook_flag` | Flag or unflag an email |
| `outlook_move_email` | Move email to a different folder |
| `outlook_delete_email` | Delete an email |
| `outlook_list_folders` | List all mail folders with unread counts |

### Calendar

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List calendar events in a date range |
| `outlook_get_event` | Get full event details including attendees |
| `outlook_create_event` | Create a meeting or appointment |
| `outlook_update_event` | Update an existing event |
| `outlook_delete_event` | Delete an event |
| `outlook_respond_to_event` | Accept, decline, or tentatively accept an invite |
| `outlook_search_events` | Search events by keyword |
| `outlook_list_calendars` | List all calendars |

### Contacts & People

| Tool | Description |
|------|-------------|
| `outlook_list_contacts` | List/search contacts |
| `outlook_get_contact` | Get contact details by ID |
| `outlook_create_contact` | Create a new contact |
| `outlook_delete_contact` | Delete a contact |
| `outlook_search_people` | Search the organisation directory |

## Session storage

Session files are stored encrypted in `~/.msoutlook-mcp-server/`:

- `session-state.json` — Playwright browser session (cookies + localStorage)
- `token-cache.json` — Extracted and cached tokens
- `browser-profile/` — Persistent browser profile for headless refresh

If your session expires, run `outlook_login` again.

## Token refresh

Tokens are refreshed automatically:

1. **HTTP refresh** (fast, no browser) — uses the cached refresh token with OWA's client ID
2. **Headless browser refresh** — fallback if HTTP refresh fails; opens a headless Edge window to silently re-acquire tokens from the saved browser session

## Requirements

- Node.js 20+
- A Chromium-based browser: Edge or Chrome (auto-detected from system default)
- A Microsoft 365 work/school account or personal Microsoft account

## Platform support

| Feature | macOS | Linux | Windows |
|---------|-------|-------|---------|
| Browser auto-detection | System default (Edge/Chrome) | Chrome fallback | Edge (pre-installed) |
| SSO cookie import | ✅ Chrome + Edge via Keychain | ✅ Chrome + Edge via libsecret / `"peanuts"` fallback | ✅ Edge via DPAPI (PowerShell) |
| Windows Chrome 127+ cookies | — | — | ⚠️ App-Bound Encryption — not supported. Use Edge instead. |
| Headed browser fallback | ✅ | ✅ | ✅ |

Cookie import is a best-effort optimisation. If it cannot run (e.g. no matching browser installed, Keychain denied), the MCP falls back to opening a headed browser where you sign in once manually — the session then persists.

## Security notes

- Uses the same auth as the Outlook web client — your access is limited to what your account can do
- Tokens are encrypted at rest (AES-256-GCM with a machine-derived key)
- Uses undocumented internal APIs — Microsoft may change these without notice
- Always confirm email content with the user before sending

## Environment variables

| Variable | Description |
|----------|-------------|
| `MSOUTLOOK_DEBUG=true` | Enable debug logging to stderr |
| `MSOUTLOOK_BROWSER=chrome` | Force a specific browser: `chrome` or `msedge`. If unset, uses the macOS system default browser; falls back to Chrome on macOS/Linux and Edge on Windows. |
| `MSOUTLOOK_CHROME_PROFILE` | Pin a specific Chrome profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default`. |
| `MSOUTLOOK_EDGE_PROFILE` | Pin a specific Edge profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default`. |
