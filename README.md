# msoutlook-mcp

[![npm version](https://img.shields.io/npm/v/msoutlook-mcp.svg)](https://www.npmjs.com/package/msoutlook-mcp)
[![npm downloads](https://img.shields.io/npm/dm/msoutlook-mcp.svg)](https://www.npmjs.com/package/msoutlook-mcp)
[![node](https://img.shields.io/node/v/msoutlook-mcp.svg)](https://www.npmjs.com/package/msoutlook-mcp)
[![license](https://img.shields.io/npm/l/msoutlook-mcp.svg)](./LICENSE)

MCP server for Microsoft Outlook Web. No app registration required.

Give any MCP client (Claude, Cursor, Devin, ...) read and write access to your Outlook mail, calendar, contacts, and the org directory. It works by reusing your existing Outlook Web session, the same way [@shayanline/msteams-mcp](https://github.com/shayanline/msteams-mcp) reuses the Teams web session: you sign in once in a browser, then tokens are cached and refreshed automatically.

## Why

Most Outlook MCP servers make you register an Azure AD application, grant admin consent, and manage client secrets. This one needs none of that. It reuses Outlook Web's own first party client ID, so your access is exactly what your account already has, and nothing leaves your own machine.

## Highlights

- **Zero setup auth.** No Azure app registration, no admin consent, no client secrets. Sign in once in a browser; tokens are cached encrypted and refreshed silently after that.
- **Over 50 tools** across mail, calendar, contacts, people, the org directory, and your own account settings.
- **Safe by default.** Email bodies are sent as HTML so line breaks always render, and the draft tools let you compose and review before anything is sent.
- **Scheduling and availability.** Read free/busy, find meeting times across attendees, and check who is out of office.
- **Cross platform.** macOS, Linux, and Windows, with automatic browser detection.

## How it works

Microsoft Outlook Web (OWA) uses MSAL to store OAuth tokens in `localStorage`. This server:

1. Opens a browser to `outlook.office.com` via Playwright
2. Extracts the MSAL token cache from `localStorage`, using OWA's own first party client ID (`9199bf20-a13f-4107-85dc-02114787ef48`)
3. Caches the access token, refresh token, and session state in `~/.msoutlook-mcp-server/` (AES-256-GCM encrypted)
4. Refreshes tokens automatically using the refresh token (HTTP, no browser) or a headless browser as fallback

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

The top-level key depends on your client. Most clients (Claude Desktop, Cursor, Windsurf) use `mcpServers` as shown above. VS Code's `.vscode/mcp.json` uses `servers` instead, so drop the block in as:

```json
{
  "servers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "msoutlook-mcp@latest"]
    }
  }
}
```

Then run `outlook_login` from your MCP client. On first use a browser opens so you can sign in; after that, logins are silent and no browser appears. Do not close the window manually, it closes itself once you are signed in.

## Sending mail: behaviour to know

- **HTML by default.** `outlook_send_email`, `outlook_create_draft`, and the reply/forward tools render their body as HTML, so use `<br>`, `<br><br>`, and `<ul><li>` for structure. Plain text is still accepted: its newlines are converted to `<br>` automatically, so a message never arrives as one collapsed block.
- **Review before sending.** Prefer the draft tools (`outlook_create_draft`, `outlook_create_reply_draft`, `outlook_create_forward_draft`) so you can review or edit with `outlook_update_draft` and then send with `outlook_send_draft`. The immediate tools (`outlook_send_email`, `outlook_reply`, `outlook_forward`) send straight away.
- **Presence caveat.** `outlook_get_availability` reports free/busy and out of office, which is what Outlook can read. The live Teams presence dot (Available / Away / DoNotDisturb) needs a Teams token and is not available here.

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
| `outlook_search_emails` | Search emails by keyword, with optional received date range and pagination. Omit the keyword to list everything in a date range |
| `outlook_send_email` | Send an email (HTML body by default; plain text newlines auto convert to `<br>`; supports file `attachments`) |
| `outlook_create_draft` | Create a new draft without sending (review-first flow; supports file `attachments`) |
| `outlook_send_draft` | Send a previously created draft |
| `outlook_update_draft` | Edit a draft's subject, recipients, body, or importance |
| `outlook_reply` | Reply to an email immediately (or reply all) |
| `outlook_create_reply_draft` | Create a reply (or reply all) as a draft to review before sending |
| `outlook_forward` | Forward an email immediately |
| `outlook_create_forward_draft` | Create a forward as a draft to review before sending |
| `outlook_add_attachment` | Attach a local file to an existing draft |
| `outlook_list_attachments` | List attachments on an email |
| `outlook_save_attachment` | Download an attachment to a local file |
| `outlook_get_conversation` | Get every message in a thread by conversation ID |
| `outlook_set_categories` | Set colour categories (labels) on an email |
| `outlook_mark_read` | Mark email as read or unread |
| `outlook_flag` | Flag or unflag an email |
| `outlook_move_email` | Move email to a different folder |
| `outlook_delete_email` | Delete an email |
| `outlook_batch` | Run one bulk action (mark read, mark unread, flag, unflag, move, delete) over many emails at once |
| `outlook_list_folders` | List all mail folders with unread counts |
| `outlook_create_folder` | Create a mail folder (optionally nested) |
| `outlook_rename_folder` | Rename a mail folder |
| `outlook_delete_folder` | Delete a mail folder |

### Calendar

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List calendar events in a date range (expands recurring series into instances) |
| `outlook_get_event` | Get full event details including attendees |
| `outlook_create_event` | Create a meeting or appointment, optionally recurring with a reminder |
| `outlook_update_event` | Update an existing event, including recurrence and reminder |
| `outlook_delete_event` | Delete an event |
| `outlook_respond_to_event` | Accept, decline, or tentatively accept an invite |
| `outlook_search_events` | Search events by keyword |
| `outlook_list_calendars` | List all calendars |
| `outlook_get_schedule` | Free/busy schedule (busy blocks + working hours) for people over a window |
| `outlook_find_meeting_times` | Suggest meeting slots that work across attendees |
| `outlook_cancel_event` | Cancel an event you organize and notify attendees |
| `outlook_forward_event` | Forward a meeting invite to more people |

#### Calendar: behaviour to know

- **Recurrence.** `outlook_create_event` and `outlook_update_event` take an optional `recurrence` object that maps to Microsoft Graph's `patternedRecurrence` (a `pattern` plus a `range`). Supported patterns: `daily`, `weekly` (with `days_of_week`), `absoluteMonthly` (`day_of_month`), `relativeMonthly` (e.g. last Friday via `index` + `days_of_week`), `absoluteYearly` (`month` + `day_of_month`), and `relativeYearly`. Set `interval` for "every N" (e.g. every 3 weeks). The `range` ends with `endDate` (`end_date`), `numbered` (`count`), or `noEnd`, and carries the recurrence `time_zone`. `range.start_date` defaults to the event's start date.
- **Reminders.** Use `reminder_minutes_before_start` and `is_reminder_on`. Graph and Outlook support only **one** reminder per event, so you cannot set, for example, both a one-month and a one-week reminder on a single event. Mirror the same event into separate calendars if you need several lead times.
- **Date window.** `outlook_list_events` uses Graph's `calendarView`, so a `start_date`/`end_date` window genuinely scopes the results and recurring series are expanded into individual instances. When only `start_date` is given, the window runs a week forward from it.
- **Mirroring personal events.** `show_as` (`free`/`tentative`/`busy`/`oof`/`workingElsewhere`), `categories`, and `is_private` help when copying personal events into a work calendar.

### Contacts & People

| Tool | Description |
|------|-------------|
| `outlook_list_contacts` | List/search contacts |
| `outlook_get_contact` | Get contact details by ID |
| `outlook_create_contact` | Create a new contact |
| `outlook_delete_contact` | Delete a contact |
| `outlook_search_people` | Search the organisation directory |
| `outlook_get_availability` | Check whether people are free/busy now, their out-of-office status, and working hours (free/busy and mail tips; not the live Teams presence dot) |
| `outlook_get_user_profile` | Get a colleague's directory profile (title, department, office, phone) |
| `outlook_get_manager` | Get a person's manager from the org directory |
| `outlook_get_direct_reports` | List a person's direct reports |
| `outlook_get_user_photo` | Download a person's profile photo to a file |

### Account / Out of office

| Tool | Description |
|------|-------------|
| `outlook_get_automatic_replies` | Read your out-of-office / automatic reply settings |
| `outlook_set_automatic_replies` | Turn your out-of-office on (always or scheduled) or off |

## Session storage

Session files are stored encrypted in `~/.msoutlook-mcp-server/`:

- `session-state.json`: Playwright browser session (cookies + localStorage)
- `token-cache.json`: Extracted and cached tokens
- `browser-profile/`: Persistent browser profile for headless refresh

If your session expires, run `outlook_login` again.

## Token refresh

Tokens are refreshed automatically:

1. **HTTP refresh** (fast, no browser): uses the cached refresh token with OWA's client ID
2. **Headless browser refresh**: fallback if HTTP refresh fails; opens a headless Edge window to silently reacquire tokens from the saved browser session

## Requirements

- Node.js 20+
- A Chromium based browser: Edge or Chrome (detected automatically from system default)
- A Microsoft 365 work/school account or personal Microsoft account

## Platform support

| Feature | macOS | Linux | Windows |
|---------|-------|-------|---------|
| Browser detection | System default (Edge/Chrome) | Chrome fallback | Edge (preinstalled) |
| SSO cookie import | ✅ Chrome + Edge via Keychain | ✅ Chrome + Edge via libsecret / `"peanuts"` fallback | ✅ Edge via DPAPI (PowerShell) |
| Windows Chrome 127+ cookies | n/a | n/a | ⚠️ App-Bound Encryption (not supported). Use Edge instead. |
| Headed browser fallback | ✅ | ✅ | ✅ |

Cookie import is a best effort optimisation. If it cannot run (e.g. no matching browser installed, Keychain denied), the MCP falls back to opening a headed browser where you sign in once manually, and the session then persists.

## Security notes

- Uses the same auth as the Outlook web client, so your access is limited to what your account can do
- Tokens are encrypted at rest (AES-256-GCM with a machine derived key)
- Uses undocumented internal APIs, which Microsoft may change without notice
- Always confirm email content with the user before sending

## Environment variables

| Variable | Description |
|----------|-------------|
| `MSOUTLOOK_DEBUG=true` | Enable debug logging to stderr |
| `MSOUTLOOK_BROWSER=chrome` | Force a specific browser: `chrome` or `msedge`. If unset, uses the macOS system default browser; falls back to Chrome on macOS/Linux and Edge on Windows. |
| `MSOUTLOOK_CHROME_PROFILE` | Pin a specific Chrome profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default`. |
| `MSOUTLOOK_EDGE_PROFILE` | Pin a specific Edge profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default`. |
| `MSOUTLOOK_SKIP_COOKIE_IMPORT=true` | Skip importing SSO cookies from your real browser (avoids the one time Keychain/keyring prompt). You'll sign in once manually in the browser; the persistent profile then remembers the session. |

## Troubleshooting

**A macOS Keychain prompt appears on first login.** That is the cookie import reading your browser's session. Click "Always Allow" once and it never prompts again. If you would rather not grant it, set `MSOUTLOOK_SKIP_COOKIE_IMPORT=true` and sign in manually the first time instead.

**Login says it failed but the browser showed my inbox.** Run `outlook_login` again and let the window close on its own. If it persists, run with `MSOUTLOOK_DEBUG=true` and check the stderr logs.

**It opens the wrong browser.** Set `MSOUTLOOK_BROWSER=chrome` or `MSOUTLOOK_BROWSER=msedge` to force one. Only Chromium based browsers (Chrome, Edge) are supported; Safari and Firefox fall back to Chrome/Edge.

**Tokens stopped working after a long break.** Refresh tokens last around 90 days. Run `outlook_login` to reauthenticate.

## License

MIT. See [LICENSE](./LICENSE).
