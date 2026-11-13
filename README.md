# outlook-mcp

An [MCP](https://modelcontextprotocol.io) server that gives an AI assistant access to your Outlook / Exchange Online
mailbox and calendar, signing in the way your Mac already does: through the **macOS system authentication session**
(Enterprise SSO / Company Portal), so it works in tenants whose Conditional Access requires a managed or compliant device.

macOS only. Talks to Exchange Web Services (EWS); no Microsoft Graph, no admin consent, no custom app registration.

## What makes this different

Most Outlook MCP servers sign in with MSAL device-code or browser OAuth against their own Azure app registration. In
tenants that require a managed or compliant device, or that block user consent to third-party apps, those sign-ins are
refused. This server takes a different route:

- **System sign-in session.** A small Swift helper runs `ASWebAuthenticationSession`, so the sign-in goes through the
  same Enterprise SSO / Company Portal path as other Mac apps and carries your device identity.
- **No app registration, no admin consent.** It uses the Microsoft Office public client and EWS
  (`EWS.AccessAsUser.All`), not Graph scopes that need consent.
- **Unread overview without side effects.** `list_mail` and `read_mail` never mark mail as read; only
  `mark_mail_read` does.
- **Separate draft and send tools**, so drafting can be allowlisted without allowing sending.
- **Everything under `~/.config/outlook-mcp/`**, files mode 600.

The trade-offs: macOS only, EWS instead of Graph (Microsoft has announced the retirement of EWS in Exchange
Online, so check your tenant's status), and the Office client ID is borrowed, so your organisation can still restrict it.

## Tools

| Tool | What it does |
|---|---|
| `get_events` | Events in your own calendar between two ISO datetimes (recurrences expanded). |
| `find_free_slots` | Free meeting slots inside working hours, optionally intersected with other people's free/busy. |
| `list_mail` | Overview of a folder, newest first: id, sender, subject, received, read state, preview. Supports `unreadOnly`, `since`, paging. |
| `search_mail` | Server-side mail search (Exchange query syntax) across all folders or one: `from:alice subject:(budget OR forecast) hasattachments:true received:>=2026-11-01`. Same result fields as `list_mail`. |
| `read_mail` | One message by id: headers and plain-text body. |
| `mark_mail_read` | Mark messages read or unread (`isRead: false`). No read receipts are sent. |
| `draft_mail` | Save a message to Drafts. Sends nothing. |
| `send_mail` | Send a message from your mailbox. Irreversible. |

Behaviour worth knowing:

- `list_mail` and `read_mail` never mark anything as read. Only `mark_mail_read` changes read state, so you can get an
  overview of unread mail without touching it.
- `search_mail` runs Exchange's Advanced Query Syntax on the server, newest first. Searching "all" covers every mail
  folder except Deleted Items and Junk Email; it queries folders a few at a time to stay under EWS's concurrent
  connection limit, so very large mailboxes with many folders are slower than a single-folder search. `since` has
  day granularity.
- `find_free_slots` treats attendees whose free/busy can't be fetched as free and says so in `warnings`. Free/busy of
  others only works if your organisation shares it. Range is limited to 62 days.
- `send_mail` and `draft_mail` take `to`, `cc`, `bcc`, `subject`, `body`, `isHtml`. Prefer `draft_mail` first.
- Read-only tools are annotated `readOnlyHint`, so you can allowlist them separately from `mark_mail_read`,
  `draft_mail` and `send_mail`.

## Setup

Requirements: macOS, Node 20+, Xcode command line tools (`swiftc`, for the sign-in helper).

```sh
npm install
npm run build                         # compiles helper/main.swift -> bin/auth-helper, then tsc
node dist/index.js init --tenant <tenant-guid-or-domain>
npm run login                         # opens the macOS sign-in sheet, once
```

Register it with Claude Code (use absolute paths; the server must be started with the script as an argument):

```sh
claude mcp add outlook -s local -- "$(which node)" "$PWD/dist/index.js"
```

Restart Claude Code (or `/mcp` and reconnect) afterwards.

`init` accepts `--tenant`, `--client-id`, `--redirect-uri` and `--force` (overwrite an existing config).

## Configuration

Everything lives in `~/.config/outlook-mcp/` (directory mode 700, files mode 600):

| File | Contents |
|---|---|
| `config.json` | `tenantId` (required), and optionally `clientId`, `redirectUri`, `resource`, `ewsUrl`, `email`, `timeZone`. |
| `token-cache.json` | OAuth refresh and access token. **Treat as a password.** |

Environment variables override the file: `OUTLOOK_MCP_CONFIG`, `OUTLOOK_MCP_CACHE`, `OUTLOOK_MCP_TENANT_ID`,
`OUTLOOK_MCP_CLIENT_ID`, `OUTLOOK_MCP_REDIRECT_URI`, `OUTLOOK_MCP_RESOURCE`, `OUTLOOK_MCP_EWS_URL`, `OUTLOOK_MCP_EMAIL`,
`OUTLOOK_MCP_TIMEZONE`.

Defaults: the Microsoft Office public client ID, redirect URI `urn:ietf:wg:oauth:2.0:oob`, resource
`https://outlook.office365.com`, EWS endpoint `https://outlook.office365.com/EWS/Exchange.asmx`. Your address is taken
from the token unless `email` is set. Working hours default to 09:00-17:00, Monday to Friday, in `timeZone` (or the
machine's zone).

## How sign-in works

1. `npm run login` builds the Azure AD v1 authorize URL and passes it to `bin/auth-helper`, a small Swift program.
2. The helper runs an `ASWebAuthenticationSession`. This is the system sign-in session, so Enterprise SSO, Company
   Portal and device-compliance checks apply exactly as they do for other Mac apps. The `urn:ietf:wg:oauth:2.0:oob`
   redirect is a custom scheme the session can intercept.
3. The authorization code is exchanged for tokens, and the refresh token is written to `token-cache.json`.
4. The server refreshes access tokens silently. If the refresh token is rejected, run `npm run login` again.

Permissions come from the app registration of the client ID (v1 is resource-based, not scope-based). The default Office
client normally carries `EWS.AccessAsUser.All`, which covers reading and sending mail and calendar access. If your
tenant restricts it, calls fail with an EWS error rather than silently.

### Compared with other Outlook MCP servers

As far as a README-level survey of GitHub projects shows (source code not audited), other Outlook MCP servers use Graph
with MSAL device-code or browser OAuth, drive the Outlook app via AppleScript, or reuse a browser session. Device-code
login is a separate flow from the Mac's SSO session and generally cannot satisfy "managed / compliant device"
Conditional Access policies. None found uses `ASWebAuthenticationSession`.

## Development

```sh
npm run build:ts   # tsc only
npm test           # tsc + unit tests for the free-slot logic
```

Layout: `src/index.ts` (tools, `init`/`login` subcommands), `src/ews.ts` (EWS requests), `src/auth.ts` (tokens),
`src/config.ts`, `src/slots.ts` (free-range maths), `helper/main.swift` (system auth session).

## Security notes

- The server can read and send mail as you. Anyone who can read `token-cache.json` has the same access until the
  session is revoked; keep `~/.config/outlook-mcp` out of backups and sync folders you share.
- `send_mail` is irreversible. Keep Claude Code's approval prompt on for it.
- Nothing secret is stored in this repository; tokens and tenant ID live only in `~/.config/outlook-mcp/`.
