<div align="center">

# ghub

Gmail, Drive, Sheets, Docs, and Calendar for AI agents, across several Google accounts, as an MCP server and a JSON CLI.

[![Verify and release](https://github.com/tszaks/ghub/actions/workflows/release.yml/badge.svg)](https://github.com/tszaks/ghub/actions/workflows/release.yml)
[![Release](https://img.shields.io/github/v/release/tszaks/ghub)](https://github.com/tszaks/ghub/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Node.js 20+](https://img.shields.io/badge/node-%3E%3D20-brightgreen)](https://nodejs.org/)

</div>

```text
$ ghub call search_emails --json '{"account":"work","query":"is:unread","max_results":10}' \
    --fields emails.id,emails.subject --limit 10
{"ok":true,"data":{"emails":[{"id":"...","subject":"..."}]},"meta":{"version":1}}
```

ghub is a local server that gives an AI assistant access to your Google accounts. It speaks the Model Context Protocol (MCP) over stdio or SSE, and the same 60 tools are also available as short CLI calls that print one JSON object. You sign in once per account with a built-in OAuth flow. After that, an agent can read and search across all your inboxes at once, or act on one account you name.

It is for people who run agents (Claude Code, Codex, Claude Desktop, or their own) and use more than one Google account. It exists because most Gmail MCP servers handle one account and need manual token setup.

## Features

- **Multi-account:** Configure as many Google accounts as you want. Read and search tools merge results from all enabled accounts when you leave `account` out. Write tools require an explicit `account`.
- **Five Google services:** 60 tools across Gmail, Drive, Sheets, Docs, and Calendar.
- **Built-in OAuth:** Onboard an account with `begin_account_auth` and `finish_account_auth`. Refreshed tokens are saved automatically.
- **Attachments:** `get_all_attachments` extracts text from PDF, Word, and Excel files. Images are read with a small macOS Vision OCR helper when `swiftc` is available at install time.
- **Agent CLI:** `ghub tools`, `ghub schema TOOL`, and `ghub call TOOL` give structured JSON, bounded output, validation-only `--dry-run`, and fixed exit codes. See [docs/agent-cli.md](docs/agent-cli.md).
- **Two MCP transports:** stdio by default, or HTTP with SSE for hosted setups.

## Install

You need Node.js 20 or newer.

From a clone (the compiled `dist/` is committed, so this works without a build step):

```sh
git clone https://github.com/tszaks/ghub.git
cd ghub
npm install
npm run build
node dist/index.js tools
```

Or from a release tarball (see [docs/releases.md](docs/releases.md)):

```sh
npm install -g ./ghub-1.7.0.tgz
ghub tools
```

`npm install` runs `scripts/build-ocr.sh`, which compiles the OCR helper only if `swiftc` is found (macOS). Without it, image attachments are saved to disk but no text is extracted.

ghub is not published to the npm registry.

## Quick start

1. Create Google OAuth credentials (see [Google Cloud setup](#google-cloud-setup)) and download the JSON file.
2. Add ghub to your MCP client (see [Client configuration](#client-configuration)).
3. Ask your agent to call `begin_account_auth`, open the URL it returns, then pass the code to `finish_account_auth`. Details are in [OAuth onboarding](#oauth-onboarding).
4. Call `list_accounts` to check that each account has a valid token.

## Client configuration

ghub runs as `node /absolute/path/to/ghub/dist/index.js` (or `ghub mcp` if installed from the tarball). With no arguments it starts the MCP server on stdio.

**Claude Code**

```sh
claude mcp add ghub -- node /absolute/path/to/ghub/dist/index.js
```

**Codex** (`~/.codex/config.toml`)

```toml
[mcp_servers.ghub]
command = "node"
args = ["/absolute/path/to/ghub/dist/index.js"]
```

**Claude Desktop** (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "ghub": {
      "command": "node",
      "args": ["/absolute/path/to/ghub/dist/index.js"]
    }
  }
}
```

The server starts and lists its tools over stdio with no extra setup. Google calls begin only after you onboard an account.

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `GMAILMCPCONFIG_DIR` | `~/.gmail-multi-mcp` | Where account files and tokens are stored. |
| `GMAIL_MCP_CONFIG_DIR` | none | Older name for the same setting. `GMAILMCPCONFIG_DIR` wins if both are set. |
| `MCP_TRANSPORT` | `stdio` | Set to `sse` to serve MCP over HTTP. |
| `PORT` | none | If set, the server switches to SSE mode and listens on it. Falls back to `RAILWAY_PORT`, then `3000`, when SSE is forced. |
| `HOST` | `0.0.0.0` | Bind address in SSE mode. |

To pass a variable to the server, add it to the client config, for example `claude mcp add ghub -e GMAILMCPCONFIG_DIR=/custom/path -- node /absolute/path/to/ghub/dist/index.js`.

### Where data is stored

Default configuration location: `~/.gmail-multi-mcp/`

```
~/.gmail-multi-mcp/
├── accounts.json           # Master account list
└── accounts/
    ├── personal/
    │   ├── credentials.json  # OAuth client credentials
    │   ├── token.json       # Access/refresh tokens
    │   └── meta.json        # Account metadata
    └── work/
        ├── credentials.json
        ├── token.json
        └── meta.json
```

### accounts.json format

```json
{
  "defaultAccount": "personal",
  "accounts": [
    {
      "id": "personal",
      "email": "user@gmail.com",
      "displayName": "Personal Gmail",
      "enabled": true,
      "credentialPath": "~/.gmail-multi-mcp/accounts/personal/credentials.json",
      "tokenPath": "~/.gmail-multi-mcp/accounts/personal/token.json"
    },
    {
      "id": "work",
      "email": "user@company.com",
      "displayName": "Work Email",
      "enabled": true,
      "credentialPath": "~/.gmail-multi-mcp/accounts/work/credentials.json",
      "tokenPath": "~/.gmail-multi-mcp/accounts/work/token.json"
    }
  ]
}
```

## Google Cloud setup

Before using ghub, you need to set up a Google Cloud project:

### 1. Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Note your project ID

### 2. Enable Gmail API

1. In your project, go to **APIs & Services** > **Library**
2. Search for "Gmail API"
3. Click **Enable**
4. Enable Google Drive, Google Sheets, Google Docs, and Google Calendar APIs
   for the corresponding tools you plan to use.

### 3. Create OAuth 2.0 Credentials

1. Go to **APIs & Services** > **Credentials**
2. Click **+ CREATE CREDENTIALS** > **OAuth client ID**
3. If prompted, configure the OAuth consent screen:
   - Choose "External" user type
   - Fill in required fields (app name, support email)
   - Add your email to test users
4. For application type, select **Desktop app**
5. Give it a name (e.g., "Gmail MCP Client")
6. Click **Create**
7. **Download the JSON file** - you'll need this for authentication

### 4. OAuth Consent Screen Setup

1. Go to **APIs & Services** > **OAuth consent screen**
2. Review the existing broad scopes requested by GHub:
   - `https://mail.google.com/`
   - `https://www.googleapis.com/auth/gmail.settings.basic`
   - `https://www.googleapis.com/auth/drive`
   - `https://www.googleapis.com/auth/spreadsheets`
   - `https://www.googleapis.com/auth/documents`
   - `https://www.googleapis.com/auth/calendar`
3. Add your Gmail address(es) as test users

These scopes permit writes and are unchanged by the CLI. Human consent is
required during the separate OAuth onboarding flow.


## SSE deployment

- Set `MCP_TRANSPORT=sse` to force HTTP/SSE mode. If `PORT` is present, SSE mode starts automatically.
- The SSE endpoint is `/sse` and message posts go to `/messages`. `GET /` returns a plain text health line.
- For Railway, point the start command at `npm start` or `node dist/index.js` and let Railway provide `PORT`.

The SSE server has no authentication of its own. Anyone who can reach the port can use every configured account. Run it only on a private network or behind your own access control.

## OAuth onboarding

Authenticate accounts directly through MCP tools:

### Step 1: Start Authentication

Call the `begin_account_auth` tool with your OAuth credentials:

```typescript
{
  "account_id": "personal",
  "email": "user@gmail.com",
  "credentials_json": {
    "installed": {
      "client_id": "YOUR_CLIENT_ID.apps.googleusercontent.com",
      "client_secret": "YOUR_CLIENT_SECRET",
      "redirect_uris": ["http://localhost"]
    }
  }
}
```

Or use a file path:

```typescript
{
  "account_id": "personal",
  "email": "user@gmail.com",
  "credentials_path": "/path/to/credentials.json"
}
```

The tool returns a Google OAuth URL. Open this URL in your browser.

### Step 2: Complete Authentication

1. In your browser, sign in with the Gmail account
2. Grant the requested permissions
3. Google redirects you to a localhost URL with a `code` parameter
4. Copy the authorization code from the URL

Call `finish_account_auth`:

```typescript
{
  "account_id": "personal",
  "authorization_code": "4/0AfJoh..."
}
```

Your account is now authenticated and ready to use.

## Usage

### Example 1: Read Recent Emails from All Accounts

```typescript
// Aggregates across all enabled accounts
{
  "max_results": 20,
  "include_body": true
}

// Returns emails with source account indicated
```

### Example 2: Search Across Multiple Accounts

```typescript
{
  "query": "from:boss@company.com is:unread",
  "max_results": 10
}

// Searches all enabled accounts, merges and sorts results
```

### Example 3: Send Email from Specific Account

```typescript
{
  "account": "work",
  "to": "colleague@company.com",
  "subject": "Project Update",
  "body": "Here's the latest on the project...",
  "html": false
}
```

### Example 4: Read Only from One Account

```typescript
{
  "account": "personal",
  "max_results": 10,
  "query": "label:important"
}
```

### Example 5: Manage Labels

```typescript
// Create a new label
{
  "account": "personal",
  "name": "Urgent-2026"
}

// Add label to messages
{
  "account": "personal",
  "message_ids": ["msg123", "msg456"],
  "label_ids": ["Label_789"]
}
```

### Example 6: Delete Drafts

```typescript
// Delete one or more drafts by their draft IDs (returned by create_draft)
{
  "account": "personal",
  "draft_ids": ["r5457071851533655344", "r774137312565667821"]
}
```


## Tools

All 60 tools, taken from the live registry (`ghub tools`). Kind is `read`, `write`, or `auth`. Parameters for the first Gmail tools are in [docs/api-reference.md](docs/api-reference.md). For the exact input schema of any tool, run `ghub schema TOOL`.

### Accounts and auth

| Tool | Kind | What it does |
| --- | --- | --- |
| `list_accounts` | read | List all configured Gmail inbox accounts and their authentication/health status. |
| `begin_account_auth` | auth | Start OAuth onboarding for an account. |
| `finish_account_auth` | auth | Complete OAuth onboarding by exchanging authorization code and storing token.json. |

### Gmail

| Tool | Kind | What it does |
| --- | --- | --- |
| `read_emails` | read | Read emails from one account or aggregate across all enabled accounts when account is omitted. |
| `search_emails` | read | Search Gmail using query syntax. |
| `get_email_thread` | read | Get a full email thread for one account. |
| `get_labels` | read | List labels for one account. |
| `mark_as_read` | write | Mark messages as read in one account. |
| `add_labels` | write | Add labels to messages in one account. |
| `remove_labels` | write | Remove labels from messages in one account. |
| `archive_emails` | write | Archive messages in one account by removing INBOX label. |
| `trash_emails` | write | Move messages to trash in one account. |
| `create_label` | write | Create a new Gmail label in one account. |
| `delete_label` | write | Delete a Gmail label in one account. |
| `block_sender` | write | Block a sender by creating a Gmail filter (same mechanism as Gmail's native "Block" button). |
| `list_blocked_senders` | read | List all Gmail filters for an account. |
| `unblock_sender` | write | Remove a Gmail filter. |
| `unsubscribe_from_email` | write | Unsubscribe from a mailing list by invoking the List-Unsubscribe header (RFC 2369/8058). |
| `mute_thread` | write | Mute a conversation: archives the thread now and (with scope="subject") installs a subject-based filter so future replies auto-archive. |
| `create_draft` | write | Create a draft email in one account. |
| `delete_drafts` | write | Permanently delete one or more drafts in one account. |
| `send_draft` | write | Send an existing Gmail draft by draft ID. |
| `list_drafts` | read | List drafts in an account sorted oldest-first by internalDate. |
| `search_drafts` | read | Search drafts by Gmail query string. |
| `send_email` | write | Send an email from one account. |
| `get_attachment` | read | Fetch a single attachment by id. |
| `get_all_attachments` | read | Fetch and extract text from every attachment on an email in one call (PDF, Word, Excel, images via OCR). |

### Drive

| Tool | Kind | What it does |
| --- | --- | --- |
| `search_drive_files` | read | Search Google Drive file metadata. |
| `list_drive_files` | read | List Google Drive files. |
| `get_drive_file` | read | Get full metadata for a specific Google Drive file including size, parents, export links. |
| `get_drive_file_content` | read | Download a Drive file and extract its text content. |
| `upload_drive_file` | write | Upload a local file to Google Drive. |
| `create_drive_folder` | write | Create a new folder in Google Drive. |
| `update_drive_file` | write | Update a Drive file: rename it, move it between folders, star/unstar, or update its description. |
| `trash_drive_file` | write | Move a Google Drive file to trash. |
| `share_drive_file` | write | Share a Google Drive file by adding a permission. |

### Sheets

| Tool | Kind | What it does |
| --- | --- | --- |
| `sheets_get` | read | Get Google Sheets spreadsheet metadata: title, sheet tabs, row/column counts. |
| `sheets_read` | read | Read cell values from a Google Sheet range (e.g., "Sheet1!A1:D10" or "A1:D10" for the first sheet). |
| `sheets_write` | write | Write values to a Google Sheet range. |
| `sheets_append` | write | Append rows to a Google Sheet after the last row with data. |
| `sheets_create` | write | Create a new Google Sheets spreadsheet. |
| `sheets_add_tab` | write | Add a new sheet tab to an existing spreadsheet. |
| `sheets_rename_tab` | write | Rename a sheet tab in a spreadsheet. |
| `sheets_delete_tab` | write | Delete a sheet tab from a spreadsheet. |
| `sheets_format` | write | Apply formatting to a range of cells: bold, italic, font size, background/text color, alignment, number format, wrap strategy. |
| `sheets_add_chart` | write | Create a chart in a Google Sheet from a data range. |
| `sheets_insert_dimension` | write | Insert rows or columns into a sheet at a specified position. |
| `sheets_delete_dimension` | write | Delete rows or columns from a sheet. |

### Docs

| Tool | Kind | What it does |
| --- | --- | --- |
| `docs_get` | read | Get the full text content and title of a Google Doc. |
| `docs_create` | write | Create a new Google Doc. |
| `docs_append` | write | Append text to the end of a Google Doc. |
| `docs_replace_text` | write | Find and replace all occurrences of text throughout a Google Doc. |
| `docs_insert_table` | write | Insert an empty table at the end of a Google Doc. |
| `docs_apply_style` | write | Apply a heading or paragraph style to a text range in a Google Doc by character indices. |

### Calendar

| Tool | Kind | What it does |
| --- | --- | --- |
| `list_calendars` | read | List all Google Calendars for one or all enabled accounts. |
| `list_events` | read | List calendar events. |
| `get_event` | read | Get full details of a single calendar event. |
| `create_event` | write | Create a new calendar event. |
| `update_event` | write | Update an existing calendar event. |
| `delete_event` | write | Delete a calendar event. |

## Permissions and privacy

- **Scopes requested:** full Gmail access (`https://mail.google.com/`), Gmail settings (filters), full Drive, Sheets, Docs, and Calendar. These allow reading and writing. ghub does not narrow them per tool.
- **What it can read:** mail, attachments, Drive files, spreadsheets, documents, and calendar events in every account you onboard.
- **What it can change:** send mail, delete drafts, trash mail, create and delete filters, upload, move, share, and trash Drive files, edit Sheets and Docs, and create, update, and delete events. There is no confirmation prompt in the server. Your MCP client's tool approval is the safety check.
- **Where data goes:** ghub talks to Google's APIs and nothing else. It sends no telemetry. Content you read is returned to your MCP client or CLI caller, and from there to whichever model that client uses.
- **Where credentials live:** OAuth client files and tokens are stored in plain JSON under the config directory (`~/.gmail-multi-mcp` by default). Protect that folder. `credentials.json` and `token.json` are git-ignored.
- **Local files:** `send_email` and `create_draft` read attachment paths you give them, `upload_drive_file` reads a local file, and attachment tools save files to disk.

## Using with agents

An agent can discover everything it needs at run time:

```sh
ghub tools                    # names, descriptions, read/write/auth kind
ghub schema search_emails     # input schema, output envelope, example
ghub call send_email --dry-run --input - <<'JSON'
{"account":"work","to":"recipient@example.test","subject":"Example","body":"Hello"}
JSON
```

Use `--dry-run` to validate a write before running it. Mutations never pick a default account. See [docs/agent-cli.md](docs/agent-cli.md) for the JSON contract and exit codes, and [docs/agent-cli-spec.md](docs/agent-cli-spec.md) for the specification.

## FAQ and troubleshooting

### "Invalid grant" Error

This usually means your authorization code has expired. Authorization codes are single-use and expire after a few minutes.

**Solution:** Run `begin_account_auth` again to get a fresh OAuth URL and authorization code.

### "Token has been expired or revoked"

Your refresh token is no longer valid.

**Solution:**
1. Delete the `token.json` file for the affected account
2. Run the OAuth flow again (`begin_account_auth` then `finish_account_auth`)

### "Insufficient permissions"

The OAuth token doesn't have the required scopes.

**Solution:**
1. Check your Google Cloud OAuth consent screen has all required scopes
2. Review the [current requested scopes](#4-oauth-consent-screen-setup), API
   enablement, and access to the specific resource.
3. If the grant is missing required scopes, have a human review consent and
   complete the separate OAuth flow again. CLI calls do not start this
   automatically.

### Account Not Found

The specified account ID doesn't exist in `accounts.json`.

**Solution:**
1. Run `list_accounts` to see available accounts
2. Ensure you completed OAuth onboarding for the account
3. Check that the account is `enabled: true` in `accounts.json`

### Rate Limiting

Gmail API has rate limits (daily quota and per-user quotas).

**Solution:**
1. Check your quota at [Google Cloud Console](https://console.cloud.google.com/apis/api/gmail.googleapis.com/quotas)
2. Implement exponential backoff in your application
3. Consider applying for increased quota if needed

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md). Report security problems as described in [SECURITY.md](SECURITY.md). Everyone taking part follows the [Code of Conduct](CODE_OF_CONDUCT.md). Agents working on the code should read [AGENTS.md](AGENTS.md).

Docs: [Agent CLI](docs/agent-cli.md), [CLI spec](docs/agent-cli-spec.md), [API reference](docs/api-reference.md), [Releases](docs/releases.md), [Changelog](CHANGELOG.md).

## License

[MIT](LICENSE)

## Acknowledgments

- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- Uses the [Google APIs Node.js Client](https://github.com/googleapis/google-api-nodejs-client)

Built by [Tyler Szakacs](https://github.com/tszaks).
