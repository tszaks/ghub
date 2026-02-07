# Gmail Multi-Inbox MCP

A brand-new MCP server for Gmail that supports multiple inboxes in one server.

- Read/search can aggregate across all enabled accounts.
- Send/draft/admin actions require an explicit `account` value.
- OAuth onboarding is built in via MCP tools.

## Features

### Read and search
- `list_accounts`
- `read_emails`
- `search_emails`
- `get_email_thread`
- `get_labels`

### Write and admin
- `send_email`
- `create_draft`
- `mark_as_read`
- `add_labels`
- `remove_labels`
- `archive_emails`
- `trash_emails`
- `create_label`
- `delete_label`

### Account onboarding
- `begin_account_auth`
- `finish_account_auth`

## Requirements

- Node.js 20+
- A Google Cloud OAuth client with Gmail API enabled

## Install and Build

```bash
git clone https://github.com/tszaks/gmail-multi-inbox-mcp.git
cd gmail-multi-inbox-mcp
npm install
npm run build
```

## MCP Config

Add to your global `.mcp.json` (adjust path to your installation):

```json
{
  "mcpServers": {
    "gmail-multi": {
      "command": "node",
      "args": [
        "/path/to/gmail-multi-inbox-mcp/dist/index.js"
      ]
    }
  }
}
```

Optional environment override:

```json
{
  "mcpServers": {
    "gmail-multi": {
      "command": "node",
      "args": [
        "/path/to/gmail-multi-inbox-mcp/dist/index.js"
      ],
      "env": {
        "GMAIL_MCP_CONFIG_DIR": "/custom/path/.gmail-multi-mcp"
      }
    }
  }
}
```

## Config Layout

Default root: `~/.gmail-multi-mcp`

```
~/.gmail-multi-mcp/
  accounts.json
  accounts/
    personal/
      credentials.json
      token.json
      meta.json
    business/
      credentials.json
      token.json
      meta.json
```

### `accounts.json` example

```json
{
  "defaultAccount": "personal",
  "accounts": [
    {
      "id": "personal",
      "email": "your-email@gmail.com",
      "displayName": "Personal",
      "enabled": true,
      "credentialPath": "~/.gmail-multi-mcp/accounts/personal/credentials.json",
      "tokenPath": "~/.gmail-multi-mcp/accounts/personal/token.json"
    }
  ]
}
```

## OAuth Onboarding Flow

### 1) Start account auth

Call `begin_account_auth` with:
- `account_id`
- `email`
- one of:
  - `credentials_json` (JSON object or JSON string)
  - `credentials_path`

The tool returns a Google OAuth URL.

### 2) Finish account auth

Call `finish_account_auth` with:
- `account_id`
- `authorization_code`

This stores `token.json`, enables the account, and validates profile access.

## Behavior Rules

- `read_emails` and `search_emails`:
  - if `account` is omitted, server aggregates across all enabled accounts
  - partial account failures are reported in `Account Errors`
- Write/admin tools:
  - always require explicit `account`

## Scripts

- `npm run build` compile TypeScript to `dist/`
- `npm run typecheck` run strict type check
- `npm run start` run server from `dist/index.js`

## Notes

- This server does not reuse existing Gmail MCP code.
- Tokens are refreshed automatically and saved back to each account `token.json`.
