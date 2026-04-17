# Gmail Multi-Inbox MCP Server

> A Model Context Protocol (MCP) server that enables AI assistants to manage multiple Gmail accounts simultaneously with built-in OAuth authentication.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D20-brightgreen)](https://nodejs.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.7-blue)](https://www.typescriptlang.org/)

## Why This MCP Server?

Unlike existing Gmail MCP servers, this implementation offers:

- **Multi-Account Support**: Manage multiple Gmail accounts from a single MCP server instance
- **Intelligent Aggregation**: Read and search across all accounts simultaneously
- **Built-in OAuth Flow**: Complete authentication setup through MCP tools - no manual token generation
- **Type-Safe**: Built with TypeScript for reliability and better IDE support
- **Auto Token Refresh**: Handles token expiration automatically and saves refreshed tokens
- **Comprehensive API**: Full Gmail API coverage including labels, threads, drafts, and more

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Google Cloud Setup](#google-cloud-setup)
- [Configuration](#configuration)
- [Railway / SSE Deployment](#railway--sse-deployment)
- [OAuth Onboarding](#oauth-onboarding)
- [Usage Examples](#usage-examples)
- [API Reference](#api-reference)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Features

### Read Operations
- **`list_accounts`** - View all configured accounts and their status
- **`read_emails`** - Fetch recent emails (aggregates across all accounts by default)
- **`search_emails`** - Search using Gmail query syntax across multiple accounts
- **`get_email_thread`** - Retrieve complete conversation threads
- **`get_labels`** - List all labels for an account

### Write Operations
- **`send_email`** - Send emails from any configured account, with optional local file attachments
- **`create_draft`** - Create draft messages, with optional local file attachments
- **`delete_drafts`** - Permanently delete one or more drafts by draft ID
- **`mark_as_read`** - Mark messages as read
- **`archive_emails`** - Archive messages (remove from inbox)
- **`trash_emails`** - Move messages to trash

### Label Management
- **`add_labels`** - Apply labels to messages
- **`remove_labels`** - Remove labels from messages
- **`create_label`** - Create new custom labels
- **`delete_label`** - Delete existing labels

### Account Management
- **`begin_account_auth`** - Start OAuth flow for new account
- **`finish_account_auth`** - Complete OAuth and save credentials

## Prerequisites

- **Node.js 20+** ([Download](https://nodejs.org/))
- **A Google Cloud Project** with Gmail API enabled
- **OAuth 2.0 Credentials** (Desktop application type)

## Installation

```bash
# Clone the repository
git clone https://github.com/tszaks/gmail-multi-inbox-mcp.git
cd gmail-multi-inbox-mcp

# Install dependencies
npm install

# Build the TypeScript code
npm run build
```

## Google Cloud Setup

Before using this MCP server, you need to set up a Google Cloud project:

### 1. Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Note your project ID

### 2. Enable Gmail API

1. In your project, go to **APIs & Services** > **Library**
2. Search for "Gmail API"
3. Click **Enable**

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
2. Add the following scopes:
   - `https://www.googleapis.com/auth/gmail.modify`
   - `https://www.googleapis.com/auth/gmail.send`
   - `https://www.googleapis.com/auth/userinfo.email`
3. Add your Gmail address(es) as test users

## Configuration

### Add to Your MCP Settings

Add this server to your `.mcp.json` configuration file (adjust path to match your installation):

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

### Custom Config Directory (Optional)

To use a custom directory for storing account data:

```json
{
  "mcpServers": {
    "gmail-multi": {
      "command": "node",
      "args": [
        "/path/to/gmail-multi-inbox-mcp/dist/index.js"
      ],
      "env": {
        "GMAILMCPCONFIG_DIR": "/custom/path/.gmail-multi-mcp"
      }
    }
  }
}
```

The server also accepts the legacy `GMAIL_MCP_CONFIG_DIR` env var.

### Directory Structure

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

### accounts.json Format

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

## Railway / SSE Deployment

This server can run as either stdio MCP or an SSE-backed HTTP server.

- Set `MCP_TRANSPORT=sse` to force HTTP/SSE mode.
- If `PORT` is present, the server automatically switches to SSE mode.
- The SSE endpoint is served at `/sse` and message posts are accepted at `/messages`.
- Configuration still respects `GMAILMCPCONFIG_DIR` and `GMAIL_MCP_CONFIG_DIR`.

For Railway, point the start command at `npm start` or `node dist/index.js` and let Railway provide `PORT`.

## OAuth Onboarding

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

## Usage Examples

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

## API Reference

### Read Operations

#### `list_accounts`
Returns all configured accounts with health status.

**Parameters:** None

**Returns:**
```typescript
{
  accounts: Array<{
    id: string
    email: string
    displayName: string
    enabled: boolean
    hasValidToken: boolean
  }>
}
```

#### `read_emails`
Fetch recent emails with optional filtering.

**Parameters:**
- `account` (optional): Account ID to read from. Omit to aggregate all accounts.
- `max_results` (optional, default: 20): Number of emails to return (1-100)
- `query` (optional): Gmail search query
- `include_body` (optional, default: false): Include plaintext body extraction

**Returns:** Array of email objects with metadata, headers, and optional body

#### `search_emails`
Search emails using Gmail query syntax.

**Parameters:**
- `query` (required): Gmail search query
- `account` (optional): Account ID to search. Omit to search all accounts.
- `max_results` (optional, default: 25): Maximum results (1-100)

**Returns:** Array of matching emails

#### `get_email_thread`
Retrieve a complete email thread.

**Parameters:**
- `account` (required): Account ID
- `thread_id` (required): Gmail thread ID

**Returns:** Thread object with all messages

#### `get_labels`
List all labels for an account.

**Parameters:**
- `account` (required): Account ID

**Returns:** Array of label objects with IDs and names

### Write Operations

#### `send_email`
Send an email from a specific account.

**Parameters:**
- `account` (required): Account ID
- `to` (required): Recipient email address(es)
- `subject` (required): Email subject
- `body` (required): Email body
- `cc` (optional): CC recipients
- `bcc` (optional): BCC recipients
- `html` (optional, default: false): Send as HTML
- `attachments` (optional): Array of local file attachments. Each item supports:
  - `path` (required): Absolute or local filesystem path
  - `filename` (optional): Override the filename shown in Gmail
  - `content_type` (optional): Override the MIME type, for example `application/pdf`

**Returns:** Sent message details

#### `create_draft`
Create a draft email.

**Parameters:** Same as `send_email`

**Returns:** Draft details including `draft_id` and `thread_id`

#### `delete_drafts`
Permanently delete one or more drafts. Uses draft IDs as returned by `create_draft`. Note: draft IDs differ from message IDs and cannot be used with `trash_emails`.

**Parameters:**
- `account` (required): Account ID
- `draft_ids` (required): Array of draft IDs to delete

**Returns:** Count of deleted drafts

#### `mark_as_read`
Mark messages as read.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

#### `archive_emails`
Archive messages (remove INBOX label).

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

#### `trash_emails`
Move messages to trash.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

### Label Operations

#### `add_labels`
Add labels to messages.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs
- `label_ids` (required): Array of label IDs

#### `remove_labels`
Remove labels from messages.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs
- `label_ids` (required): Array of label IDs

#### `create_label`
Create a new Gmail label.

**Parameters:**
- `account` (required): Account ID
- `name` (required): Label name
- `label_list_visibility` (optional, default: "labelShow")
- `message_list_visibility` (optional, default: "show")

#### `delete_label`
Delete a Gmail label.

**Parameters:**
- `account` (required): Account ID
- `label_id` (required): Label ID to delete

## Troubleshooting

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
2. Re-run the OAuth flow to grant new permissions
3. Required scopes:
   - `https://www.googleapis.com/auth/gmail.modify`
   - `https://www.googleapis.com/auth/gmail.send`
   - `https://www.googleapis.com/auth/userinfo.email`

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

Contributions are welcome. Here's how:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add your feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

### Development Scripts

```bash
# Watch mode for development
npm run dev

# Type checking
npm run typecheck

# Build for production
npm run build

# Run the server
npm run start
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- Uses [Google APIs Node.js Client](https://github.com/googleapis/google-api-nodejs-client)

## Links

- [MCP Documentation](https://modelcontextprotocol.io/)
- [Gmail API Documentation](https://developers.google.com/gmail/api)
- [Issues & Bug Reports](https://github.com/tszaks/gmail-multi-inbox-mcp/issues)

---

Built by [Tyler Szakacs](https://github.com/tszaks)

## Quickstart TL;DR

```bash
npm install
npm run build
node dist/index.js
```

Then add the server to MCP config and onboard each account with `begin_account_auth` + `finish_account_auth`.

## How It Works (TL;DR)

- Server stores account-level credentials/tokens in local config directory
- OAuth flow is handled via MCP tools
- Read/search can aggregate across all enabled accounts
- Gmail API calls execute per selected account

## LLM Quick Copy

Use the copy button on this code block in GitHub.

```txt
Repo: gmail-multi-inbox-mcp
Goal: Multi-account Gmail MCP with built-in OAuth onboarding.
Setup:
1) npm install && npm run build
2) Add to MCP config
3) Run begin_account_auth + finish_account_auth for each inbox
Use:
- list_accounts to verify health
- read_emails/search_emails aggregated or per account
- send_email/create_draft/delete_drafts/label tools for write actions
How it works:
- Node MCP server maintains per-account token files and calls Gmail API
```
