# API reference

Parameters for the Gmail tools that were documented first. For the full, current input schema of any of the 60 tools, run `ghub schema TOOL` (see [agent-cli.md](agent-cli.md)).

## Read Operations

### `list_accounts`
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

### `read_emails`
Fetch recent emails with optional filtering.

**Parameters:**
- `account` (optional): Account ID to read from. Omit to aggregate all accounts.
- `max_results` (optional, default: 20): Number of emails to return (1-100)
- `query` (optional): Gmail search query
- `include_body` (optional, default: false): Include plaintext body extraction

**Returns:** Array of email objects with metadata, headers, and optional body

### `search_emails`
Search emails using Gmail query syntax.

**Parameters:**
- `query` (required): Gmail search query
- `account` (optional): Account ID to search. Omit to search all accounts.
- `max_results` (optional, default: 25): Maximum results (1-100)

**Returns:** Array of matching emails

### `get_email_thread`
Retrieve a complete email thread.

**Parameters:**
- `account` (required): Account ID
- `thread_id` (required): Gmail thread ID

**Returns:** Thread object with all messages

### `get_labels`
List all labels for an account.

**Parameters:**
- `account` (required): Account ID

**Returns:** Array of label objects with IDs and names

## Write Operations

### `send_email`
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
  - `path` (required): Absolute path to the file (a leading `~` is expanded). A plain path string is also accepted in place of the object. A path that cannot be read fails the call instead of sending without it
  - `filename` (optional): Override the filename shown in Gmail
  - `content_type` (optional): Override the MIME type, for example `application/pdf`
- `thread_id` (optional): Send as a reply in this Gmail thread
- `in_reply_to` (optional): Message-ID being replied to. Derived from `thread_id` when omitted
- `references` (optional): References header value. Derived from `thread_id` when omitted

**Returns:** Sent message details

### `create_draft`
Create a draft email.

**Parameters:** Same as `send_email`

**Returns:** Draft details including `draft_id` and `thread_id`

### `delete_drafts`
Permanently delete one or more drafts. Uses draft IDs as returned by `create_draft`. Note: draft IDs differ from message IDs and cannot be used with `trash_emails`.

**Parameters:**
- `account` (required): Account ID
- `draft_ids` (required): Array of draft IDs to delete

**Returns:** Count of deleted drafts

### `mark_as_read`
Mark messages as read.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

### `archive_emails`
Archive messages (remove INBOX label).

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

### `trash_emails`
Move messages to trash.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs

## Label Operations

### `add_labels`
Add labels to messages.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs
- `label_ids` (required): Array of label IDs

### `remove_labels`
Remove labels from messages.

**Parameters:**
- `account` (required): Account ID
- `message_ids` (required): Array of message IDs
- `label_ids` (required): Array of label IDs

### `create_label`
Create a new Gmail label.

**Parameters:**
- `account` (required): Account ID
- `name` (required): Label name
- `label_list_visibility` (optional, default: "labelShow")
- `message_list_visibility` (optional, default: "show")

### `delete_label`
Delete a Gmail label.

**Parameters:**
- `account` (required): Account ID
- `label_id` (required): Label ID to delete
