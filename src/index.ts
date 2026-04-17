#!/usr/bin/env node

import { createServer } from 'node:http';
import { promises as fs } from 'node:fs';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';
import {
  ensureConfigLayout,
  getAccountPaths,
  getConfigRoot,
  getDefaultAccountPaths,
  loadAccountsConfig,
  saveAccountsConfig,
  upsertAccount,
  validateAccountId,
  type AccountConfig,
  type AccountsConfig,
} from './config.js';
import {
  getAccountHealth,
  getAccountOrThrow,
  resolveReadAccounts,
  resolveWriteAccount,
} from './accounts.js';
import {
  GmailAccountClient,
  exchangeCodeForToken,
  generateAuthUrlFromCredentials,
  readCredentialsFile,
  type AttachmentMetadata,
  type DriveFileSummary,
  type ParsedEmail,
} from './gmail-client.js';
import { saveAndExtract, type AttachmentContent } from './attachments.js';

interface ReadEmailsArgs {
  account?: string;
  query?: string;
  max_results?: number;
  include_body?: boolean;
}

interface SearchEmailsArgs {
  account?: string;
  query: string;
  max_results?: number;
}

interface SearchDriveFilesArgs {
  account?: string;
  query: string;
  max_results?: number;
}

interface ThreadArgs {
  account: string;
  thread_id: string;
}

interface LabelsArgs {
  account: string;
}

interface MessageIdsArgs {
  account: string;
  message_ids: string[];
}

interface AddLabelsArgs extends MessageIdsArgs {
  label_ids: string[];
}

interface DeleteLabelArgs {
  account: string;
  label_id: string;
}

interface CreateLabelArgs {
  account: string;
  name: string;
  label_list_visibility?: string;
  message_list_visibility?: string;
}

interface DeleteDraftsArgs {
  account: string;
  draft_ids: string[];
}

interface OutgoingEmailArgs {
  account: string;
  to: string;
  subject: string;
  body: string;
  cc?: string;
  bcc?: string;
  html?: boolean;
  attachments?: OutgoingEmailAttachmentArgs[];
}

interface OutgoingEmailAttachmentArgs {
  path: string;
  filename?: string;
  content_type?: string;
}

interface BeginAuthArgs {
  account_id: string;
  email: string;
  display_name?: string;
  credentials_json?: unknown;
  credentials_path?: string;
}

interface FinishAuthArgs {
  account_id: string;
  authorization_code: string;
}

interface GetAttachmentArgs {
  account: string;
  email_id: string;
  attachment_id: string;
}

interface GetAllAttachmentsArgs {
  account: string;
  email_id: string;
}

type BlockAction = 'trash' | 'archive' | 'spam';

interface BlockSenderArgs {
  account: string;
  sender: string;
  action?: BlockAction;
  also_trash_existing?: boolean;
}

interface UnblockSenderArgs {
  account: string;
  filter_id?: string;
  sender?: string;
}

interface UnsubscribeArgs {
  account: string;
  message_id: string;
  dry_run?: boolean;
}

type MuteScope = 'subject' | 'thread_only';

interface MuteThreadArgs {
  account: string;
  thread_id: string;
  scope?: MuteScope;
}

function textResult(text: string): CallToolResult {
  return {
    content: [
      {
        type: 'text',
        text,
      },
    ],
  };
}

function clamp(value: number, minValue: number, maxValue: number): number {
  return Math.max(minValue, Math.min(value, maxValue));
}

function valueToString(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback;
}

function valueToBoolean(value: unknown, fallback = false): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function valueToNumber(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function valueToStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.map((item) => String(item)).map((item) => item.trim()).filter(Boolean);
}

function valueToAttachmentArray(value: unknown): OutgoingEmailAttachmentArgs[] {
  if (!Array.isArray(value)) return [];

  return value.flatMap((item) => {
    if (!item || typeof item !== 'object') return [];
    const candidate = item as Record<string, unknown>;
    const filePath = valueToString(candidate.path).trim();
    if (!filePath) return [];

    return [
      {
        path: filePath,
        filename: valueToString(candidate.filename, '').trim() || undefined,
        content_type: valueToString(candidate.content_type, '').trim() || undefined,
      },
    ];
  });
}

function emailDateForSort(email: ParsedEmail): number {
  return Number.isFinite(email.internalDate) ? email.internalDate : 0;
}

function parseListUnsubscribeHeader(raw: string): string[] {
  const urls: string[] = [];
  for (const match of raw.matchAll(/<([^>]+)>/g)) {
    const url = match[1]?.trim();
    if (url) urls.push(url);
  }
  return urls;
}

function parseMailto(mailtoUrl: string): { to: string; subject: string; body: string } {
  const stripped = mailtoUrl.replace(/^mailto:/i, '');
  const [addressPart, queryPart = ''] = stripped.split('?', 2);
  const to = (addressPart ?? '').trim();
  const params = new URLSearchParams(queryPart);
  const subject = (params.get('subject') ?? 'unsubscribe').trim() || 'unsubscribe';
  const body = (params.get('body') ?? '').trim();
  return { to, subject, body };
}

const GENERIC_SUBJECTS = new Set([
  'hi',
  'hey',
  'hello',
  'test',
  'update',
  'updates',
  'news',
  'newsletter',
  '(no subject)',
  'notification',
  'notifications',
  'alert',
  'alerts',
]);

function isGenericSubject(subject: string): boolean {
  const normalized = subject.trim().toLowerCase();
  if (normalized.length < 4) return true;
  return GENERIC_SUBJECTS.has(normalized);
}

function driveFileDateForSort(file: DriveFileSummary): number {
  if (!file.modifiedTime) return 0;
  const timestamp = Date.parse(file.modifiedTime);
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function formatEmailItem(email: ParsedEmail, includeBody: boolean): string {
  const lines = [
    `**Account**: ${email.accountId} (${email.accountEmail})`,
    `**From**: ${email.from || '(unknown)'}`,
    `**To**: ${email.to || '(unknown)'}`,
    `**Date**: ${email.date || (email.internalDate ? new Date(email.internalDate).toISOString() : '(unknown)')}`,
    `**Message ID**: ${email.id}`,
    `**Thread ID**: ${email.threadId}`,
    `**Preview**: ${email.snippet || '(none)'}`,
  ];

  if (includeBody) {
    const body = (email.body || '').trim();
    if (body) {
      const trimmed = body.length > 600 ? `${body.slice(0, 600)}...` : body;
      lines.push(`**Body**:\n${trimmed}`);
    }
  }

  if (email.attachments.length > 0) {
    lines.push(`**Attachments** (${email.attachments.length}):`);
    for (const att of email.attachments) {
      lines.push(
        `  - ${att.filename} (${att.contentType}, ${att.sizeBytes} bytes) — id: ${att.id}`,
      );
    }
  }

  return lines.join('\n');
}

function formatAttachmentContent(content: AttachmentContent): string {
  const lines = [
    `**Filename**: ${content.filename}`,
    `**Content-Type**: ${content.contentType}`,
    `**Size**: ${content.sizeBytes} bytes`,
    `**Saved to**: ${content.savedPath}`,
    `**Extraction Method**: ${content.extractionMethod}`,
  ];

  if (content.textTruncated) {
    lines.push('**Text Truncated**: true (exceeds 500 KB; saved file is complete)');
  }

  if (content.extractionError) {
    lines.push(`**Extraction Error**: ${content.extractionError}`);
  }

  if (content.text !== null) {
    lines.push('', '**Extracted Text**:', '', content.text);
  }

  return lines.join('\n');
}

function formatDriveFileItem(file: DriveFileSummary): string {
  const lines = [
    `**Account**: ${file.accountId} (${file.accountEmail})`,
    `**File ID**: ${file.id}`,
    `**MIME Type**: ${file.mimeType}`,
    `**Modified**: ${file.modifiedTime || '(unknown)'}`,
    `**Owners**: ${file.owners.length > 0 ? file.owners.join(', ') : '(unknown)'}`,
  ];

  if (file.webViewLink) {
    lines.push(`**Open**: ${file.webViewLink}`);
  }

  return lines.join('\n');
}

function formatEmailListOutput(input: {
  title: string;
  scopeText: string;
  queryText: string;
  totalFound: number;
  returned: ParsedEmail[];
  includeBody: boolean;
  errors: string[];
}): string {
  const sections: string[] = [];

  sections.push(`# ${input.title}`);
  sections.push('');
  sections.push(`**Scope**: ${input.scopeText}`);
  sections.push(`**Query**: ${input.queryText || '(none)'}`);
  sections.push(`**Total Found**: ${input.totalFound}`);
  sections.push(`**Returned**: ${input.returned.length}`);

  if (input.returned.length > 0) {
    sections.push('');
    input.returned.forEach((email, index) => {
      sections.push(`## ${index + 1}. ${email.subject || '(no subject)'}`);
      sections.push(formatEmailItem(email, input.includeBody));
      sections.push('');
    });
  }

  if (input.errors.length > 0) {
    sections.push('## Account Errors');
    sections.push(input.errors.map((error) => `- ${error}`).join('\n'));
  }

  return sections.join('\n');
}

function formatDriveFileListOutput(input: {
  title: string;
  scopeText: string;
  queryText: string;
  totalFound: number;
  returned: DriveFileSummary[];
  errors: string[];
}): string {
  const sections: string[] = [];

  sections.push(`# ${input.title}`);
  sections.push('');
  sections.push(`**Scope**: ${input.scopeText}`);
  sections.push(`**Query**: ${input.queryText}`);
  sections.push(`**Total Found**: ${input.totalFound}`);
  sections.push(`**Returned**: ${input.returned.length}`);

  if (input.returned.length > 0) {
    sections.push('');
    input.returned.forEach((file, index) => {
      sections.push(`## ${index + 1}. ${file.name}`);
      sections.push(formatDriveFileItem(file));
      sections.push('');
    });
  }

  if (input.errors.length > 0) {
    sections.push('## Account Errors');
    sections.push(input.errors.map((error) => `- ${error}`).join('\n'));
  }

  return sections.join('\n');
}

class GmailMultiInboxServer {
  private readonly server: Server;
  private readonly configRoot: string;

  constructor() {
    this.server = new Server(
      {
        name: 'gmail-multi-inbox-mcp',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.configRoot = getConfigRoot();
    this.setupHandlers();

    this.server.onerror = (error) => {
      console.error('[gmail-multi-inbox-mcp] MCP error:', error);
    };
  }

  private setupHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_accounts',
          description:
            'List all configured Gmail inbox accounts and their authentication/health status.',
          inputSchema: {
            type: 'object',
            properties: {},
            additionalProperties: false,
          },
        },
        {
          name: 'read_emails',
          description:
            'Read emails from one account or aggregate across all enabled accounts when account is omitted.',
          inputSchema: {
            type: 'object',
            properties: {
              account: {
                type: 'string',
                description: 'Optional account id. Omit to aggregate across all enabled accounts.',
              },
              query: {
                type: 'string',
                description: 'Optional Gmail query string.',
                default: '',
              },
              max_results: {
                type: 'number',
                description: 'Maximum emails to return (1-100).',
                default: 20,
              },
              include_body: {
                type: 'boolean',
                description: 'Include plaintext body extraction in each returned email.',
                default: false,
              },
            },
            additionalProperties: false,
          },
        },
        {
          name: 'search_emails',
          description:
            'Search Gmail using query syntax. Omitting account searches all enabled inboxes and merges results.',
          inputSchema: {
            type: 'object',
            properties: {
              account: {
                type: 'string',
                description: 'Optional account id. Omit to aggregate across all enabled accounts.',
              },
              query: {
                type: 'string',
                description: 'Gmail search query.',
              },
              max_results: {
                type: 'number',
                description: 'Maximum emails to return (1-100).',
                default: 25,
              },
            },
            required: ['query'],
            additionalProperties: false,
          },
        },
        {
          name: 'search_drive_files',
          description:
            'Search Google Drive file metadata. Omitting account searches all enabled accounts and merges results.',
          inputSchema: {
            type: 'object',
            properties: {
              account: {
                type: 'string',
                description: 'Optional account id. Omit to aggregate across all enabled accounts.',
              },
              query: {
                type: 'string',
                description: 'Plain text to search in Drive file names and indexed content.',
              },
              max_results: {
                type: 'number',
                description: 'Maximum files to return (1-100).',
                default: 25,
              },
            },
            required: ['query'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_email_thread',
          description: 'Get a full email thread for one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              thread_id: { type: 'string', description: 'Gmail thread id.' },
            },
            required: ['account', 'thread_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_labels',
          description: 'List labels for one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'mark_as_read',
          description: 'Mark messages as read in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Message IDs to mark as read.',
              },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'add_labels',
          description: 'Add labels to messages in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Message IDs to update.',
              },
              label_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Label IDs to add.',
              },
            },
            required: ['account', 'message_ids', 'label_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'remove_labels',
          description: 'Remove labels from messages in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Message IDs to update.',
              },
              label_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Label IDs to remove.',
              },
            },
            required: ['account', 'message_ids', 'label_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'archive_emails',
          description: 'Archive messages in one account by removing INBOX label (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Message IDs to archive.',
              },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'trash_emails',
          description: 'Move messages to trash in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Message IDs to trash.',
              },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_label',
          description: 'Create a new Gmail label in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              name: { type: 'string', description: 'Label name.' },
              label_list_visibility: {
                type: 'string',
                description: 'Gmail labelListVisibility value (default: labelShow).',
                default: 'labelShow',
              },
              message_list_visibility: {
                type: 'string',
                description: 'Gmail messageListVisibility value (default: show).',
                default: 'show',
              },
            },
            required: ['account', 'name'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_label',
          description: 'Delete a Gmail label in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              label_id: { type: 'string', description: 'Label id to delete.' },
            },
            required: ['account', 'label_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'block_sender',
          description:
            'Block a sender by creating a Gmail filter (same mechanism as Gmail\'s native "Block" button). Optionally also trashes existing mail from that sender.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              sender: {
                type: 'string',
                description:
                  'Email address (e.g. "spam@foo.com"), domain ("@foo.com"), or Gmail search fragment ("from:x subject:promo").',
              },
              action: {
                type: 'string',
                enum: ['trash', 'archive', 'spam'],
                description:
                  'What to do with matching mail. "trash" (default) mirrors Gmail\'s Block button.',
                default: 'trash',
              },
              also_trash_existing: {
                type: 'boolean',
                description:
                  'If true (default), retroactively trashes up to 100 existing messages from this sender.',
                default: true,
              },
            },
            required: ['account', 'sender'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_blocked_senders',
          description:
            'List all Gmail filters for an account. Use to find filter_id for unblock_sender, or to audit what is being auto-trashed/archived/spammed.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'unblock_sender',
          description:
            'Remove a Gmail filter. Specify either filter_id (exact match) or sender (matches filters whose "from" criteria equals this value).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              filter_id: {
                type: 'string',
                description: 'Filter id from list_blocked_senders. Prefer this when known.',
              },
              sender: {
                type: 'string',
                description:
                  'Sender string to match against filters\' from criteria. Used when filter_id is not provided.',
              },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'unsubscribe_from_email',
          description:
            'Unsubscribe from a mailing list by invoking the List-Unsubscribe header (RFC 2369/8058). Same mechanism as Gmail\'s native "Unsubscribe" button. Falls back to mailto:. Returns method used.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_id: {
                type: 'string',
                description: 'Gmail message id of a message from the sender to unsubscribe from.',
              },
              dry_run: {
                type: 'boolean',
                description:
                  'If true, report what method would be used without executing the unsubscribe. Default false.',
                default: false,
              },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'mute_thread',
          description:
            'Mute a conversation: archives the thread now and (with scope="subject") installs a subject-based filter so future replies auto-archive. Approximates Gmail\'s client-side mute.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              thread_id: {
                type: 'string',
                description: 'Thread id (from search_emails / read_emails / get_email_thread).',
              },
              scope: {
                type: 'string',
                enum: ['subject', 'thread_only'],
                description:
                  '"subject" (default) archives now + auto-archives future replies with the same subject. "thread_only" just archives this thread without creating a filter.',
                default: 'subject',
              },
            },
            required: ['account', 'thread_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_draft',
          description: 'Create a draft email in one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              to: { type: 'string', description: 'Recipient email address(es).' },
              subject: { type: 'string', description: 'Email subject.' },
              body: { type: 'string', description: 'Email body.' },
              cc: { type: 'string', description: 'Optional CC list.' },
              bcc: { type: 'string', description: 'Optional BCC list.' },
              html: {
                type: 'boolean',
                description: 'Set true to send body as text/html.',
                default: false,
              },
              attachments: {
                type: 'array',
                description: 'Optional local file attachments.',
                items: {
                  type: 'object',
                  properties: {
                    path: { type: 'string', description: 'Absolute or local filesystem path.' },
                    filename: {
                      type: 'string',
                      description: 'Optional override filename shown in Gmail.',
                    },
                    content_type: {
                      type: 'string',
                      description: 'Optional MIME type override (for example application/pdf).',
                    },
                  },
                  required: ['path'],
                  additionalProperties: false,
                },
              },
            },
            required: ['account', 'to', 'subject', 'body'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_drafts',
          description: 'Permanently delete one or more drafts in one account (account required). Draft IDs are returned by create_draft.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              draft_ids: {
                type: 'array',
                items: { type: 'string' },
                description: 'Draft IDs to delete (as returned by create_draft).',
              },
            },
            required: ['account', 'draft_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'send_email',
          description: 'Send an email from one account (account required).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              to: { type: 'string', description: 'Recipient email address(es).' },
              subject: { type: 'string', description: 'Email subject.' },
              body: { type: 'string', description: 'Email body.' },
              cc: { type: 'string', description: 'Optional CC list.' },
              bcc: { type: 'string', description: 'Optional BCC list.' },
              html: {
                type: 'boolean',
                description: 'Set true to send body as text/html.',
                default: false,
              },
              attachments: {
                type: 'array',
                description: 'Optional local file attachments.',
                items: {
                  type: 'object',
                  properties: {
                    path: { type: 'string', description: 'Absolute or local filesystem path.' },
                    filename: {
                      type: 'string',
                      description: 'Optional override filename shown in Gmail.',
                    },
                    content_type: {
                      type: 'string',
                      description: 'Optional MIME type override (for example application/pdf).',
                    },
                  },
                  required: ['path'],
                  additionalProperties: false,
                },
              },
            },
            required: ['account', 'to', 'subject', 'body'],
            additionalProperties: false,
          },
        },
        {
          name: 'begin_account_auth',
          description:
            'Start OAuth onboarding for an account. Accepts credentials JSON or a path to credentials.json.',
          inputSchema: {
            type: 'object',
            properties: {
              account_id: {
                type: 'string',
                description: 'Stable id for this inbox account (letters/numbers/_/- only).',
              },
              email: {
                type: 'string',
                description: 'Email address for this account.',
              },
              display_name: {
                type: 'string',
                description: 'Optional display name (e.g., Personal, Work).',
              },
              credentials_json: {
                description: 'OAuth client JSON object or JSON string from Google Cloud.',
              },
              credentials_path: {
                type: 'string',
                description: 'Path to an existing credentials.json file.',
              },
            },
            required: ['account_id', 'email'],
            additionalProperties: false,
          },
        },
        {
          name: 'finish_account_auth',
          description:
            'Complete OAuth onboarding by exchanging authorization code and storing token.json.',
          inputSchema: {
            type: 'object',
            properties: {
              account_id: {
                type: 'string',
                description: 'Account id used in begin_account_auth.',
              },
              authorization_code: {
                type: 'string',
                description: 'OAuth authorization code from Google redirect.',
              },
            },
            required: ['account_id', 'authorization_code'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_attachment',
          description:
            'Fetch a single email attachment by id. Saves to ~/Downloads/mcp-attachments/ and returns extracted text for supported formats (PDF, DOCX, XLSX, PPTX, text, images via OCR).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              email_id: { type: 'string', description: 'Gmail message id.' },
              attachment_id: {
                type: 'string',
                description: 'Attachment id from read_emails / get_email_thread metadata.',
              },
            },
            required: ['account', 'email_id', 'attachment_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_all_attachments',
          description:
            'Fetch every attachment on an email in one call. Saves each to disk and returns extracted text for each.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              email_id: { type: 'string', description: 'Gmail message id.' },
            },
            required: ['account', 'email_id'],
            additionalProperties: false,
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const name = request.params.name;
      const args = (request.params.arguments ?? {}) as Record<string, unknown>;

      try {
        switch (name) {
          case 'list_accounts':
            return await this.handleListAccounts();
          case 'read_emails':
            return await this.handleReadEmails(args);
          case 'search_emails':
            return await this.handleSearchEmails(args);
          case 'search_drive_files':
            return await this.handleSearchDriveFiles(args);
          case 'get_email_thread':
            return await this.handleGetThread(args);
          case 'get_labels':
            return await this.handleGetLabels(args);
          case 'mark_as_read':
            return await this.handleMarkAsRead(args);
          case 'add_labels':
            return await this.handleAddLabels(args);
          case 'remove_labels':
            return await this.handleRemoveLabels(args);
          case 'archive_emails':
            return await this.handleArchiveEmails(args);
          case 'trash_emails':
            return await this.handleTrashEmails(args);
          case 'create_label':
            return await this.handleCreateLabel(args);
          case 'delete_label':
            return await this.handleDeleteLabel(args);
          case 'block_sender':
            return await this.handleBlockSender(args);
          case 'list_blocked_senders':
            return await this.handleListBlockedSenders(args);
          case 'unblock_sender':
            return await this.handleUnblockSender(args);
          case 'unsubscribe_from_email':
            return await this.handleUnsubscribeFromEmail(args);
          case 'mute_thread':
            return await this.handleMuteThread(args);
          case 'create_draft':
            return await this.handleCreateDraft(args);
          case 'delete_drafts':
            return await this.handleDeleteDrafts(args);
          case 'send_email':
            return await this.handleSendEmail(args);
          case 'begin_account_auth':
            return await this.handleBeginAccountAuth(args);
          case 'finish_account_auth':
            return await this.handleFinishAccountAuth(args);
          case 'get_attachment':
            return await this.handleGetAttachment(args);
          case 'get_all_attachments':
            return await this.handleGetAllAttachments(args);
          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        return textResult(`Error executing ${name}: ${(error as Error).message}`);
      }
    });
  }

  private async loadConfig(): Promise<AccountsConfig> {
    return loadAccountsConfig(this.configRoot);
  }

  private async getClientForAccount(account: AccountConfig): Promise<GmailAccountClient> {
    return GmailAccountClient.create(this.configRoot, account);
  }

  private async handleListAccounts(): Promise<CallToolResult> {
    const config = await this.loadConfig();

    if (config.accounts.length === 0) {
      return textResult(
        [
          '# Gmail Accounts',
          '',
          'No accounts configured yet.',
          '',
          'Use `begin_account_auth` then `finish_account_auth` to add the first inbox.',
        ].join('\n')
      );
    }

    const healthList = await Promise.all(
      config.accounts.map((account) => getAccountHealth(this.configRoot, account))
    );

    const lines: string[] = ['# Gmail Accounts', '', `**Config Root**: ${this.configRoot}`, ''];

    healthList.forEach((health, index) => {
      const { account } = health;
      const status = health.ready ? 'ready' : account.enabled ? 'needs-auth-files' : 'disabled';
      const defaultMarker = config.defaultAccount === account.id ? ' (default)' : '';
      lines.push(`## ${index + 1}. ${account.id}${defaultMarker}`);
      lines.push(`- Email: ${account.email}`);
      lines.push(`- Display Name: ${account.displayName || '(none)'}`);
      lines.push(`- Enabled: ${account.enabled}`);
      lines.push(`- Status: ${status}`);
      lines.push(`- Credentials File: ${health.hasCredentialsFile ? 'present' : 'missing'}`);
      lines.push(`- Token File: ${health.hasTokenFile ? 'present' : 'missing'}`);
      lines.push('');
    });

    return textResult(lines.join('\n'));
  }

  private async handleReadEmails(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: ReadEmailsArgs = {
      account: valueToString(rawArgs.account, '').trim() || undefined,
      query: valueToString(rawArgs.query, ''),
      max_results: valueToNumber(rawArgs.max_results, 20),
      include_body: valueToBoolean(rawArgs.include_body, false),
    };

    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, args.account);
    const query = args.query?.trim() ?? '';
    const includeBody = args.include_body ?? false;
    const maxResults = clamp(args.max_results ?? 20, 1, 100);

    const accountResults = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const emails = await client.readEmails(query, maxResults, includeBody);
          return { account, emails, error: null as string | null };
        } catch (error) {
          return {
            account,
            emails: [] as ParsedEmail[],
            error: `${account.id}: ${(error as Error).message}`,
          };
        }
      })
    );

    const merged = accountResults
      .flatMap((result) => result.emails)
      .sort((a, b) => emailDateForSort(b) - emailDateForSort(a));

    const errors = accountResults
      .map((result) => result.error)
      .filter((error): error is string => Boolean(error));

    const returned = merged.slice(0, maxResults);

    return textResult(
      formatEmailListOutput({
        title: 'Gmail Emails',
        scopeText: args.account
          ? `account ${args.account}`
          : `all enabled accounts (${targetAccounts.length})`,
        queryText: query,
        totalFound: merged.length,
        returned,
        includeBody,
        errors,
      })
    );
  }

  private async handleSearchEmails(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: SearchEmailsArgs = {
      account: valueToString(rawArgs.account, '').trim() || undefined,
      query: valueToString(rawArgs.query, ''),
      max_results: valueToNumber(rawArgs.max_results, 25),
    };

    if (!args.query || args.query.trim() === '') {
      throw new Error('query is required.');
    }

    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, args.account);
    const maxResults = clamp(args.max_results ?? 25, 1, 100);

    const accountResults = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const emails = await client.searchEmails(args.query, maxResults);
          return { account, emails, error: null as string | null };
        } catch (error) {
          return {
            account,
            emails: [] as ParsedEmail[],
            error: `${account.id}: ${(error as Error).message}`,
          };
        }
      })
    );

    const merged = accountResults
      .flatMap((result) => result.emails)
      .sort((a, b) => emailDateForSort(b) - emailDateForSort(a));

    const errors = accountResults
      .map((result) => result.error)
      .filter((error): error is string => Boolean(error));

    const returned = merged.slice(0, maxResults);

    return textResult(
      formatEmailListOutput({
        title: 'Gmail Search Results',
        scopeText: args.account
          ? `account ${args.account}`
          : `all enabled accounts (${targetAccounts.length})`,
        queryText: args.query,
        totalFound: merged.length,
        returned,
        includeBody: false,
        errors,
      })
    );
  }

  private async handleSearchDriveFiles(
    rawArgs: Record<string, unknown>
  ): Promise<CallToolResult> {
    const args: SearchDriveFilesArgs = {
      account: valueToString(rawArgs.account, '').trim() || undefined,
      query: valueToString(rawArgs.query, ''),
      max_results: valueToNumber(rawArgs.max_results, 25),
    };

    if (!args.query || args.query.trim() === '') {
      throw new Error('query is required.');
    }

    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, args.account);
    const maxResults = clamp(args.max_results ?? 25, 1, 100);

    const accountResults = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const files = await client.searchDriveFiles(args.query, maxResults);
          return { account, files, error: null as string | null };
        } catch (error) {
          return {
            account,
            files: [] as DriveFileSummary[],
            error: `${account.id}: ${(error as Error).message}`,
          };
        }
      })
    );

    const merged = accountResults
      .flatMap((result) => result.files)
      .sort((a, b) => driveFileDateForSort(b) - driveFileDateForSort(a));

    const errors = accountResults
      .map((result) => result.error)
      .filter((error): error is string => Boolean(error));

    const returned = merged.slice(0, maxResults);

    return textResult(
      formatDriveFileListOutput({
        title: 'Google Drive Search Results',
        scopeText: args.account
          ? `account ${args.account}`
          : `all enabled accounts (${targetAccounts.length})`,
        queryText: args.query,
        totalFound: merged.length,
        returned,
        errors,
      })
    );
  }

  private async handleGetThread(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: ThreadArgs = {
      account: valueToString(rawArgs.account),
      thread_id: valueToString(rawArgs.thread_id),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const thread = await client.getThread(args.thread_id);

    const lines = [
      '# Gmail Thread',
      '',
      `**Account**: ${account.id} (${account.email})`,
      `**Thread ID**: ${thread.threadId}`,
      `**Messages**: ${thread.messages.length}`,
      '',
    ];

    thread.messages.forEach((message, index) => {
      lines.push(`## ${index + 1}. ${message.subject}`);
      lines.push(formatEmailItem(message, true));
      lines.push('');
    });

    return textResult(lines.join('\n'));
  }

  private async handleGetLabels(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: LabelsArgs = {
      account: valueToString(rawArgs.account),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const labels = await client.getLabels();

    const lines = [
      '# Gmail Labels',
      '',
      `**Account**: ${account.id} (${account.email})`,
      `**Count**: ${labels.length}`,
      '',
    ];

    labels.forEach((label, index) => {
      lines.push(
        `${index + 1}. ${label.name} | id=${label.id} | type=${label.type ?? 'unknown'} | messages=${label.messagesTotal ?? 0}`
      );
    });

    return textResult(lines.join('\n'));
  }

  private async handleMarkAsRead(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: MessageIdsArgs = {
      account: valueToString(rawArgs.account),
      message_ids: valueToStringArray(rawArgs.message_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const updated = await client.markAsRead(args.message_ids);

    return textResult(`✅ Marked ${updated} message(s) as read in account ${account.id}.`);
  }

  private async handleAddLabels(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: AddLabelsArgs = {
      account: valueToString(rawArgs.account),
      message_ids: valueToStringArray(rawArgs.message_ids),
      label_ids: valueToStringArray(rawArgs.label_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const updated = await client.addLabels(args.message_ids, args.label_ids);

    return textResult(
      `✅ Added label(s) to ${updated} message(s) in account ${account.id}. Labels: ${args.label_ids.join(', ')}`
    );
  }

  private async handleRemoveLabels(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: AddLabelsArgs = {
      account: valueToString(rawArgs.account),
      message_ids: valueToStringArray(rawArgs.message_ids),
      label_ids: valueToStringArray(rawArgs.label_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const updated = await client.removeLabels(args.message_ids, args.label_ids);

    return textResult(
      `✅ Removed label(s) from ${updated} message(s) in account ${account.id}. Labels: ${args.label_ids.join(', ')}`
    );
  }

  private async handleArchiveEmails(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: MessageIdsArgs = {
      account: valueToString(rawArgs.account),
      message_ids: valueToStringArray(rawArgs.message_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const updated = await client.archiveEmails(args.message_ids);

    return textResult(`✅ Archived ${updated} message(s) in account ${account.id}.`);
  }

  private async handleTrashEmails(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: MessageIdsArgs = {
      account: valueToString(rawArgs.account),
      message_ids: valueToStringArray(rawArgs.message_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const updated = await client.trashEmails(args.message_ids);

    return textResult(`✅ Trashed ${updated} message(s) in account ${account.id}.`);
  }

  private async handleCreateLabel(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: CreateLabelArgs = {
      account: valueToString(rawArgs.account),
      name: valueToString(rawArgs.name),
      label_list_visibility: valueToString(rawArgs.label_list_visibility, 'labelShow'),
      message_list_visibility: valueToString(rawArgs.message_list_visibility, 'show'),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const label = await client.createLabel(
      args.name,
      args.label_list_visibility,
      args.message_list_visibility
    );

    return textResult(
      `✅ Created label in account ${account.id}: ${label.name} (id=${label.id}, type=${label.type ?? 'unknown'})`
    );
  }

  private async handleDeleteLabel(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: DeleteLabelArgs = {
      account: valueToString(rawArgs.account),
      label_id: valueToString(rawArgs.label_id),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    await client.deleteLabel(args.label_id);

    return textResult(`✅ Deleted label ${args.label_id} in account ${account.id}.`);
  }

  private async handleBlockSender(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const rawAction = valueToString(rawArgs.action, 'trash').trim().toLowerCase();
    const action: BlockAction =
      rawAction === 'archive' || rawAction === 'spam' ? rawAction : 'trash';

    const args: BlockSenderArgs = {
      account: valueToString(rawArgs.account),
      sender: valueToString(rawArgs.sender).trim(),
      action,
      also_trash_existing: valueToBoolean(rawArgs.also_trash_existing, true),
    };

    if (!args.sender) throw new Error('sender is required.');

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const filter = await client.createBlockFilter(args.sender, action);

    let sweptCount = 0;
    let sweepError: string | null = null;
    if (args.also_trash_existing && action !== 'archive') {
      try {
        const existing = await client.searchEmails(`from:${args.sender}`, 100);
        if (existing.length > 0) {
          sweptCount = await client.trashEmails(existing.map((email) => email.id));
        }
      } catch (error) {
        sweepError = (error as Error).message;
      }
    }

    const lines = [
      `✅ Blocked \`${args.sender}\` in account ${account.id}.`,
      `- Action: ${action}`,
      `- Filter id: ${filter.id ?? '(unknown)'}`,
    ];
    if (args.also_trash_existing && action !== 'archive') {
      lines.push(`- Existing messages trashed: ${sweptCount}`);
      if (sweepError) {
        lines.push(`- Sweep warning: ${sweepError}`);
      }
    }
    if (action === 'archive' && args.also_trash_existing) {
      lines.push('- Note: `also_trash_existing` ignored for `archive` action (would be destructive).');
    }

    return textResult(lines.join('\n'));
  }

  private async handleListBlockedSenders(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args = { account: valueToString(rawArgs.account) };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const filters = await client.listFilters();

    if (filters.length === 0) {
      return textResult(`# Filters for account ${account.id}\n\nNo filters configured.`);
    }

    const lines: string[] = [`# Filters for account ${account.id}`, '', `**Total**: ${filters.length}`, ''];

    filters.forEach((filter, index) => {
      const c = filter.criteria ?? {};
      const a = filter.action ?? {};
      const criteriaParts: string[] = [];
      if (c.from) criteriaParts.push(`from: ${c.from}`);
      if (c.to) criteriaParts.push(`to: ${c.to}`);
      if (c.subject) criteriaParts.push(`subject: ${c.subject}`);
      if (c.query) criteriaParts.push(`query: ${c.query}`);
      if (c.hasAttachment) criteriaParts.push('hasAttachment: true');

      const actionParts: string[] = [];
      if (a.addLabelIds?.length) actionParts.push(`add: ${a.addLabelIds.join(', ')}`);
      if (a.removeLabelIds?.length) actionParts.push(`remove: ${a.removeLabelIds.join(', ')}`);
      if (a.forward) actionParts.push(`forward: ${a.forward}`);

      lines.push(`## ${index + 1}. Filter \`${filter.id ?? '(no id)'}\``);
      lines.push(`- Criteria: ${criteriaParts.length > 0 ? criteriaParts.join(' | ') : '(none)'}`);
      lines.push(`- Action: ${actionParts.length > 0 ? actionParts.join(' | ') : '(none)'}`);
      lines.push('');
    });

    return textResult(lines.join('\n'));
  }

  private async handleUnblockSender(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: UnblockSenderArgs = {
      account: valueToString(rawArgs.account),
      filter_id: valueToString(rawArgs.filter_id, '').trim() || undefined,
      sender: valueToString(rawArgs.sender, '').trim() || undefined,
    };

    if (!args.filter_id && !args.sender) {
      throw new Error('Provide either filter_id or sender.');
    }

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    if (args.filter_id) {
      await client.deleteFilter(args.filter_id);
      return textResult(`✅ Deleted filter ${args.filter_id} in account ${account.id}.`);
    }

    const senderNeedle = args.sender!;
    const filters = await client.listFilters();
    const matches = filters.filter((f) => (f.criteria?.from ?? '').trim() === senderNeedle);

    if (matches.length === 0) {
      return textResult(
        `No filters in account ${account.id} matched sender \`${senderNeedle}\`. Use list_blocked_senders to inspect.`
      );
    }

    const deletedIds: string[] = [];
    for (const match of matches) {
      if (!match.id) continue;
      await client.deleteFilter(match.id);
      deletedIds.push(match.id);
    }

    return textResult(
      `✅ Deleted ${deletedIds.length} filter(s) matching \`${senderNeedle}\` in account ${account.id}: ${deletedIds.join(', ')}.`
    );
  }

  private async handleUnsubscribeFromEmail(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: UnsubscribeArgs = {
      account: valueToString(rawArgs.account),
      message_id: valueToString(rawArgs.message_id).trim(),
      dry_run: valueToBoolean(rawArgs.dry_run, false),
    };

    if (!args.message_id) throw new Error('message_id is required.');

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const headers = await client.getMessageHeaders(args.message_id, [
      'List-Unsubscribe',
      'List-Unsubscribe-Post',
    ]);

    const listUnsubscribe = headers['List-Unsubscribe']?.trim() ?? '';
    const listUnsubscribePost = headers['List-Unsubscribe-Post']?.trim() ?? '';

    if (!listUnsubscribe) {
      return textResult(
        [
          `# Unsubscribe: unavailable`,
          ``,
          `Message ${args.message_id} has no \`List-Unsubscribe\` header. Nothing to unsubscribe from.`,
        ].join('\n')
      );
    }

    const urls = parseListUnsubscribeHeader(listUnsubscribe);
    const httpsUrl = urls.find((u) => u.startsWith('https://') || u.startsWith('http://'));
    const mailtoUrl = urls.find((u) => u.startsWith('mailto:'));

    const isOneClick = /List-Unsubscribe\s*=\s*One-Click/i.test(listUnsubscribePost);

    if (httpsUrl && isOneClick) {
      if (args.dry_run) {
        return textResult(
          `[dry_run] Would POST one-click unsubscribe to ${httpsUrl}`
        );
      }
      try {
        const response = await fetch(httpsUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: 'List-Unsubscribe=One-Click',
        });
        const success = response.status >= 200 && response.status < 300;
        return textResult(
          [
            `# Unsubscribe: ${success ? '✅ success' : '⚠️ non-2xx response'}`,
            ``,
            `- Method: http_one_click (RFC 8058)`,
            `- URL: ${httpsUrl}`,
            `- HTTP status: ${response.status} ${response.statusText}`,
          ].join('\n')
        );
      } catch (error) {
        return textResult(
          [
            `# Unsubscribe: ⚠️ network error`,
            ``,
            `- Method: http_one_click`,
            `- URL: ${httpsUrl}`,
            `- Error: ${(error as Error).message}`,
          ].join('\n')
        );
      }
    }

    if (mailtoUrl) {
      const mailto = parseMailto(mailtoUrl);
      if (args.dry_run) {
        return textResult(
          `[dry_run] Would send mailto unsubscribe to ${mailto.to} (subject: "${mailto.subject}").`
        );
      }
      try {
        const result = await client.sendEmail({
          to: mailto.to,
          subject: mailto.subject,
          body: mailto.body,
          html: false,
        });
        return textResult(
          [
            `# Unsubscribe: ✅ sent`,
            ``,
            `- Method: mailto`,
            `- To: ${mailto.to}`,
            `- Subject: ${mailto.subject}`,
            `- Sent message id: ${result.messageId}`,
          ].join('\n')
        );
      } catch (error) {
        return textResult(
          [
            `# Unsubscribe: ⚠️ mailto send failed`,
            ``,
            `- To: ${mailto.to}`,
            `- Error: ${(error as Error).message}`,
          ].join('\n')
        );
      }
    }

    if (httpsUrl) {
      return textResult(
        [
          `# Unsubscribe: manual confirmation required`,
          ``,
          `- Method: http_manual (no one-click header)`,
          `- URL: ${httpsUrl}`,
          ``,
          `This sender provides an HTTPS unsubscribe URL but did not opt in to RFC 8058 one-click. Not auto-clicking (some senders count the click as consent). Open the URL manually if you trust it.`,
        ].join('\n')
      );
    }

    return textResult(
      [
        `# Unsubscribe: unavailable`,
        ``,
        `Header present but no parseable URL found: \`${listUnsubscribe}\``,
      ].join('\n')
    );
  }

  private async handleMuteThread(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const rawScope = valueToString(rawArgs.scope, 'subject').trim().toLowerCase();
    const scope: MuteScope = rawScope === 'thread_only' ? 'thread_only' : 'subject';

    const args: MuteThreadArgs = {
      account: valueToString(rawArgs.account),
      thread_id: valueToString(rawArgs.thread_id).trim(),
      scope,
    };

    if (!args.thread_id) throw new Error('thread_id is required.');

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    await client.modifyThread(args.thread_id, {
      removeLabelIds: ['INBOX'],
    });

    const result: {
      thread_archived: boolean;
      filter_created: boolean;
      filter_id?: string;
      subject_matched?: string;
      warning?: string;
    } = {
      thread_archived: true,
      filter_created: false,
    };

    if (scope === 'subject') {
      const subject = await client.getThreadSubject(args.thread_id);
      const cleaned = subject.trim();

      if (isGenericSubject(cleaned)) {
        result.warning = `Subject "${cleaned}" is too generic; skipped filter to avoid collateral muting. Thread was archived only.`;
      } else {
        const filter = await client.createFilter(
          { subject: cleaned },
          { removeLabelIds: ['INBOX'] }
        );
        result.filter_created = true;
        result.filter_id = filter.id ?? undefined;
        result.subject_matched = cleaned;
      }
    }

    const lines = [
      `✅ Muted thread ${args.thread_id} in account ${account.id}.`,
      `- Thread archived: yes`,
      `- Filter created: ${result.filter_created ? 'yes' : 'no'}`,
    ];
    if (result.filter_id) lines.push(`- Filter id: ${result.filter_id}`);
    if (result.subject_matched) lines.push(`- Subject matched: "${result.subject_matched}"`);
    if (result.warning) lines.push(`- Warning: ${result.warning}`);

    return textResult(lines.join('\n'));
  }

  private async handleCreateDraft(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: OutgoingEmailArgs = {
      account: valueToString(rawArgs.account),
      to: valueToString(rawArgs.to),
      subject: valueToString(rawArgs.subject),
      body: valueToString(rawArgs.body),
      cc: valueToString(rawArgs.cc, '') || undefined,
      bcc: valueToString(rawArgs.bcc, '') || undefined,
      html: valueToBoolean(rawArgs.html, false),
      attachments: valueToAttachmentArray(rawArgs.attachments),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const result = await client.createDraft({
      to: args.to,
      subject: args.subject,
      body: args.body,
      cc: args.cc,
      bcc: args.bcc,
      html: args.html,
      attachments: args.attachments?.map((attachment) => ({
        path: attachment.path,
        filename: attachment.filename,
        contentType: attachment.content_type,
      })),
    });

    return textResult(
      [
        '✅ Draft created.',
        `Account: ${account.id} (${account.email})`,
        `Draft ID: ${result.draftId || '(unknown)'}`,
        `Thread ID: ${result.threadId || '(unknown)'}`,
        `Attachments: ${args.attachments?.length ?? 0}`,
      ].join('\n')
    );
  }

  private async handleDeleteDrafts(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: DeleteDraftsArgs = {
      account: valueToString(rawArgs.account),
      draft_ids: valueToStringArray(rawArgs.draft_ids),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);
    const deleted = await client.deleteDrafts(args.draft_ids);

    return textResult(`✅ Deleted ${deleted} draft(s) in account ${account.id}.`);
  }

  private async handleSendEmail(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: OutgoingEmailArgs = {
      account: valueToString(rawArgs.account),
      to: valueToString(rawArgs.to),
      subject: valueToString(rawArgs.subject),
      body: valueToString(rawArgs.body),
      cc: valueToString(rawArgs.cc, '') || undefined,
      bcc: valueToString(rawArgs.bcc, '') || undefined,
      html: valueToBoolean(rawArgs.html, false),
      attachments: valueToAttachmentArray(rawArgs.attachments),
    };

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const result = await client.sendEmail({
      to: args.to,
      subject: args.subject,
      body: args.body,
      cc: args.cc,
      bcc: args.bcc,
      html: args.html,
      attachments: args.attachments?.map((attachment) => ({
        path: attachment.path,
        filename: attachment.filename,
        contentType: attachment.content_type,
      })),
    });

    return textResult(
      [
        '✅ Email sent.',
        `Account: ${account.id} (${account.email})`,
        `Message ID: ${result.messageId || '(unknown)'}`,
        `Thread ID: ${result.threadId || '(unknown)'}`,
        `Attachments: ${args.attachments?.length ?? 0}`,
      ].join('\n')
    );
  }

  private async parseCredentialsInput(args: BeginAuthArgs, credentialsPath: string): Promise<unknown> {
    if (args.credentials_json !== undefined) {
      if (typeof args.credentials_json === 'string') {
        return JSON.parse(args.credentials_json);
      }
      if (typeof args.credentials_json === 'object' && args.credentials_json !== null) {
        return args.credentials_json;
      }
      throw new Error('credentials_json must be either a JSON string or object.');
    }

    if (args.credentials_path && args.credentials_path.trim() !== '') {
      const raw = await fs.readFile(args.credentials_path, 'utf8');
      return JSON.parse(raw);
    }

    return readCredentialsFile(credentialsPath);
  }

  private async handleBeginAccountAuth(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: BeginAuthArgs = {
      account_id: valueToString(rawArgs.account_id),
      email: valueToString(rawArgs.email),
      display_name: valueToString(rawArgs.display_name, '') || undefined,
      credentials_json: rawArgs.credentials_json,
      credentials_path: valueToString(rawArgs.credentials_path, '') || undefined,
    };

    if (!args.account_id) throw new Error('account_id is required.');
    if (!args.email) throw new Error('email is required.');

    validateAccountId(args.account_id);
    await ensureConfigLayout(this.configRoot);

    const defaultPaths = getDefaultAccountPaths(this.configRoot, args.account_id);
    await fs.mkdir(defaultPaths.accountDir, { recursive: true });

    const credentials = await this.parseCredentialsInput(args, defaultPaths.credentialsPath);
    await fs.writeFile(defaultPaths.credentialsPath, `${JSON.stringify(credentials, null, 2)}\n`, 'utf8');

    const { authUrl } = generateAuthUrlFromCredentials(credentials);

    let config = await this.loadConfig();

    config = upsertAccount(config, {
      id: args.account_id,
      email: args.email,
      displayName: args.display_name,
      enabled: false,
      credentialPath: defaultPaths.credentialsPath,
      tokenPath: defaultPaths.tokenPath,
    });

    await saveAccountsConfig(this.configRoot, config);

    return textResult(
      [
        '# Google Account OAuth Started',
        '',
        `**Account ID**: ${args.account_id}`,
        `**Email**: ${args.email}`,
        `**Credentials File**: ${defaultPaths.credentialsPath}`,
        '',
        'Open this URL and approve access for Gmail and Google Drive metadata:',
        authUrl,
        '',
        'Then run `finish_account_auth` with the same account_id and the returned authorization_code.',
      ].join('\n')
    );
  }

  private async handleFinishAccountAuth(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const args: FinishAuthArgs = {
      account_id: valueToString(rawArgs.account_id),
      authorization_code: valueToString(rawArgs.authorization_code),
    };

    if (!args.account_id) throw new Error('account_id is required.');
    if (!args.authorization_code) throw new Error('authorization_code is required.');

    validateAccountId(args.account_id);

    let config = await this.loadConfig();
    const account = getAccountOrThrow(config, args.account_id);
    const paths = getAccountPaths(this.configRoot, account);

    const credentials = await readCredentialsFile(paths.credentialsPath);
    const tokens = await exchangeCodeForToken(credentials, args.authorization_code);

    if (!tokens.access_token && !tokens.refresh_token) {
      throw new Error('OAuth exchange succeeded but no token payload was returned.');
    }

    await fs.writeFile(paths.tokenPath, `${JSON.stringify(tokens, null, 2)}\n`, 'utf8');

    const updatedAccount: AccountConfig = {
      ...account,
      enabled: true,
      credentialPath: paths.credentialsPath,
      tokenPath: paths.tokenPath,
    };

    const tempConfig = upsertAccount(config, updatedAccount);
    await saveAccountsConfig(this.configRoot, tempConfig);

    const refreshedAccount = getAccountOrThrow(tempConfig, args.account_id);
    const client = await this.getClientForAccount(refreshedAccount);

    let profileEmail = refreshedAccount.email;
    try {
      profileEmail = await client.getProfileEmail();
    } catch {
      profileEmail = refreshedAccount.email;
    }

    config = upsertAccount(tempConfig, {
      ...refreshedAccount,
      email: profileEmail,
      enabled: true,
    });

    if (!config.defaultAccount) {
      config.defaultAccount = args.account_id;
    }

    await saveAccountsConfig(this.configRoot, config);

    return textResult(
      [
        '✅ Google account OAuth completed.',
        `Account ID: ${args.account_id}`,
        `Email: ${profileEmail}`,
        `Token File: ${paths.tokenPath}`,
        '',
        'This account is now enabled for Gmail tools and Google Drive metadata search.',
      ].join('\n')
    );
  }

  private async handleGetAttachment(
    rawArgs: Record<string, unknown>,
  ): Promise<CallToolResult> {
    const args: GetAttachmentArgs = {
      account: valueToString(rawArgs.account),
      email_id: valueToString(rawArgs.email_id),
      attachment_id: valueToString(rawArgs.attachment_id),
    };

    if (!args.account || !args.email_id || !args.attachment_id) {
      throw new Error('account, email_id, and attachment_id are required.');
    }

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const { bytes, metadata } = await client.getAttachment(
      args.email_id,
      args.attachment_id,
    );
    const content = await saveAndExtract(bytes, metadata);

    return textResult(
      [
        '# Gmail Attachment',
        '',
        `**Account**: ${account.id} (${account.email})`,
        `**Message ID**: ${args.email_id}`,
        `**Attachment ID**: ${args.attachment_id}`,
        '',
        formatAttachmentContent(content),
      ].join('\n'),
    );
  }

  private async handleGetAllAttachments(
    rawArgs: Record<string, unknown>,
  ): Promise<CallToolResult> {
    const args: GetAllAttachmentsArgs = {
      account: valueToString(rawArgs.account),
      email_id: valueToString(rawArgs.email_id),
    };

    if (!args.account || !args.email_id) {
      throw new Error('account and email_id are required.');
    }

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, args.account);
    const client = await this.getClientForAccount(account);

    const list: AttachmentMetadata[] = await client.listAttachments(args.email_id);

    const settled = await Promise.allSettled(
      list.map(async (meta) => {
        const { bytes, metadata } = await client.getAttachment(args.email_id, meta.id);
        return saveAndExtract(bytes, metadata);
      }),
    );

    const attachments: AttachmentContent[] = [];
    const errors: Array<{ attachment_id: string; error: string }> = [];

    settled.forEach((result, index) => {
      const originalId = list[index]?.id ?? '(unknown)';
      if (result.status === 'fulfilled') {
        attachments.push(result.value);
      } else {
        errors.push({
          attachment_id: originalId,
          error:
            result.reason instanceof Error
              ? result.reason.message
              : String(result.reason),
        });
      }
    });

    const lines = [
      '# Gmail Attachments',
      '',
      `**Account**: ${account.id} (${account.email})`,
      `**Message ID**: ${args.email_id}`,
      `**Attachments**: ${attachments.length}`,
      `**Errors**: ${errors.length}`,
      '',
    ];

    attachments.forEach((att, index) => {
      lines.push(`## ${index + 1}. ${att.filename}`);
      lines.push(formatAttachmentContent(att));
      lines.push('');
    });

    if (errors.length > 0) {
      lines.push('## Errors');
      errors.forEach((err) => {
        lines.push(`- ${err.attachment_id}: ${err.error}`);
      });
    }

    return textResult(lines.join('\n'));
  }

  async connectTransport(transport: StdioServerTransport | SSEServerTransport): Promise<void> {
    await ensureConfigLayout(this.configRoot);
    await this.server.connect(transport);
  }

  async close(): Promise<void> {
    await this.server.close();
  }

  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.connectTransport(transport);
    console.error(`[gmail-multi-inbox-mcp] Running on stdio. Config root: ${this.configRoot}`);
  }
}


function resolveTransportMode(): 'stdio' | 'sse' {
  const explicit = process.env.MCP_TRANSPORT?.trim().toLowerCase();
  if (explicit === 'stdio' || explicit === 'sse') return explicit;
  if (process.env.PORT?.trim()) return 'sse';
  return 'stdio';
}

async function runSseServer(): Promise<void> {
  const port = Number.parseInt(process.env.PORT ?? '3000', 10);
  const host = process.env.HOST?.trim() || '0.0.0.0';
  const sessions = new Map<string, { app: GmailMultiInboxServer; transport: SSEServerTransport }>();

  const closeAll = async (): Promise<void> => {
    const entries = [...sessions.values()];
    sessions.clear();
    await Promise.allSettled(entries.map(async ({ app, transport }) => {
      await Promise.allSettled([transport.close(), app.close()]);
    }));
  };

  const httpServer = createServer(async (req, res) => {
    try {
      const requestUrl = new URL(req.url ?? '/', `http://${req.headers.host ?? host}`);

      if (req.method === 'GET' && requestUrl.pathname === '/') {
        res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('gmail-multi-inbox-mcp SSE server');
        return;
      }

      if (req.method === 'GET' && requestUrl.pathname === '/sse') {
        const app = new GmailMultiInboxServer();
        const transport = new SSEServerTransport('/messages', res);
        const sessionId = transport.sessionId;
        sessions.set(sessionId, { app, transport });

        transport.onclose = () => {
          sessions.delete(sessionId);
        };

        try {
          await app.connectTransport(transport);
          console.error(`[gmail-multi-inbox-mcp] Running on SSE. Session: ${sessionId}. Port: ${port}`);
        } catch (error) {
          sessions.delete(sessionId);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
          }
          res.end(error instanceof Error ? error.message : String(error));
        }
        return;
      }

      if (req.method === 'POST' && requestUrl.pathname === '/messages') {
        const sessionId = requestUrl.searchParams.get('sessionId') ?? '';
        const session = sessions.get(sessionId);
        if (!session) {
          res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
          res.end('Unknown SSE session');
          return;
        }

        await session.transport.handlePostMessage(req, res);
        return;
      }

      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not found');
    } catch (error) {
      console.error('[gmail-multi-inbox-mcp] SSE server error:', error);
      if (!res.headersSent) {
        res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      }
      res.end(error instanceof Error ? error.message : String(error));
    }
  });

  await new Promise<void>((resolve) => {
    httpServer.listen(port, host, () => resolve());
  });

  console.error(`[gmail-multi-inbox-mcp] Running on SSE at http://${host}:${port}. Config dir uses GMAILMCPCONFIG_DIR or GMAIL_MCP_CONFIG_DIR.`);

  process.on('SIGINT', async () => {
    await closeAll();
    await new Promise<void>((resolve) => httpServer.close(() => resolve()));
    process.exit(0);
  });
}

async function main(): Promise<void> {
  if (resolveTransportMode() === 'sse') {
    await runSseServer();
    return;
  }

  const server = new GmailMultiInboxServer();
  process.on('SIGINT', async () => {
    await server.close();
    process.exit(0);
  });
  await server.run();
}

main().catch((error) => {
  console.error('[gmail-multi-inbox-mcp] Fatal error:', error);
  process.exit(1);
});
