#!/usr/bin/env node

import { promises as fs } from 'node:fs';
import express, { type Request, type Response } from 'express';
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
  type CalendarEvent,
  type CalendarInfo,
  type DriveFileDetail,
  type DriveFileSummary,
  type ParsedEmail,
  type SpreadsheetMetadata,
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

// ─── Drive arg interfaces ────────────────────────────────────────────────────

interface ListDriveFilesArgs {
  account?: string;
  folder_id?: string;
  query?: string;
  max_results?: number;
  page_token?: string;
}

interface DriveFileIdArgs {
  account: string;
  file_id: string;
}

interface UploadDriveFileArgs {
  account: string;
  local_path: string;
  name?: string;
  folder_id?: string;
  mime_type?: string;
}

interface CreateDriveFolderArgs {
  account: string;
  name: string;
  parent_id?: string;
}

interface UpdateDriveFileArgs {
  account: string;
  file_id: string;
  name?: string;
  add_parents?: string;
  remove_parents?: string;
  starred?: boolean;
  description?: string;
}

interface ShareDriveFileArgs {
  account: string;
  file_id: string;
  email?: string;
  role: string;
  type: string;
  send_notification?: boolean;
  notification_message?: string;
}

// ─── Sheets arg interfaces ───────────────────────────────────────────────────

interface SheetsSpreadsheetArgs {
  account: string;
  spreadsheet_id: string;
}

interface SheetsReadArgs extends SheetsSpreadsheetArgs {
  range: string;
  value_render_option?: string;
}

interface SheetsWriteArgs extends SheetsSpreadsheetArgs {
  range: string;
  values: unknown[][];
  value_input_option?: string;
}

interface SheetsAppendArgs extends SheetsSpreadsheetArgs {
  range: string;
  values: unknown[][];
}

interface SheetsCreateArgs {
  account: string;
  title: string;
}

interface SheetsTabArgs extends SheetsSpreadsheetArgs {
  sheet_title: string;
}

interface SheetsAddTabArgs extends SheetsSpreadsheetArgs {
  title: string;
  index?: number;
}

interface SheetsRenameTabArgs extends SheetsSpreadsheetArgs {
  current_title: string;
  new_title: string;
}

interface SheetsDimensionArgs extends SheetsSpreadsheetArgs {
  sheet_title: string;
  dimension: 'ROWS' | 'COLUMNS';
  start_index: number;
  count: number;
}

interface SheetsFormatArgs extends SheetsSpreadsheetArgs {
  sheet_title: string;
  range: string;
  bold?: boolean;
  italic?: boolean;
  font_size?: number;
  background_color?: string;
  text_color?: string;
  horizontal_alignment?: 'LEFT' | 'CENTER' | 'RIGHT';
  number_format?: string;
  wrap_strategy?: 'OVERFLOW_CELL' | 'LEGACY_WRAP' | 'CLIP' | 'WRAP';
}

interface SheetsAddChartArgs extends SheetsSpreadsheetArgs {
  sheet_title: string;
  chart_type: 'BAR' | 'LINE' | 'PIE' | 'COLUMN' | 'AREA' | 'SCATTER';
  data_range: string;
  title?: string;
  anchor_row?: number;
  anchor_col?: number;
  width_pixels?: number;
  height_pixels?: number;
}

// ─── Docs arg interfaces ─────────────────────────────────────────────────────

interface DocsDocumentArgs {
  account: string;
  document_id: string;
}

interface DocsCreateArgs {
  account: string;
  title: string;
}

interface DocsAppendArgs extends DocsDocumentArgs {
  text: string;
  style?: 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3';
  bold?: boolean;
  italic?: boolean;
}

interface DocsReplaceArgs extends DocsDocumentArgs {
  find: string;
  replace_with: string;
  match_case?: boolean;
}

interface DocsInsertTableArgs extends DocsDocumentArgs {
  rows: number;
  columns: number;
}

interface DocsApplyStyleArgs extends DocsDocumentArgs {
  start_index: number;
  end_index: number;
  style: 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3' | 'HEADING_4';
}

// ─── Calendar arg interfaces ─────────────────────────────────────────────────

interface CalendarAccountArgs {
  account?: string;
}

interface ListCalendarEventsArgs {
  account?: string;
  calendar_id?: string;
  time_min?: string;
  time_max?: string;
  query?: string;
  max_results?: number;
  single_events?: boolean;
}

interface CalendarEventArgs {
  account: string;
  calendar_id: string;
  event_id: string;
}

interface CreateCalendarEventArgs {
  account: string;
  calendar_id?: string;
  summary: string;
  description?: string;
  location?: string;
  start_date_time?: string;
  start_date?: string;
  end_date_time?: string;
  end_date?: string;
  time_zone?: string;
  attendees?: string[];
  recurrence?: string[];
  send_notifications?: boolean;
}

interface UpdateCalendarEventArgs {
  account: string;
  calendar_id: string;
  event_id: string;
  summary?: string;
  description?: string;
  location?: string;
  start_date_time?: string;
  start_date?: string;
  end_date_time?: string;
  end_date?: string;
  time_zone?: string;
  attendees?: string[];
  status?: string;
  send_notifications?: boolean;
}

interface DeleteCalendarEventArgs {
  account: string;
  calendar_id: string;
  event_id: string;
  send_notifications?: boolean;
}

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

function formatDriveFileDetail(file: DriveFileDetail): string {
  const lines = [
    `**Account**: ${file.accountId} (${file.accountEmail})`,
    `**File ID**: ${file.id}`,
    `**Name**: ${file.name}`,
    `**MIME Type**: ${file.mimeType}`,
    `**Size**: ${file.size !== undefined ? `${file.size.toLocaleString()} bytes` : '(unavailable)'}`,
    `**Created**: ${file.createdTime || '(unknown)'}`,
    `**Modified**: ${file.modifiedTime || '(unknown)'}`,
    `**Owners**: ${file.owners.length > 0 ? file.owners.join(', ') : '(unknown)'}`,
    `**Shared**: ${file.shared ? 'yes' : 'no'}`,
    `**Trashed**: ${file.trashed ? 'yes' : 'no'}`,
    `**Starred**: ${file.starred ? 'yes' : 'no'}`,
  ];
  if (file.parents?.length) lines.push(`**Parent IDs**: ${file.parents.join(', ')}`);
  if (file.webViewLink) lines.push(`**Open**: ${file.webViewLink}`);
  if (file.exportLinks && Object.keys(file.exportLinks).length > 0) {
    lines.push('**Export Links**:');
    for (const [mime, link] of Object.entries(file.exportLinks)) {
      lines.push(`  - ${mime}: ${link}`);
    }
  }
  return lines.join('\n');
}

function formatSpreadsheetMetadata(meta: SpreadsheetMetadata): string {
  const lines = [
    `**Spreadsheet ID**: ${meta.id}`,
    `**Title**: ${meta.title}`,
    `**URL**: ${meta.url}`,
    `**Sheets** (${meta.sheets.length}):`,
  ];
  for (const s of meta.sheets) {
    lines.push(`  - [${s.index}] "${s.title}" — ${s.rowCount} rows × ${s.columnCount} cols (sheetId: ${s.sheetId})`);
  }
  return lines.join('\n');
}

function formatCalendarInfo(cal: CalendarInfo): string {
  const primary = cal.primary ? ' (primary)' : '';
  const lines = [
    `**Calendar ID**: ${cal.id}`,
    `**Name**: ${cal.summary}${primary}`,
    `**Account**: ${cal.accountId} (${cal.accountEmail})`,
    `**Access Role**: ${cal.accessRole ?? '(unknown)'}`,
  ];
  if (cal.timeZone) lines.push(`**Time Zone**: ${cal.timeZone}`);
  if (cal.description) lines.push(`**Description**: ${cal.description}`);
  return lines.join('\n');
}

function formatCalendarEvent(event: CalendarEvent): string {
  const startTime = event.start.dateTime || event.start.date || '(unknown)';
  const endTime = event.end.dateTime || event.end.date || '(unknown)';
  const lines = [
    `**Account**: ${event.accountId} (${event.accountEmail})`,
    `**Calendar**: ${event.calendarId}`,
    `**Event ID**: ${event.id}`,
    `**Status**: ${event.status ?? 'confirmed'}`,
    `**Start**: ${startTime}`,
    `**End**: ${endTime}`,
  ];
  if (event.location) lines.push(`**Location**: ${event.location}`);
  if (event.description) {
    const desc = event.description.length > 400 ? `${event.description.slice(0, 400)}…` : event.description;
    lines.push(`**Description**: ${desc}`);
  }
  if (event.organizer) lines.push(`**Organizer**: ${event.organizer.displayName || event.organizer.email}`);
  if (event.attendees.length > 0) {
    lines.push(`**Attendees** (${event.attendees.length}):`);
    for (const a of event.attendees) {
      const status = a.responseStatus ? ` [${a.responseStatus}]` : '';
      const self = a.self ? ' (you)' : '';
      lines.push(`  - ${a.displayName || a.email}${status}${self}`);
    }
  }
  if (event.htmlLink) lines.push(`**Open**: ${event.htmlLink}`);
  if (event.recurrence?.length) lines.push(`**Recurrence**: ${event.recurrence.join(', ')}`);
  if (event.conferenceData?.entryPoints?.length) {
    const meet = event.conferenceData.entryPoints.find((ep) => ep.entryPointType === 'video');
    if (meet) lines.push(`**Meet Link**: ${meet.uri}`);
  }
  return lines.join('\n');
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
        name: 'ghub',
        version: '1.4.0',
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
      console.error('[ghub] MCP error:', error);
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
                description: 'Maximum emails to return (1-500).',
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
                description: 'Maximum emails to return (1-500).',
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
                description: 'Maximum files to return (1-500).',
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
              thread_id: {
                type: 'string',
                description: 'Optional Gmail thread ID. When set, the draft is created as a reply in that thread.',
              },
              in_reply_to: {
                type: 'string',
                description: 'Optional RFC 2822 Message-ID of the email being replied to. Sets the In-Reply-To header for proper threading.',
              },
              references: {
                type: 'string',
                description: 'Optional RFC 2822 References header value for threading.',
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

        // ─── Drive tools ────────────────────────────────────────────────────
        {
          name: 'list_drive_files',
          description: 'List Google Drive files. Omitting account aggregates across all enabled accounts. Optionally filter by folder or text query.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Optional account id. Omit to aggregate across all enabled accounts.' },
              folder_id: { type: 'string', description: 'Optional folder ID to list contents of a specific folder.' },
              query: { type: 'string', description: 'Optional text to search in file names and content.' },
              max_results: { type: 'number', description: 'Maximum files to return (1-500, default 25).' },
              page_token: { type: 'string', description: 'Pagination token from a previous response.' },
            },
            additionalProperties: false,
          },
        },
        {
          name: 'get_drive_file',
          description: 'Get full metadata for a specific Google Drive file including size, parents, export links.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              file_id: { type: 'string', description: 'Google Drive file ID.' },
            },
            required: ['account', 'file_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_drive_file_content',
          description: 'Download a Drive file and extract its text content. Google Docs/Sheets/Slides are exported to Office formats then parsed. PDFs, images, and text files are also supported.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              file_id: { type: 'string', description: 'Google Drive file ID.' },
            },
            required: ['account', 'file_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'upload_drive_file',
          description: 'Upload a local file to Google Drive.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              local_path: { type: 'string', description: 'Absolute local path to the file to upload.' },
              name: { type: 'string', description: 'Optional filename override in Drive.' },
              folder_id: { type: 'string', description: 'Optional parent folder ID.' },
              mime_type: { type: 'string', description: 'Optional MIME type override.' },
            },
            required: ['account', 'local_path'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_drive_folder',
          description: 'Create a new folder in Google Drive.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              name: { type: 'string', description: 'Folder name.' },
              parent_id: { type: 'string', description: 'Optional parent folder ID.' },
            },
            required: ['account', 'name'],
            additionalProperties: false,
          },
        },
        {
          name: 'update_drive_file',
          description: 'Update a Drive file — rename it, move it between folders, star/unstar, or update its description.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              file_id: { type: 'string', description: 'Google Drive file ID.' },
              name: { type: 'string', description: 'New file name.' },
              add_parents: { type: 'string', description: 'Comma-separated folder IDs to add as parents (moves file).' },
              remove_parents: { type: 'string', description: 'Comma-separated folder IDs to remove from parents.' },
              starred: { type: 'boolean', description: 'Star or unstar the file.' },
              description: { type: 'string', description: 'New file description.' },
            },
            required: ['account', 'file_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'trash_drive_file',
          description: 'Move a Google Drive file to trash.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              file_id: { type: 'string', description: 'Google Drive file ID.' },
            },
            required: ['account', 'file_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'share_drive_file',
          description: 'Share a Google Drive file by adding a permission.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              file_id: { type: 'string', description: 'Google Drive file ID.' },
              email: { type: 'string', description: 'Email address of the person to share with (for type=user or type=group).' },
              role: { type: 'string', enum: ['reader', 'commenter', 'writer', 'owner'], description: 'Permission role.' },
              type: { type: 'string', enum: ['user', 'group', 'domain', 'anyone'], description: 'Permission type.' },
              send_notification: { type: 'boolean', description: 'Send notification email (default true).', default: true },
              notification_message: { type: 'string', description: 'Optional message included in the share notification email.' },
            },
            required: ['account', 'file_id', 'role', 'type'],
            additionalProperties: false,
          },
        },

        // ─── Sheets tools ────────────────────────────────────────────────────
        {
          name: 'sheets_get',
          description: 'Get Google Sheets spreadsheet metadata: title, sheet tabs, row/column counts.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID (from the URL).' },
            },
            required: ['account', 'spreadsheet_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_read',
          description: 'Read cell values from a Google Sheet range (e.g., "Sheet1!A1:D10" or "A1:D10" for the first sheet).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              range: { type: 'string', description: 'A1 notation range, e.g. "Sheet1!A1:D10".' },
              value_render_option: { type: 'string', enum: ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'], description: 'How values are rendered (default FORMATTED_VALUE).' },
            },
            required: ['account', 'spreadsheet_id', 'range'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_write',
          description: 'Write values to a Google Sheet range. Overwrites existing values in the specified range.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              range: { type: 'string', description: 'A1 notation range, e.g. "Sheet1!A1".' },
              values: { type: 'array', items: { type: 'array' }, description: '2D array of values to write (rows × columns).' },
              value_input_option: { type: 'string', enum: ['RAW', 'USER_ENTERED'], description: 'How values are parsed (default USER_ENTERED).' },
            },
            required: ['account', 'spreadsheet_id', 'range', 'values'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_append',
          description: 'Append rows to a Google Sheet after the last row with data.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              range: { type: 'string', description: 'Sheet name or range to append to, e.g. "Sheet1".' },
              values: { type: 'array', items: { type: 'array' }, description: '2D array of rows to append.' },
            },
            required: ['account', 'spreadsheet_id', 'range', 'values'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_create',
          description: 'Create a new Google Sheets spreadsheet.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              title: { type: 'string', description: 'Spreadsheet title.' },
            },
            required: ['account', 'title'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_add_tab',
          description: 'Add a new sheet tab to an existing spreadsheet.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              title: { type: 'string', description: 'Name for the new sheet tab.' },
              index: { type: 'number', description: 'Optional zero-based position for the new tab.' },
            },
            required: ['account', 'spreadsheet_id', 'title'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_rename_tab',
          description: 'Rename a sheet tab in a spreadsheet.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              current_title: { type: 'string', description: 'Current name of the sheet tab.' },
              new_title: { type: 'string', description: 'New name for the sheet tab.' },
            },
            required: ['account', 'spreadsheet_id', 'current_title', 'new_title'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_delete_tab',
          description: 'Delete a sheet tab from a spreadsheet.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              sheet_title: { type: 'string', description: 'Name of the sheet tab to delete.' },
            },
            required: ['account', 'spreadsheet_id', 'sheet_title'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_format',
          description: 'Apply formatting to a range of cells: bold, italic, font size, background/text color, alignment, number format, wrap strategy.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              sheet_title: { type: 'string', description: 'Sheet tab name.' },
              range: { type: 'string', description: 'A1 notation range, e.g. "A1:D1".' },
              bold: { type: 'boolean', description: 'Apply bold.' },
              italic: { type: 'boolean', description: 'Apply italic.' },
              font_size: { type: 'number', description: 'Font size in points.' },
              background_color: { type: 'string', description: 'Background color as hex, e.g. "#FF0000".' },
              text_color: { type: 'string', description: 'Text color as hex, e.g. "#FFFFFF".' },
              horizontal_alignment: { type: 'string', enum: ['LEFT', 'CENTER', 'RIGHT'], description: 'Horizontal text alignment.' },
              number_format: { type: 'string', description: 'Number format pattern, e.g. "#,##0.00" or "MM/DD/YYYY".' },
              wrap_strategy: { type: 'string', enum: ['OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP'], description: 'Cell text wrap strategy.' },
            },
            required: ['account', 'spreadsheet_id', 'sheet_title', 'range'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_add_chart',
          description: 'Create a chart in a Google Sheet from a data range.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              sheet_title: { type: 'string', description: 'Sheet tab where the chart will be placed.' },
              chart_type: { type: 'string', enum: ['BAR', 'LINE', 'PIE', 'COLUMN', 'AREA', 'SCATTER'], description: 'Chart type.' },
              data_range: { type: 'string', description: 'A1 notation range for chart data, e.g. "A1:B10". First column = labels, second = values.' },
              title: { type: 'string', description: 'Optional chart title.' },
              anchor_row: { type: 'number', description: 'Row index (0-based) to anchor the chart (default 0).' },
              anchor_col: { type: 'number', description: 'Column index (0-based) to anchor the chart (default 0).' },
              width_pixels: { type: 'number', description: 'Chart width in pixels (default 600).' },
              height_pixels: { type: 'number', description: 'Chart height in pixels (default 400).' },
            },
            required: ['account', 'spreadsheet_id', 'sheet_title', 'chart_type', 'data_range'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_insert_dimension',
          description: 'Insert rows or columns into a sheet at a specified position.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              sheet_title: { type: 'string', description: 'Sheet tab name.' },
              dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'], description: 'Whether to insert rows or columns.' },
              start_index: { type: 'number', description: 'Zero-based index where rows/columns will be inserted.' },
              count: { type: 'number', description: 'Number of rows or columns to insert.' },
            },
            required: ['account', 'spreadsheet_id', 'sheet_title', 'dimension', 'start_index', 'count'],
            additionalProperties: false,
          },
        },
        {
          name: 'sheets_delete_dimension',
          description: 'Delete rows or columns from a sheet.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              spreadsheet_id: { type: 'string', description: 'Spreadsheet ID.' },
              sheet_title: { type: 'string', description: 'Sheet tab name.' },
              dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'], description: 'Whether to delete rows or columns.' },
              start_index: { type: 'number', description: 'Zero-based index of the first row/column to delete.' },
              count: { type: 'number', description: 'Number of rows or columns to delete.' },
            },
            required: ['account', 'spreadsheet_id', 'sheet_title', 'dimension', 'start_index', 'count'],
            additionalProperties: false,
          },
        },

        // ─── Docs tools ──────────────────────────────────────────────────────
        {
          name: 'docs_get',
          description: 'Get the full text content and title of a Google Doc.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              document_id: { type: 'string', description: 'Google Docs document ID.' },
            },
            required: ['account', 'document_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'docs_create',
          description: 'Create a new Google Doc.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              title: { type: 'string', description: 'Document title.' },
            },
            required: ['account', 'title'],
            additionalProperties: false,
          },
        },
        {
          name: 'docs_append',
          description: 'Append text to the end of a Google Doc. Optionally apply a heading style and bold/italic.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              document_id: { type: 'string', description: 'Google Docs document ID.' },
              text: { type: 'string', description: 'Text to append.' },
              style: { type: 'string', enum: ['NORMAL_TEXT', 'HEADING_1', 'HEADING_2', 'HEADING_3'], description: 'Paragraph style (default NORMAL_TEXT).' },
              bold: { type: 'boolean', description: 'Apply bold formatting.' },
              italic: { type: 'boolean', description: 'Apply italic formatting.' },
            },
            required: ['account', 'document_id', 'text'],
            additionalProperties: false,
          },
        },
        {
          name: 'docs_replace_text',
          description: 'Find and replace all occurrences of text throughout a Google Doc.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              document_id: { type: 'string', description: 'Google Docs document ID.' },
              find: { type: 'string', description: 'Text to find.' },
              replace_with: { type: 'string', description: 'Replacement text.' },
              match_case: { type: 'boolean', description: 'Case-sensitive match (default false).' },
            },
            required: ['account', 'document_id', 'find', 'replace_with'],
            additionalProperties: false,
          },
        },
        {
          name: 'docs_insert_table',
          description: 'Insert an empty table at the end of a Google Doc.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              document_id: { type: 'string', description: 'Google Docs document ID.' },
              rows: { type: 'number', description: 'Number of rows.' },
              columns: { type: 'number', description: 'Number of columns.' },
            },
            required: ['account', 'document_id', 'rows', 'columns'],
            additionalProperties: false,
          },
        },
        {
          name: 'docs_apply_style',
          description: 'Apply a heading or paragraph style to a text range in a Google Doc by character indices.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              document_id: { type: 'string', description: 'Google Docs document ID.' },
              start_index: { type: 'number', description: 'Start character index (from docs_get response).' },
              end_index: { type: 'number', description: 'End character index.' },
              style: { type: 'string', enum: ['NORMAL_TEXT', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4'], description: 'Style to apply.' },
            },
            required: ['account', 'document_id', 'start_index', 'end_index', 'style'],
            additionalProperties: false,
          },
        },

        // ─── Calendar tools ──────────────────────────────────────────────────
        {
          name: 'list_calendars',
          description: 'List all Google Calendars for one or all enabled accounts.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Optional account id. Omit to list calendars for all enabled accounts.' },
            },
            additionalProperties: false,
          },
        },
        {
          name: 'list_events',
          description: 'List calendar events. Omitting account aggregates across all enabled accounts. Supports time range and text search.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Optional account id.' },
              calendar_id: { type: 'string', description: 'Calendar ID (default "primary").' },
              time_min: { type: 'string', description: 'Lower bound for event start time (RFC3339, e.g. "2025-01-01T00:00:00Z").' },
              time_max: { type: 'string', description: 'Upper bound for event end time (RFC3339).' },
              query: { type: 'string', description: 'Free text search query.' },
              max_results: { type: 'number', description: 'Maximum events to return (1-250, default 25).' },
              single_events: { type: 'boolean', description: 'Expand recurring events into single instances (default true).', default: true },
            },
            additionalProperties: false,
          },
        },
        {
          name: 'get_event',
          description: 'Get full details of a single calendar event.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              calendar_id: { type: 'string', description: 'Calendar ID.' },
              event_id: { type: 'string', description: 'Event ID.' },
            },
            required: ['account', 'calendar_id', 'event_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_event',
          description: 'Create a new calendar event. Use start_date_time/end_date_time for timed events or start_date/end_date for all-day events.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              calendar_id: { type: 'string', description: 'Calendar ID (default "primary").' },
              summary: { type: 'string', description: 'Event title.' },
              description: { type: 'string', description: 'Event description.' },
              location: { type: 'string', description: 'Event location.' },
              start_date_time: { type: 'string', description: 'Start time as RFC3339, e.g. "2025-06-01T10:00:00".' },
              start_date: { type: 'string', description: 'All-day event start date as YYYY-MM-DD.' },
              end_date_time: { type: 'string', description: 'End time as RFC3339.' },
              end_date: { type: 'string', description: 'All-day event end date as YYYY-MM-DD.' },
              time_zone: { type: 'string', description: 'Time zone for start/end (e.g. "America/New_York").' },
              attendees: { type: 'array', items: { type: 'string' }, description: 'Email addresses of attendees.' },
              recurrence: { type: 'array', items: { type: 'string' }, description: 'RRULE strings, e.g. ["RRULE:FREQ=WEEKLY;COUNT=5"].' },
              send_notifications: { type: 'boolean', description: 'Send invite notifications to attendees (default true).' },
            },
            required: ['account', 'summary'],
            additionalProperties: false,
          },
        },
        {
          name: 'update_event',
          description: 'Update an existing calendar event. Only provided fields are changed.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              calendar_id: { type: 'string', description: 'Calendar ID.' },
              event_id: { type: 'string', description: 'Event ID.' },
              summary: { type: 'string', description: 'New event title.' },
              description: { type: 'string', description: 'New event description.' },
              location: { type: 'string', description: 'New event location.' },
              start_date_time: { type: 'string', description: 'New start time as RFC3339.' },
              start_date: { type: 'string', description: 'New all-day start date.' },
              end_date_time: { type: 'string', description: 'New end time as RFC3339.' },
              end_date: { type: 'string', description: 'New all-day end date.' },
              time_zone: { type: 'string', description: 'Time zone for updated start/end.' },
              attendees: { type: 'array', items: { type: 'string' }, description: 'Replace full attendees list.' },
              status: { type: 'string', enum: ['confirmed', 'tentative', 'cancelled'], description: 'Event status.' },
              send_notifications: { type: 'boolean', description: 'Notify attendees of changes (default true).' },
            },
            required: ['account', 'calendar_id', 'event_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_event',
          description: 'Delete a calendar event.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              calendar_id: { type: 'string', description: 'Calendar ID.' },
              event_id: { type: 'string', description: 'Event ID.' },
              send_notifications: { type: 'boolean', description: 'Notify attendees of cancellation (default true).' },
            },
            required: ['account', 'calendar_id', 'event_id'],
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

          // Drive
          case 'list_drive_files': return await this.handleListDriveFiles(args);
          case 'get_drive_file': return await this.handleGetDriveFile(args);
          case 'get_drive_file_content': return await this.handleGetDriveFileContent(args);
          case 'upload_drive_file': return await this.handleUploadDriveFile(args);
          case 'create_drive_folder': return await this.handleCreateDriveFolder(args);
          case 'update_drive_file': return await this.handleUpdateDriveFile(args);
          case 'trash_drive_file': return await this.handleTrashDriveFile(args);
          case 'share_drive_file': return await this.handleShareDriveFile(args);

          // Sheets
          case 'sheets_get': return await this.handleSheetsGet(args);
          case 'sheets_read': return await this.handleSheetsRead(args);
          case 'sheets_write': return await this.handleSheetsWrite(args);
          case 'sheets_append': return await this.handleSheetsAppend(args);
          case 'sheets_create': return await this.handleSheetsCreate(args);
          case 'sheets_add_tab': return await this.handleSheetsAddTab(args);
          case 'sheets_rename_tab': return await this.handleSheetsRenameTab(args);
          case 'sheets_delete_tab': return await this.handleSheetsDeleteTab(args);
          case 'sheets_format': return await this.handleSheetsFormat(args);
          case 'sheets_add_chart': return await this.handleSheetsAddChart(args);
          case 'sheets_insert_dimension': return await this.handleSheetsInsertDimension(args);
          case 'sheets_delete_dimension': return await this.handleSheetsDeleteDimension(args);

          // Docs
          case 'docs_get': return await this.handleDocsGet(args);
          case 'docs_create': return await this.handleDocsCreate(args);
          case 'docs_append': return await this.handleDocsAppend(args);
          case 'docs_replace_text': return await this.handleDocsReplaceText(args);
          case 'docs_insert_table': return await this.handleDocsInsertTable(args);
          case 'docs_apply_style': return await this.handleDocsApplyStyle(args);

          // Calendar
          case 'list_calendars': return await this.handleListCalendars(args);
          case 'list_events': return await this.handleListEvents(args);
          case 'get_event': return await this.handleGetEvent(args);
          case 'create_event': return await this.handleCreateEvent(args);
          case 'update_event': return await this.handleUpdateEvent(args);
          case 'delete_event': return await this.handleDeleteEvent(args);

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
    const maxResults = clamp(args.max_results ?? 20, 1, 500);

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
    const maxResults = clamp(args.max_results ?? 25, 1, 500);

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
    const maxResults = clamp(args.max_results ?? 25, 1, 500);

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
      threadId: valueToString(rawArgs.thread_id, '') || undefined,
      inReplyTo: valueToString(rawArgs.in_reply_to, '') || undefined,
      references: valueToString(rawArgs.references, '') || undefined,
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

  // ─── Drive handlers ────────────────────────────────────────────────────────

  private async handleListDriveFiles(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const accountId = valueToString(rawArgs.account, '').trim() || undefined;
    const folderId = valueToString(rawArgs.folder_id, '').trim() || undefined;
    const query = valueToString(rawArgs.query, '').trim() || undefined;
    const maxResults = clamp(valueToNumber(rawArgs.max_results, 25), 1, 500);
    const pageToken = valueToString(rawArgs.page_token, '').trim() || undefined;

    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, accountId);

    const results = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const result = await client.listDriveFiles({ folderId, query, maxResults, pageToken });
          return { account, files: result.files, nextPageToken: result.nextPageToken, error: null as string | null };
        } catch (error) {
          return { account, files: [] as DriveFileSummary[], nextPageToken: undefined, error: `${account.id}: ${(error as Error).message}` };
        }
      })
    );

    const merged = results.flatMap((r) => r.files).sort((a, b) => driveFileDateForSort(b) - driveFileDateForSort(a));
    const errors = results.map((r) => r.error).filter((e): e is string => Boolean(e));
    const returned = merged.slice(0, maxResults);
    const nextTokens = results.filter((r) => r.nextPageToken).map((r) => `${r.account.id}: ${r.nextPageToken}`);

    return textResult(
      formatDriveFileListOutput({
        title: 'Google Drive Files',
        scopeText: accountId ? `account ${accountId}` : `all enabled accounts (${targetAccounts.length})`,
        queryText: query ?? (folderId ? `folder: ${folderId}` : '(all files)'),
        totalFound: merged.length,
        returned,
        errors: nextTokens.length > 0 ? [...errors, `Next page tokens: ${nextTokens.join(', ')}`] : errors,
      })
    );
  }

  private async handleGetDriveFile(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const fileId = valueToString(rawArgs.file_id).trim();
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const file = await client.getDriveFile(fileId);
    return textResult(['# Google Drive File', '', formatDriveFileDetail(file)].join('\n'));
  }

  private async handleGetDriveFileContent(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const fileId = valueToString(rawArgs.file_id).trim();
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);

    const { bytes, contentType, filename } = await client.getDriveFileContent(fileId);
    const fakeMetadata: AttachmentMetadata = { id: fileId, filename, contentType, sizeBytes: bytes.length, isInline: false };
    const content = await saveAndExtract(bytes, fakeMetadata);

    const lines = [
      '# Google Drive File Content',
      '',
      `**Account**: ${acc.id} (${acc.email})`,
      `**File ID**: ${fileId}`,
      `**Filename**: ${filename}`,
      `**Content Type**: ${contentType}`,
      `**Size**: ${bytes.length.toLocaleString()} bytes`,
      `**Saved To**: ${content.savedPath}`,
      `**Extraction Method**: ${content.extractionMethod}`,
    ];
    if (content.textTruncated) lines.push('**Text Truncated**: true (file saved in full)');
    if (content.extractionError) lines.push(`**Extraction Error**: ${content.extractionError}`);
    if (content.text) { lines.push('', '**Content**:', '', content.text); }
    else { lines.push('', '*(no text extractable — file saved to disk)*'); }

    return textResult(lines.join('\n'));
  }

  private async handleUploadDriveFile(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const file = await client.uploadDriveFile({
      localPath: valueToString(rawArgs.local_path),
      name: valueToString(rawArgs.name, '').trim() || undefined,
      folderId: valueToString(rawArgs.folder_id, '').trim() || undefined,
      mimeType: valueToString(rawArgs.mime_type, '').trim() || undefined,
    });
    return textResult(['✅ File uploaded to Google Drive.', '', formatDriveFileDetail(file)].join('\n'));
  }

  private async handleCreateDriveFolder(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const folder = await client.createDriveFolder(
      valueToString(rawArgs.name),
      valueToString(rawArgs.parent_id, '').trim() || undefined
    );
    return textResult([
      '✅ Folder created.',
      '',
      formatDriveFileItem(folder),
    ].join('\n'));
  }

  private async handleUpdateDriveFile(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const file = await client.updateDriveFile(valueToString(rawArgs.file_id).trim(), {
      name: valueToString(rawArgs.name, '').trim() || undefined,
      addParents: valueToString(rawArgs.add_parents, '').trim() || undefined,
      removeParents: valueToString(rawArgs.remove_parents, '').trim() || undefined,
      starred: rawArgs.starred !== undefined ? valueToBoolean(rawArgs.starred, false) : undefined,
      description: valueToString(rawArgs.description, '').trim() || undefined,
    });
    return textResult(['✅ Drive file updated.', '', formatDriveFileItem(file)].join('\n'));
  }

  private async handleTrashDriveFile(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const fileId = valueToString(rawArgs.file_id).trim();
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.trashDriveFile(fileId);
    return textResult(`✅ File ${fileId} moved to trash in account ${acc.id}.`);
  }

  private async handleShareDriveFile(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.shareDriveFile(valueToString(rawArgs.file_id).trim(), {
      email: valueToString(rawArgs.email, '').trim() || undefined,
      role: valueToString(rawArgs.role),
      type: valueToString(rawArgs.type),
      sendNotification: rawArgs.send_notification !== undefined ? valueToBoolean(rawArgs.send_notification, true) : true,
      notificationMessage: valueToString(rawArgs.notification_message, '').trim() || undefined,
    });
    return textResult([
      `✅ Drive file shared in account ${acc.id}.`,
      `- File ID: ${valueToString(rawArgs.file_id)}`,
      `- Role: ${valueToString(rawArgs.role)}`,
      `- Type: ${valueToString(rawArgs.type)}`,
      rawArgs.email ? `- Email: ${valueToString(rawArgs.email)}` : '',
      `- Permission ID: ${result.permissionId}`,
    ].filter(Boolean).join('\n'));
  }

  // ─── Sheets handlers ───────────────────────────────────────────────────────

  private async handleSheetsGet(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const meta = await client.getSheetsMetadata(valueToString(rawArgs.spreadsheet_id));
    return textResult(['# Google Sheets Metadata', '', `**Account**: ${acc.id} (${acc.email})`, '', formatSpreadsheetMetadata(meta)].join('\n'));
  }

  private async handleSheetsRead(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.readSheetValues(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.range),
      valueToString(rawArgs.value_render_option, 'FORMATTED_VALUE') || 'FORMATTED_VALUE',
    );

    const lines: string[] = [
      '# Sheets Values',
      '',
      `**Account**: ${acc.id} (${acc.email})`,
      `**Spreadsheet**: ${valueToString(rawArgs.spreadsheet_id)}`,
      `**Range**: ${result.range}`,
      `**Rows**: ${result.values.length}`,
      '',
    ];

    if (result.values.length === 0) {
      lines.push('*(empty range — no data)*');
    } else {
      lines.push('```');
      for (const row of result.values) {
        lines.push((row as unknown[]).map((cell) => String(cell ?? '')).join('\t'));
      }
      lines.push('```');
    }

    return textResult(lines.join('\n'));
  }

  private async handleSheetsWrite(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);

    if (!Array.isArray(rawArgs.values)) throw new Error('values must be a 2D array.');
    const values = rawArgs.values as unknown[][];

    const result = await client.writeSheetValues(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.range),
      values,
      valueToString(rawArgs.value_input_option, 'USER_ENTERED') || 'USER_ENTERED',
    );

    return textResult([
      '✅ Sheets values written.',
      `Account: ${acc.id} (${acc.email})`,
      `Updated Range: ${result.updatedRange}`,
      `Rows Updated: ${result.updatedRows}`,
      `Cells Updated: ${result.updatedCells}`,
    ].join('\n'));
  }

  private async handleSheetsAppend(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);

    if (!Array.isArray(rawArgs.values)) throw new Error('values must be a 2D array.');
    const values = rawArgs.values as unknown[][];

    const result = await client.appendSheetValues(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.range),
      values,
    );

    return textResult([
      '✅ Rows appended.',
      `Account: ${acc.id} (${acc.email})`,
      `Updated Range: ${result.updatedRange}`,
      `Rows Appended: ${result.updatedRows}`,
    ].join('\n'));
  }

  private async handleSheetsCreate(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.createSpreadsheet(valueToString(rawArgs.title));
    return textResult([
      '✅ Spreadsheet created.',
      `Account: ${acc.id} (${acc.email})`,
      `Spreadsheet ID: ${result.id}`,
      `URL: ${result.url}`,
    ].join('\n'));
  }

  private async handleSheetsAddTab(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.addSheetTab(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.title),
      rawArgs.index !== undefined ? valueToNumber(rawArgs.index, 0) : undefined,
    );
    return textResult(`✅ Sheet tab "${result.title}" added (sheetId: ${result.sheetId}) in spreadsheet ${valueToString(rawArgs.spreadsheet_id)}.`);
  }

  private async handleSheetsRenameTab(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.renameSheetTab(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.current_title),
      valueToString(rawArgs.new_title),
    );
    return textResult(`✅ Sheet tab renamed from "${valueToString(rawArgs.current_title)}" to "${valueToString(rawArgs.new_title)}".`);
  }

  private async handleSheetsDeleteTab(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.deleteSheetTab(valueToString(rawArgs.spreadsheet_id), valueToString(rawArgs.sheet_title));
    return textResult(`✅ Sheet tab "${valueToString(rawArgs.sheet_title)}" deleted.`);
  }

  private async handleSheetsFormat(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.formatCells(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.sheet_title),
      valueToString(rawArgs.range),
      {
        bold: rawArgs.bold !== undefined ? valueToBoolean(rawArgs.bold) : undefined,
        italic: rawArgs.italic !== undefined ? valueToBoolean(rawArgs.italic) : undefined,
        fontSize: rawArgs.font_size !== undefined ? valueToNumber(rawArgs.font_size, 0) || undefined : undefined,
        backgroundColor: valueToString(rawArgs.background_color, '').trim() || undefined,
        textColor: valueToString(rawArgs.text_color, '').trim() || undefined,
        horizontalAlignment: (valueToString(rawArgs.horizontal_alignment, '').trim() || undefined) as 'LEFT' | 'CENTER' | 'RIGHT' | undefined,
        numberFormat: valueToString(rawArgs.number_format, '').trim() || undefined,
        wrapStrategy: (valueToString(rawArgs.wrap_strategy, '').trim() || undefined) as 'OVERFLOW_CELL' | 'LEGACY_WRAP' | 'CLIP' | 'WRAP' | undefined,
      },
    );
    return textResult(`✅ Formatting applied to ${valueToString(rawArgs.range)} in sheet "${valueToString(rawArgs.sheet_title)}".`);
  }

  private async handleSheetsAddChart(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.addChart(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.sheet_title),
      {
        chartType: valueToString(rawArgs.chart_type) as 'BAR' | 'LINE' | 'PIE' | 'COLUMN' | 'AREA' | 'SCATTER',
        dataRange: valueToString(rawArgs.data_range),
        title: valueToString(rawArgs.title, '').trim() || undefined,
        anchorRow: rawArgs.anchor_row !== undefined ? valueToNumber(rawArgs.anchor_row, 0) : undefined,
        anchorCol: rawArgs.anchor_col !== undefined ? valueToNumber(rawArgs.anchor_col, 0) : undefined,
        widthPixels: rawArgs.width_pixels !== undefined ? valueToNumber(rawArgs.width_pixels, 600) : undefined,
        heightPixels: rawArgs.height_pixels !== undefined ? valueToNumber(rawArgs.height_pixels, 400) : undefined,
      },
    );
    return textResult([
      '✅ Chart created.',
      `Account: ${acc.id} (${acc.email})`,
      `Chart ID: ${result.chartId}`,
      `Type: ${valueToString(rawArgs.chart_type)}`,
      `Data Range: ${valueToString(rawArgs.data_range)}`,
    ].join('\n'));
  }

  private async handleSheetsInsertDimension(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const dim = valueToString(rawArgs.dimension) as 'ROWS' | 'COLUMNS';
    const count = valueToNumber(rawArgs.count, 1);
    await client.insertDimension(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.sheet_title),
      dim,
      valueToNumber(rawArgs.start_index, 0),
      count,
    );
    return textResult(`✅ Inserted ${count} ${dim.toLowerCase()} at index ${valueToNumber(rawArgs.start_index, 0)} in sheet "${valueToString(rawArgs.sheet_title)}".`);
  }

  private async handleSheetsDeleteDimension(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const dim = valueToString(rawArgs.dimension) as 'ROWS' | 'COLUMNS';
    const count = valueToNumber(rawArgs.count, 1);
    await client.deleteDimension(
      valueToString(rawArgs.spreadsheet_id),
      valueToString(rawArgs.sheet_title),
      dim,
      valueToNumber(rawArgs.start_index, 0),
      count,
    );
    return textResult(`✅ Deleted ${count} ${dim.toLowerCase()} starting at index ${valueToNumber(rawArgs.start_index, 0)} in sheet "${valueToString(rawArgs.sheet_title)}".`);
  }

  // ─── Docs handlers ─────────────────────────────────────────────────────────

  private async handleDocsGet(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const doc = await client.getDocument(valueToString(rawArgs.document_id));
    const preview = doc.body.length > 2000 ? `${doc.body.slice(0, 2000)}\n\n… (${doc.body.length - 2000} more characters)` : doc.body;
    return textResult([
      '# Google Doc',
      '',
      `**Account**: ${acc.id} (${acc.email})`,
      `**Document ID**: ${doc.documentId}`,
      `**Title**: ${doc.title}`,
      `**URL**: ${doc.url}`,
      `**Length**: ${doc.body.length} characters`,
      '',
      '**Content**:',
      '',
      preview,
    ].join('\n'));
  }

  private async handleDocsCreate(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.createDocument(valueToString(rawArgs.title));
    return textResult([
      '✅ Google Doc created.',
      `Account: ${acc.id} (${acc.email})`,
      `Document ID: ${result.documentId}`,
      `URL: ${result.url}`,
    ].join('\n'));
  }

  private async handleDocsAppend(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.appendToDocument(
      valueToString(rawArgs.document_id),
      valueToString(rawArgs.text),
      {
        style: (valueToString(rawArgs.style, 'NORMAL_TEXT') || 'NORMAL_TEXT') as 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3',
        bold: rawArgs.bold !== undefined ? valueToBoolean(rawArgs.bold) : undefined,
        italic: rawArgs.italic !== undefined ? valueToBoolean(rawArgs.italic) : undefined,
      },
    );
    return textResult(`✅ Text appended to document ${valueToString(rawArgs.document_id)} in account ${acc.id}.`);
  }

  private async handleDocsReplaceText(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const result = await client.replaceInDocument(
      valueToString(rawArgs.document_id),
      valueToString(rawArgs.find),
      valueToString(rawArgs.replace_with),
      valueToBoolean(rawArgs.match_case, false),
    );
    return textResult(`✅ Replaced ${result.occurrencesChanged} occurrence(s) of "${valueToString(rawArgs.find)}" in document ${valueToString(rawArgs.document_id)}.`);
  }

  private async handleDocsInsertTable(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const rows = valueToNumber(rawArgs.rows, 2);
    const cols = valueToNumber(rawArgs.columns, 2);
    await client.insertTableInDocument(valueToString(rawArgs.document_id), rows, cols);
    return textResult(`✅ Inserted ${rows}×${cols} table at end of document ${valueToString(rawArgs.document_id)}.`);
  }

  private async handleDocsApplyStyle(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.applyDocHeadingStyle(
      valueToString(rawArgs.document_id),
      valueToNumber(rawArgs.start_index, 0),
      valueToNumber(rawArgs.end_index, 0),
      valueToString(rawArgs.style) as 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3' | 'HEADING_4',
    );
    return textResult(`✅ Applied style "${valueToString(rawArgs.style)}" to indices ${valueToNumber(rawArgs.start_index, 0)}-${valueToNumber(rawArgs.end_index, 0)} in document ${valueToString(rawArgs.document_id)}.`);
  }

  // ─── Calendar handlers ─────────────────────────────────────────────────────

  private async handleListCalendars(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const accountId = valueToString(rawArgs.account, '').trim() || undefined;
    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, accountId);

    const results = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const cals = await client.listCalendars();
          return { cals, error: null as string | null };
        } catch (error) {
          return { cals: [] as CalendarInfo[], error: `${account.id}: ${(error as Error).message}` };
        }
      })
    );

    const all = results.flatMap((r) => r.cals);
    const errors = results.map((r) => r.error).filter((e): e is string => Boolean(e));

    const lines: string[] = [
      '# Google Calendars',
      '',
      `**Scope**: ${accountId ? `account ${accountId}` : `all enabled accounts (${targetAccounts.length})`}`,
      `**Total**: ${all.length}`,
      '',
    ];

    all.forEach((cal, i) => {
      lines.push(`## ${i + 1}. ${cal.summary}${cal.primary ? ' (primary)' : ''}`);
      lines.push(formatCalendarInfo(cal));
      lines.push('');
    });

    if (errors.length > 0) { lines.push('## Account Errors'); lines.push(errors.map((e) => `- ${e}`).join('\n')); }
    return textResult(lines.join('\n'));
  }

  private async handleListEvents(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const accountId = valueToString(rawArgs.account, '').trim() || undefined;
    const calendarId = valueToString(rawArgs.calendar_id, 'primary').trim() || 'primary';
    const timeMin = valueToString(rawArgs.time_min, '').trim() || undefined;
    const timeMax = valueToString(rawArgs.time_max, '').trim() || undefined;
    const query = valueToString(rawArgs.query, '').trim() || undefined;
    const maxResults = clamp(valueToNumber(rawArgs.max_results, 25), 1, 250);
    const singleEvents = valueToBoolean(rawArgs.single_events, true);

    const config = await this.loadConfig();
    const targetAccounts = resolveReadAccounts(config, accountId);

    const results = await Promise.all(
      targetAccounts.map(async (account) => {
        try {
          const client = await this.getClientForAccount(account);
          const events = await client.listCalendarEvents(calendarId, { timeMin, timeMax, query, maxResults, singleEvents });
          return { events, error: null as string | null };
        } catch (error) {
          return { events: [] as CalendarEvent[], error: `${account.id}: ${(error as Error).message}` };
        }
      })
    );

    const merged = results
      .flatMap((r) => r.events)
      .sort((a, b) => {
        const at = a.start.dateTime || a.start.date || '';
        const bt = b.start.dateTime || b.start.date || '';
        return at < bt ? -1 : at > bt ? 1 : 0;
      });

    const errors = results.map((r) => r.error).filter((e): e is string => Boolean(e));
    const returned = merged.slice(0, maxResults);

    const lines: string[] = [
      '# Calendar Events',
      '',
      `**Scope**: ${accountId ? `account ${accountId}` : `all enabled accounts (${targetAccounts.length})`}`,
      `**Calendar**: ${calendarId}`,
    ];
    if (timeMin) lines.push(`**From**: ${timeMin}`);
    if (timeMax) lines.push(`**To**: ${timeMax}`);
    if (query) lines.push(`**Query**: ${query}`);
    lines.push(`**Total Found**: ${merged.length}`, `**Returned**: ${returned.length}`, '');

    returned.forEach((event, i) => {
      const startTime = event.start.dateTime || event.start.date || '(unknown)';
      lines.push(`## ${i + 1}. ${event.summary}`);
      lines.push(`**Start**: ${startTime}`);
      lines.push(formatCalendarEvent(event));
      lines.push('');
    });

    if (errors.length > 0) { lines.push('## Account Errors'); lines.push(errors.map((e) => `- ${e}`).join('\n')); }
    return textResult(lines.join('\n'));
  }

  private async handleGetEvent(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    const event = await client.getCalendarEvent(valueToString(rawArgs.calendar_id), valueToString(rawArgs.event_id));
    return textResult(['# Calendar Event', '', `**Title**: ${event.summary}`, '', formatCalendarEvent(event)].join('\n'));
  }

  private async handleCreateEvent(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);

    const tz = valueToString(rawArgs.time_zone, '').trim() || undefined;
    const startDT = valueToString(rawArgs.start_date_time, '').trim() || undefined;
    const startD = valueToString(rawArgs.start_date, '').trim() || undefined;
    const endDT = valueToString(rawArgs.end_date_time, '').trim() || undefined;
    const endD = valueToString(rawArgs.end_date, '').trim() || undefined;

    if (!startDT && !startD) throw new Error('Provide start_date_time or start_date.');
    if (!endDT && !endD) throw new Error('Provide end_date_time or end_date.');

    const event = await client.createCalendarEvent(
      valueToString(rawArgs.calendar_id, 'primary').trim() || 'primary',
      {
        summary: valueToString(rawArgs.summary),
        description: valueToString(rawArgs.description, '').trim() || undefined,
        location: valueToString(rawArgs.location, '').trim() || undefined,
        start: startDT ? { dateTime: startDT, timeZone: tz } : { date: startD },
        end: endDT ? { dateTime: endDT, timeZone: tz } : { date: endD },
        attendees: valueToStringArray(rawArgs.attendees),
        recurrence: valueToStringArray(rawArgs.recurrence),
        sendNotifications: rawArgs.send_notifications !== undefined ? valueToBoolean(rawArgs.send_notifications, true) : true,
      },
    );

    return textResult(['✅ Calendar event created.', '', formatCalendarEvent(event)].join('\n'));
  }

  private async handleUpdateEvent(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);

    const tz = valueToString(rawArgs.time_zone, '').trim() || undefined;
    const startDT = valueToString(rawArgs.start_date_time, '').trim() || undefined;
    const startD = valueToString(rawArgs.start_date, '').trim() || undefined;
    const endDT = valueToString(rawArgs.end_date_time, '').trim() || undefined;
    const endD = valueToString(rawArgs.end_date, '').trim() || undefined;

    const startUpdate = startDT ? { dateTime: startDT, timeZone: tz } : startD ? { date: startD } : undefined;
    const endUpdate = endDT ? { dateTime: endDT, timeZone: tz } : endD ? { date: endD } : undefined;

    const event = await client.updateCalendarEvent(
      valueToString(rawArgs.calendar_id),
      valueToString(rawArgs.event_id),
      {
        summary: valueToString(rawArgs.summary, '').trim() || undefined,
        description: valueToString(rawArgs.description, '').trim() || undefined,
        location: valueToString(rawArgs.location, '').trim() || undefined,
        start: startUpdate,
        end: endUpdate,
        attendees: rawArgs.attendees !== undefined ? valueToStringArray(rawArgs.attendees) : undefined,
        status: (valueToString(rawArgs.status, '').trim() || undefined) as string | undefined,
        sendNotifications: rawArgs.send_notifications !== undefined ? valueToBoolean(rawArgs.send_notifications, true) : true,
      },
    );

    return textResult(['✅ Calendar event updated.', '', formatCalendarEvent(event)].join('\n'));
  }

  private async handleDeleteEvent(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account = valueToString(rawArgs.account);
    const config = await this.loadConfig();
    const acc = resolveWriteAccount(config, account);
    const client = await this.getClientForAccount(acc);
    await client.deleteCalendarEvent(
      valueToString(rawArgs.calendar_id),
      valueToString(rawArgs.event_id),
      rawArgs.send_notifications !== undefined ? valueToBoolean(rawArgs.send_notifications, true) : true,
    );
    return textResult(`✅ Event ${valueToString(rawArgs.event_id)} deleted from calendar ${valueToString(rawArgs.calendar_id)} in account ${acc.id}.`);
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
    console.error(`[ghub] Running on stdio. Config root: ${this.configRoot}`);
  }
}


function resolveTransportMode(): 'stdio' | 'sse' {
  const explicit = process.env.MCP_TRANSPORT?.trim().toLowerCase();
  if (explicit === 'stdio' || explicit === 'sse') return explicit;
  if (process.env.PORT?.trim()) return 'sse';
  return 'stdio';
}

async function runSseServer(): Promise<void> {
  const port = Number.parseInt(process.env.PORT ?? process.env.RAILWAY_PORT ?? '3000', 10);
  if (!Number.isFinite(port) || port <= 0) {
    throw new Error('Invalid PORT value: ' + (process.env.PORT ?? '(missing)'));
  }

  const host = process.env.HOST?.trim() || '0.0.0.0';
  const app = express();
  const sessions = new Map<string, { app: GmailMultiInboxServer; transport: SSEServerTransport }>();

  const closeAll = async (): Promise<void> => {
    const entries = [...sessions.values()];
    sessions.clear();
    await Promise.allSettled(entries.map(async ({ app, transport }) => {
      await Promise.allSettled([transport.close(), app.close()]);
    }));
  };

  app.get('/', (_req: Request, res: Response) => {
    res.status(200).type('text/plain').send('ghub SSE server');
  });

  app.use((req: Request, _res: Response, next) => {
    console.error('[ghub] HTTP ' + req.method + ' ' + req.path);
    next();
  });

  app.get('/sse', async (_req: Request, res: Response) => {
    const serverApp = new GmailMultiInboxServer();
    const transport = new SSEServerTransport('/messages', res);
    const sessionId = transport.sessionId;
    sessions.set(sessionId, { app: serverApp, transport });

    transport.onclose = () => {
      sessions.delete(sessionId);
    };

    try {
      await serverApp.connectTransport(transport);
      console.error('[ghub] SSE session started: ' + sessionId);
    } catch (error) {
      sessions.delete(sessionId);
      if (!res.headersSent) {
        res.status(500).type('text/plain');
      }
      res.end(error instanceof Error ? error.message : String(error));
    }
  });

  app.post('/messages', async (req: Request, res: Response) => {
    const sessionId = typeof req.query.sessionId === 'string' ? req.query.sessionId : '';
    const session = sessions.get(sessionId);
    if (!session) {
      res.status(404).type('text/plain').send('Unknown SSE session');
      return;
    }

    await session.transport.handlePostMessage(req, res);
  });

  const httpServer = app.listen(port, () => {
    console.error('[ghub] Running on SSE at port ' + port + '. Routes: GET /sse, POST /messages');
  });

  const shutdown = async (): Promise<void> => {
    await closeAll();
    await new Promise<void>((resolve) => httpServer.close(() => resolve()));
  };

  process.on('SIGINT', async () => {
    await shutdown();
    process.exit(0);
  });
  process.on('SIGTERM', async () => {
    await shutdown();
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
  process.on('SIGTERM', async () => {
    await server.close();
    process.exit(0);
  });
  await server.run();
}

main().catch((error) => {
  console.error('[ghub] Fatal error:', error);
  process.exit(1);
});
