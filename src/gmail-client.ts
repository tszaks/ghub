import { OAuth2Client, type Credentials } from 'google-auth-library';
import { google, type calendar_v3, type docs_v1, type drive_v3, type gmail_v1, type sheets_v4 } from 'googleapis';
import { promises as fs, createReadStream } from 'node:fs';
import path from 'node:path';
import { AccountConfig, type AccountPaths, getAccountPaths } from './config.js';

export const GMAIL_SCOPES = [
  'https://mail.google.com/',
] as const;

export const DRIVE_METADATA_SCOPE =
  'https://www.googleapis.com/auth/drive.metadata.readonly' as const;

export const DRIVE_FULL_SCOPE =
  'https://www.googleapis.com/auth/drive' as const;

export const SHEETS_SCOPE =
  'https://www.googleapis.com/auth/spreadsheets' as const;

export const DOCS_SCOPE =
  'https://www.googleapis.com/auth/documents' as const;

export const CALENDAR_SCOPE =
  'https://www.googleapis.com/auth/calendar' as const;

export const GOOGLE_ACCOUNT_SCOPES = [
  ...GMAIL_SCOPES,
  DRIVE_FULL_SCOPE,
  SHEETS_SCOPE,
  DOCS_SCOPE,
  CALENDAR_SCOPE,
] as const;

export interface AttachmentMetadata {
  id: string;
  filename: string;
  contentType: string;
  sizeBytes: number;
  isInline: boolean;
}

export interface ParsedEmail {
  id: string;
  threadId: string;
  snippet: string;
  from: string;
  to: string;
  cc: string;
  subject: string;
  date: string;
  internalDate: number;
  messageHeaderId: string;
  inReplyTo: string;
  references: string;
  body?: string;
  labels: string[];
  attachments: AttachmentMetadata[];
  accountId: string;
  accountEmail: string;
}

export interface ParsedThread {
  threadId: string;
  messages: ParsedEmail[];
}

export interface LabelInfo {
  id: string;
  name: string;
  type?: string;
  messagesTotal?: number;
}

export interface DriveFileSummary {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime?: string;
  owners: string[];
  webViewLink?: string;
  accountId: string;
  accountEmail: string;
}

export interface DriveFileDetail extends DriveFileSummary {
  createdTime?: string;
  size?: number;
  parents?: string[];
  trashed?: boolean;
  starred?: boolean;
  shared?: boolean;
  exportLinks?: Record<string, string>;
}

export interface SpreadsheetSheetInfo {
  sheetId: number;
  title: string;
  index: number;
  rowCount: number;
  columnCount: number;
}

export interface SpreadsheetMetadata {
  id: string;
  title: string;
  url: string;
  sheets: SpreadsheetSheetInfo[];
}

export interface CalendarInfo {
  id: string;
  summary: string;
  description?: string;
  primary?: boolean;
  accessRole?: string;
  backgroundColor?: string;
  timeZone?: string;
  accountId: string;
  accountEmail: string;
}

export interface CalendarEventAttendee {
  email: string;
  displayName?: string;
  responseStatus?: string;
  self?: boolean;
}

export interface CalendarEvent {
  id: string;
  calendarId: string;
  summary: string;
  description?: string;
  location?: string;
  start: { dateTime?: string; date?: string; timeZone?: string };
  end: { dateTime?: string; date?: string; timeZone?: string };
  attendees: CalendarEventAttendee[];
  organizer?: { email: string; displayName?: string; self?: boolean };
  status?: string;
  htmlLink?: string;
  recurrence?: string[];
  recurringEventId?: string;
  created?: string;
  updated?: string;
  conferenceData?: { entryPoints?: Array<{ uri: string; entryPointType: string }> };
  accountId: string;
  accountEmail: string;
}

interface OAuthClientOptions {
  credentials: unknown;
}

export interface EmailAttachment {
  path: string;
  filename?: string;
  contentType?: string;
}

function decodeBase64Url(value: string): string {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
  return Buffer.from(padded, 'base64').toString('utf8');
}

function decodeBase64UrlBuffer(value: string): Buffer {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
  return Buffer.from(padded, 'base64');
}

function isInlineDisposition(
  headers: gmail_v1.Schema$MessagePartHeader[] | undefined,
): boolean {
  if (!headers) return false;
  const disposition = headers.find((h) => h.name?.toLowerCase() === 'content-disposition');
  return Boolean(disposition?.value?.trim().toLowerCase().startsWith('inline'));
}

function extractAttachmentsMetadata(
  payload: gmail_v1.Schema$MessagePart | undefined,
): AttachmentMetadata[] {
  if (!payload) return [];

  const out: AttachmentMetadata[] = [];
  const consider = (part: gmail_v1.Schema$MessagePart) => {
    if (part.filename && part.body?.attachmentId) {
      out.push({
        id: part.body.attachmentId,
        filename: part.filename,
        contentType: part.mimeType ?? 'application/octet-stream',
        sizeBytes: Number(part.body.size ?? 0),
        isInline: isInlineDisposition(part.headers),
      });
    }
  };

  consider(payload);

  const stack: gmail_v1.Schema$MessagePart[] = payload.parts ? [...payload.parts] : [];
  while (stack.length > 0) {
    const part = stack.shift();
    if (!part) continue;
    consider(part);
    if (part.parts?.length) stack.push(...part.parts);
  }

  return out;
}

function findAttachmentPart(
  payload: gmail_v1.Schema$MessagePart | undefined,
  attachmentId: string,
): gmail_v1.Schema$MessagePart | null {
  if (!payload) return null;
  if (payload.body?.attachmentId === attachmentId) return payload;

  const stack: gmail_v1.Schema$MessagePart[] = payload.parts ? [...payload.parts] : [];
  while (stack.length > 0) {
    const part = stack.shift();
    if (!part) continue;
    if (part.body?.attachmentId === attachmentId) return part;
    if (part.parts?.length) stack.push(...part.parts);
  }

  return null;
}

function stripHtmlTags(input: string): string {
  return input.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function getHeaderValue(
  headers: gmail_v1.Schema$MessagePartHeader[] | undefined,
  name: string
): string {
  if (!headers) return '';
  const found = headers.find((header) => header.name?.toLowerCase() === name.toLowerCase());
  return found?.value?.trim() ?? '';
}

function extractEmailBody(payload?: gmail_v1.Schema$MessagePart): string {
  if (!payload) return '';

  if (payload.body?.data && !payload.parts?.length) {
    return decodeBase64Url(payload.body.data);
  }

  if (!payload.parts || payload.parts.length === 0) {
    return '';
  }

  let textPlain = '';
  let textHtml = '';

  const stack = [...payload.parts];
  while (stack.length > 0) {
    const part = stack.shift();
    if (!part) continue;

    if (part.parts?.length) {
      stack.push(...part.parts);
    }

    if (!part.body?.data) continue;

    if (part.mimeType === 'text/plain' && !textPlain) {
      textPlain = decodeBase64Url(part.body.data);
    } else if (part.mimeType === 'text/html' && !textHtml) {
      textHtml = decodeBase64Url(part.body.data);
    }
  }

  if (textPlain) return textPlain;
  if (textHtml) return stripHtmlTags(textHtml);

  return '';
}

function normalizeOutgoingAddressList(value?: string): string | null {
  if (!value || value.trim() === '') return null;
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
    .join(', ');
}

function encodeBase64Url(value: string): string {
  return Buffer.from(value)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function wrapBase64(value: string): string {
  return value.replace(/.{1,76}/g, '$&\r\n').trimEnd();
}

function inferContentType(filename: string): string {
  const extension = path.extname(filename).toLowerCase();
  switch (extension) {
    case '.pdf':
      return 'application/pdf';
    case '.txt':
      return 'text/plain';
    case '.csv':
      return 'text/csv';
    case '.json':
      return 'application/json';
    case '.png':
      return 'image/png';
    case '.jpg':
    case '.jpeg':
      return 'image/jpeg';
    case '.gif':
      return 'image/gif';
    case '.webp':
      return 'image/webp';
    case '.doc':
      return 'application/msword';
    case '.docx':
      return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    default:
      return 'application/octet-stream';
  }
}

function sanitizeHeaderValue(value: string): string {
  return value.replace(/[\r\n"]/g, ' ').trim();
}

async function buildRawEmailMessage(input: {
  to: string;
  subject: string;
  body: string;
  cc?: string;
  bcc?: string;
  html?: boolean;
  attachments?: EmailAttachment[];
  inReplyTo?: string;
  references?: string;
}): Promise<string> {
  const to = normalizeOutgoingAddressList(input.to);
  if (!to) {
    throw new Error('Recipient "to" is required.');
  }

  const attachments = (input.attachments ?? []).filter((attachment) => attachment.path.trim() !== '');

  if (attachments.length === 0) {
    const lines: string[] = [
      `To: ${to}`,
      `Subject: ${input.subject}`,
      'MIME-Version: 1.0',
      `Content-Type: text/${input.html ? 'html' : 'plain'}; charset=utf-8`,
    ];

    const cc = normalizeOutgoingAddressList(input.cc);
    if (cc) lines.push(`Cc: ${cc}`);

    const bcc = normalizeOutgoingAddressList(input.bcc);
    if (bcc) lines.push(`Bcc: ${bcc}`);

    if (input.inReplyTo) lines.push(`In-Reply-To: ${input.inReplyTo}`);
    if (input.references) lines.push(`References: ${input.references}`);

    lines.push('', input.body);
    return encodeBase64Url(lines.join('\r\n'));
  }

  const lines: string[] = [
    `To: ${to}`,
    `Subject: ${input.subject}`,
    'MIME-Version: 1.0',
  ];

  const cc = normalizeOutgoingAddressList(input.cc);
  if (cc) lines.push(`Cc: ${cc}`);

  const bcc = normalizeOutgoingAddressList(input.bcc);
  if (bcc) lines.push(`Bcc: ${bcc}`);

  const boundary = `gmail-multi-inbox-mcp-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 10)}`;
  lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`, '');

  lines.push(
    `--${boundary}`,
    `Content-Type: text/${input.html ? 'html' : 'plain'}; charset=utf-8`,
    'Content-Transfer-Encoding: base64',
    '',
    wrapBase64(Buffer.from(input.body, 'utf8').toString('base64'))
  );

  for (const attachment of attachments) {
    const filePath = attachment.path.trim();
    const fileBuffer = await fs.readFile(filePath);
    const filename = sanitizeHeaderValue(
      attachment.filename?.trim() || path.basename(filePath)
    );
    const contentType = attachment.contentType?.trim() || inferContentType(filename);

    lines.push(
      `--${boundary}`,
      `Content-Type: ${contentType}; name="${filename}"`,
      'Content-Transfer-Encoding: base64',
      `Content-Disposition: attachment; filename="${filename}"`,
      '',
      wrapBase64(fileBuffer.toString('base64'))
    );
  }

  lines.push(`--${boundary}--`);

  return encodeBase64Url(lines.join('\r\n'));
}

function normalizeAttachments(attachments?: EmailAttachment[]): EmailAttachment[] {
  return (attachments ?? [])
    .map((attachment) => ({
      path: attachment.path.trim(),
      filename: attachment.filename?.trim() || undefined,
      contentType: attachment.contentType?.trim() || undefined,
    }))
    .filter((attachment) => attachment.path !== '');
}

async function createRawEmailMessage(input: {
  to: string;
  subject: string;
  body: string;
  cc?: string;
  bcc?: string;
  html?: boolean;
  attachments?: EmailAttachment[];
  inReplyTo?: string;
  references?: string;
}): Promise<string> {
  return buildRawEmailMessage({
    ...input,
    attachments: normalizeAttachments(input.attachments),
  });
}

export function createOAuthClientFromCredentials(options: OAuthClientOptions): OAuth2Client {
  if (!options.credentials || typeof options.credentials !== 'object') {
    throw new Error('Invalid credentials content.');
  }

  const credentialsObject = options.credentials as {
    installed?: {
      client_id?: string;
      client_secret?: string;
      redirect_uris?: string[];
    };
    web?: {
      client_id?: string;
      client_secret?: string;
      redirect_uris?: string[];
    };
  };

  const source = credentialsObject.installed ?? credentialsObject.web;
  if (!source?.client_id || !source.client_secret) {
    throw new Error('Credentials must include client_id and client_secret under "installed" or "web".');
  }

  const redirectUri = source.redirect_uris?.[0] ?? 'http://localhost';
  return new OAuth2Client(source.client_id, source.client_secret, redirectUri);
}

export async function readCredentialsFile(credentialsPath: string): Promise<unknown> {
  const raw = await fs.readFile(credentialsPath, 'utf8');
  return JSON.parse(raw);
}

export async function buildOAuthClientFromCredentialsFile(
  credentialsPath: string
): Promise<OAuth2Client> {
  const credentials = await readCredentialsFile(credentialsPath);
  return createOAuthClientFromCredentials({ credentials });
}

export function generateAuthUrlFromCredentials(credentials: unknown): {
  oauth2Client: OAuth2Client;
  authUrl: string;
} {
  const oauth2Client = createOAuthClientFromCredentials({ credentials });
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: [...GOOGLE_ACCOUNT_SCOPES],
    prompt: 'consent',
  });
  return { oauth2Client, authUrl };
}

export async function exchangeCodeForToken(
  credentials: unknown,
  authorizationCode: string
): Promise<Credentials> {
  const oauth2Client = createOAuthClientFromCredentials({ credentials });
  const { tokens } = await oauth2Client.getToken(authorizationCode);
  return tokens;
}

function sanitizeMessageIds(messageIds: string[]): string[] {
  return Array.from(
    new Set(
      messageIds
        .map((messageId) => messageId.trim())
        .filter((messageId) => messageId.length > 0)
    )
  );
}

function escapeDriveQueryValue(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export function buildDriveSearchQuery(query: string): string {
  const normalizedQuery = query.trim();
  if (!normalizedQuery) {
    throw new Error('Drive search query is required.');
  }

  const escapedQuery = escapeDriveQueryValue(normalizedQuery);
  return `trashed = false and (name contains '${escapedQuery}' or fullText contains '${escapedQuery}')`;
}

export function describeDriveApiError(error: unknown, fallback: string): string {
  const googleError = error as {
    code?: number;
    message?: string;
    response?: {
      data?: {
        error?: {
          message?: string;
          details?: Array<{
            '@type'?: string;
            reason?: string;
            metadata?: {
              activationUrl?: string;
              containerInfo?: string;
            };
          }>;
        };
      };
    };
  };

  const message = googleError.response?.data?.error?.message || googleError.message || fallback;
  const details = googleError.response?.data?.error?.details ?? [];
  const serviceDisabled = details.find((detail) => detail.reason === 'SERVICE_DISABLED');

  if (serviceDisabled) {
    const activationUrl = serviceDisabled.metadata?.activationUrl;
    const projectId = serviceDisabled.metadata?.containerInfo;
    const parts = ['Enable the Google Drive API in Google Cloud for this OAuth client project first.'];

    if (projectId) {
      parts.push(`Project: ${projectId}.`);
    }

    if (activationUrl) {
      parts.push(`Activation URL: ${activationUrl}`);
    }

    return parts.join(' ');
  }

  if (
    googleError.code === 403 &&
    /insufficient.*scope|insufficient.*permission/i.test(message)
  ) {
    return [
      'Drive access is not granted for this account yet.',
      'Re-run `begin_account_auth` and `finish_account_auth` to grant Google Drive metadata access.',
    ].join(' ');
  }

  return message;
}

const WORKSPACE_EXPORT_MAP: Record<string, { exportMime: string; ext: string; contentType: string }> = {
  'application/vnd.google-apps.document': {
    exportMime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    ext: '.docx',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  },
  'application/vnd.google-apps.spreadsheet': {
    exportMime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ext: '.xlsx',
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  'application/vnd.google-apps.presentation': {
    exportMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ext: '.pptx',
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  },
};

export class GmailAccountClient {
  readonly account: AccountConfig;
  readonly paths: AccountPaths;
  private readonly gmail: gmail_v1.Gmail;
  private readonly drive: drive_v3.Drive;
  private readonly sheets: sheets_v4.Sheets;
  private readonly docs: docs_v1.Docs;
  private readonly calendar: calendar_v3.Calendar;

  private constructor(
    account: AccountConfig,
    paths: AccountPaths,
    gmail: gmail_v1.Gmail,
    drive: drive_v3.Drive,
    sheets: sheets_v4.Sheets,
    docs: docs_v1.Docs,
    calendar: calendar_v3.Calendar,
  ) {
    this.account = account;
    this.paths = paths;
    this.gmail = gmail;
    this.drive = drive;
    this.sheets = sheets;
    this.docs = docs;
    this.calendar = calendar;
  }

  static async create(configRoot: string, account: AccountConfig): Promise<GmailAccountClient> {
    const paths = getAccountPaths(configRoot, account);

    const oauth2Client = await buildOAuthClientFromCredentialsFile(paths.credentialsPath);

    let cachedTokens: Credentials;
    try {
      const rawToken = await fs.readFile(paths.tokenPath, 'utf8');
      cachedTokens = JSON.parse(rawToken) as Credentials;
    } catch (error) {
      throw new Error(
        `Token file missing or invalid for account "${account.id}" at ${paths.tokenPath}: ${(error as Error).message}`
      );
    }

    oauth2Client.setCredentials(cachedTokens);
    oauth2Client.on('tokens', (incomingTokens) => {
      cachedTokens = { ...cachedTokens, ...incomingTokens };
      void fs
        .writeFile(paths.tokenPath, `${JSON.stringify(cachedTokens, null, 2)}\n`, 'utf8')
        .catch((error) => {
          console.error(
            `[ghub] Failed to persist refreshed token for account ${account.id}:`,
            error
          );
        });
    });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
    const drive = google.drive({ version: 'v3', auth: oauth2Client });
    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
    const docs = google.docs({ version: 'v1', auth: oauth2Client });
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
    return new GmailAccountClient(account, paths, gmail, drive, sheets, docs, calendar);
  }

  async getProfileEmail(): Promise<string> {
    const profile = await this.gmail.users.getProfile({ userId: 'me' });
    if (!profile.data.emailAddress) {
      throw new Error(`Gmail profile did not return an email address for account "${this.account.id}".`);
    }
    return profile.data.emailAddress;
  }

  async readEmails(query: string, maxResults: number, includeBody: boolean): Promise<ParsedEmail[]> {
    return this.fetchMessages(query, maxResults, includeBody);
  }

  async searchEmails(query: string, maxResults: number): Promise<ParsedEmail[]> {
    if (!query || query.trim() === '') {
      throw new Error('Search query is required.');
    }
    return this.fetchMessages(query, maxResults, false);
  }

  async searchDriveFiles(query: string, maxResults: number): Promise<DriveFileSummary[]> {
    const boundedMax = Math.max(1, Math.min(maxResults, 500));

    try {
      const response = await this.drive.files.list({
        q: buildDriveSearchQuery(query),
        pageSize: boundedMax,
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        orderBy: 'modifiedTime desc',
        fields:
          'files(id,name,mimeType,modifiedTime,webViewLink,owners(displayName,emailAddress))',
      });

      return (response.data.files ?? []).map((file) => ({
        id: file.id ?? '',
        name: file.name ?? '(untitled)',
        mimeType: file.mimeType ?? 'application/octet-stream',
        modifiedTime: file.modifiedTime ?? undefined,
        owners: (file.owners ?? []).map((owner) => owner.displayName || owner.emailAddress || '(unknown)'),
        webViewLink: file.webViewLink ?? undefined,
        accountId: this.account.id,
        accountEmail: this.account.email,
      }));
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive search failed.'));
    }
  }

  private async fetchMessages(
    query: string,
    maxResults: number,
    includeBody: boolean
  ): Promise<ParsedEmail[]> {
    const boundedMax = Math.max(1, Math.min(maxResults, 500));

    const listResponse = await this.gmail.users.messages.list({
      userId: 'me',
      q: query.trim() === '' ? undefined : query,
      maxResults: boundedMax,
    });

    const messageIds = (listResponse.data.messages ?? [])
      .map((message) => message.id)
      .filter((id): id is string => Boolean(id));

    if (messageIds.length === 0) {
      return [];
    }

    const fullMessages = await Promise.all(
      messageIds.map((messageId) =>
        this.gmail.users.messages.get({
          userId: 'me',
          id: messageId,
          format: 'full',
        })
      )
    );

    return fullMessages
      .map((response) => this.parseMessage(response.data, includeBody))
      .sort((a, b) => b.internalDate - a.internalDate);
  }

  async listAttachments(messageId: string): Promise<AttachmentMetadata[]> {
    if (!messageId || messageId.trim() === '') {
      throw new Error('message_id is required.');
    }

    const response = await this.gmail.users.messages.get({
      userId: 'me',
      id: messageId,
      format: 'full',
    });

    return extractAttachmentsMetadata(response.data.payload).filter((a) => !a.isInline);
  }

  async getAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ bytes: Buffer; metadata: AttachmentMetadata }> {
    if (!messageId || messageId.trim() === '') {
      throw new Error('message_id is required.');
    }
    if (!attachmentId || attachmentId.trim() === '') {
      throw new Error('attachment_id is required.');
    }

    const messageResponse = await this.gmail.users.messages.get({
      userId: 'me',
      id: messageId,
      format: 'full',
    });

    const part = findAttachmentPart(messageResponse.data.payload, attachmentId);
    if (!part || !part.filename) {
      throw new Error(
        `Attachment ${attachmentId} not found on message ${messageId}.`,
      );
    }

    const attachmentResponse = await this.gmail.users.messages.attachments.get({
      userId: 'me',
      messageId,
      id: attachmentId,
    });

    const data = attachmentResponse.data.data;
    if (!data) {
      throw new Error(`Attachment ${attachmentId} has no data payload.`);
    }

    const bytes = decodeBase64UrlBuffer(data);

    return {
      bytes,
      metadata: {
        id: attachmentId,
        filename: part.filename,
        contentType: part.mimeType ?? 'application/octet-stream',
        sizeBytes: Number(part.body?.size ?? bytes.length),
        isInline: isInlineDisposition(part.headers),
      },
    };
  }

  async getThread(threadId: string): Promise<ParsedThread> {
    if (!threadId || threadId.trim() === '') {
      throw new Error('thread_id is required.');
    }

    const threadResponse = await this.gmail.users.threads.get({
      userId: 'me',
      id: threadId,
      format: 'full',
    });

    const messages = (threadResponse.data.messages ?? [])
      .map((message) => this.parseMessage(message, true))
      .sort((a, b) => a.internalDate - b.internalDate);

    return {
      threadId,
      messages,
    };
  }

  async getLabels(): Promise<LabelInfo[]> {
    const labelsResponse = await this.gmail.users.labels.list({ userId: 'me' });
    return (labelsResponse.data.labels ?? []).map((label) => ({
      id: label.id ?? '',
      name: label.name ?? '(unnamed)',
      type: label.type ?? undefined,
      messagesTotal: label.messagesTotal ?? undefined,
    }));
  }

  async markAsRead(messageIds: string[]): Promise<number> {
    const ids = sanitizeMessageIds(messageIds);
    if (ids.length === 0) {
      throw new Error('message_ids must include at least one value.');
    }

    await this.gmail.users.messages.batchModify({
      userId: 'me',
      requestBody: {
        ids,
        removeLabelIds: ['UNREAD'],
      },
    });

    return ids.length;
  }

  async addLabels(messageIds: string[], labelIds: string[]): Promise<number> {
    const ids = sanitizeMessageIds(messageIds);
    const labels = labelIds.map((labelId) => labelId.trim()).filter(Boolean);

    if (ids.length === 0) throw new Error('message_ids must include at least one value.');
    if (labels.length === 0) throw new Error('label_ids must include at least one value.');

    await this.gmail.users.messages.batchModify({
      userId: 'me',
      requestBody: {
        ids,
        addLabelIds: labels,
      },
    });

    return ids.length;
  }

  async removeLabels(messageIds: string[], labelIds: string[]): Promise<number> {
    const ids = sanitizeMessageIds(messageIds);
    const labels = labelIds.map((labelId) => labelId.trim()).filter(Boolean);

    if (ids.length === 0) throw new Error('message_ids must include at least one value.');
    if (labels.length === 0) throw new Error('label_ids must include at least one value.');

    await this.gmail.users.messages.batchModify({
      userId: 'me',
      requestBody: {
        ids,
        removeLabelIds: labels,
      },
    });

    return ids.length;
  }

  async archiveEmails(messageIds: string[]): Promise<number> {
    const ids = sanitizeMessageIds(messageIds);
    if (ids.length === 0) throw new Error('message_ids must include at least one value.');

    await this.gmail.users.messages.batchModify({
      userId: 'me',
      requestBody: {
        ids,
        removeLabelIds: ['INBOX'],
      },
    });

    return ids.length;
  }

  async trashEmails(messageIds: string[]): Promise<number> {
    const ids = sanitizeMessageIds(messageIds);
    if (ids.length === 0) throw new Error('message_ids must include at least one value.');

    await Promise.all(
      ids.map((messageId) =>
        this.gmail.users.messages.trash({
          userId: 'me',
          id: messageId,
        })
      )
    );

    return ids.length;
  }

  async createLabel(
    name: string,
    labelListVisibility = 'labelShow',
    messageListVisibility = 'show'
  ): Promise<LabelInfo> {
    if (!name || name.trim() === '') {
      throw new Error('Label name is required.');
    }

    const response = await this.gmail.users.labels.create({
      userId: 'me',
      requestBody: {
        name: name.trim(),
        labelListVisibility,
        messageListVisibility,
      },
    });

    return {
      id: response.data.id ?? '',
      name: response.data.name ?? name,
      type: response.data.type ?? undefined,
      messagesTotal: response.data.messagesTotal ?? undefined,
    };
  }

  async deleteLabel(labelId: string): Promise<void> {
    if (!labelId || labelId.trim() === '') {
      throw new Error('label_id is required.');
    }

    await this.gmail.users.labels.delete({
      userId: 'me',
      id: labelId,
    });
  }

  async createFilter(
    criteria: gmail_v1.Schema$FilterCriteria,
    action: gmail_v1.Schema$FilterAction
  ): Promise<gmail_v1.Schema$Filter> {
    const response = await this.gmail.users.settings.filters.create({
      userId: 'me',
      requestBody: { criteria, action },
    });
    return response.data;
  }

  async listFilters(): Promise<gmail_v1.Schema$Filter[]> {
    const response = await this.gmail.users.settings.filters.list({ userId: 'me' });
    return response.data.filter ?? [];
  }

  async deleteFilter(filterId: string): Promise<void> {
    if (!filterId || filterId.trim() === '') {
      throw new Error('filter_id is required.');
    }
    await this.gmail.users.settings.filters.delete({
      userId: 'me',
      id: filterId.trim(),
    });
  }

  async createBlockFilter(
    sender: string,
    action: 'trash' | 'archive' | 'spam'
  ): Promise<gmail_v1.Schema$Filter> {
    const trimmed = sender.trim();
    if (!trimmed) throw new Error('sender is required.');

    const criteria: gmail_v1.Schema$FilterCriteria = { from: trimmed };
    const filterAction: gmail_v1.Schema$FilterAction =
      action === 'archive'
        ? { removeLabelIds: ['INBOX'] }
        : action === 'spam'
          ? { addLabelIds: ['SPAM'], removeLabelIds: ['INBOX'] }
          : { addLabelIds: ['TRASH'], removeLabelIds: ['INBOX', 'UNREAD'] };

    return this.createFilter(criteria, filterAction);
  }

  async modifyThread(
    threadId: string,
    modifications: { addLabelIds?: string[]; removeLabelIds?: string[] }
  ): Promise<void> {
    if (!threadId || threadId.trim() === '') {
      throw new Error('thread_id is required.');
    }
    const addLabelIds = (modifications.addLabelIds ?? [])
      .map((id) => id.trim())
      .filter(Boolean);
    const removeLabelIds = (modifications.removeLabelIds ?? [])
      .map((id) => id.trim())
      .filter(Boolean);
    if (addLabelIds.length === 0 && removeLabelIds.length === 0) {
      throw new Error('modifyThread requires at least one label add or remove.');
    }

    await this.gmail.users.threads.modify({
      userId: 'me',
      id: threadId.trim(),
      requestBody: { addLabelIds, removeLabelIds },
    });
  }

  async getThreadSubject(threadId: string): Promise<string> {
    if (!threadId || threadId.trim() === '') {
      throw new Error('thread_id is required.');
    }

    const response = await this.gmail.users.threads.get({
      userId: 'me',
      id: threadId.trim(),
      format: 'metadata',
      metadataHeaders: ['Subject'],
    });

    const firstMessage = response.data.messages?.[0];
    const raw = getHeaderValue(firstMessage?.payload?.headers, 'Subject');
    return raw.replace(/^(?:\s*(?:re|fwd?|aw)\s*:\s*)+/i, '').trim();
  }

  async getMessageHeaders(
    messageId: string,
    headerNames: string[]
  ): Promise<Record<string, string>> {
    if (!messageId || messageId.trim() === '') {
      throw new Error('message_id is required.');
    }
    const names = headerNames.map((name) => name.trim()).filter(Boolean);
    if (names.length === 0) {
      throw new Error('headerNames must include at least one value.');
    }

    const response = await this.gmail.users.messages.get({
      userId: 'me',
      id: messageId.trim(),
      format: 'metadata',
      metadataHeaders: names,
    });

    const headers = response.data.payload?.headers ?? [];
    const result: Record<string, string> = {};
    for (const name of names) {
      result[name] = getHeaderValue(headers, name);
    }
    return result;
  }

  async createDraft(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
    attachments?: EmailAttachment[];
    threadId?: string;
    inReplyTo?: string;
    references?: string;
  }): Promise<{ draftId: string; threadId?: string }> {
    const raw = await createRawEmailMessage(input);

    const message: { raw: string; threadId?: string } = { raw };
    if (input.threadId) message.threadId = input.threadId;

    const response = await this.gmail.users.drafts.create({
      userId: 'me',
      requestBody: { message },
    });

    return {
      draftId: response.data.id ?? '',
      threadId: response.data.message?.threadId ?? undefined,
    };
  }

  async deleteDrafts(draftIds: string[]): Promise<number> {
    const ids = draftIds.map((id) => id.trim()).filter(Boolean);
    if (ids.length === 0) throw new Error('draft_ids must include at least one value.');

    await Promise.all(
      ids.map((draftId) =>
        this.gmail.users.drafts.delete({
          userId: 'me',
          id: draftId,
        })
      )
    );

    return ids.length;
  }

  async sendDraft(draftId: string): Promise<{ messageId: string; threadId?: string }> {
    const response = await this.gmail.users.drafts.send({
      userId: 'me',
      requestBody: { id: draftId },
    });

    return {
      messageId: response.data.id ?? '',
      threadId: response.data.threadId ?? undefined,
    };
  }

  async listDrafts(maxResults = 20): Promise<Array<{ draftId: string; messageId: string; threadId?: string; subject: string; to: string; internalDate: number }>> {
    const listRes = await this.gmail.users.drafts.list({
      userId: 'me',
      maxResults,
    });

    const drafts = listRes.data.drafts ?? [];
    if (drafts.length === 0) return [];

    const details = await Promise.all(
      drafts.map((d) =>
        this.gmail.users.drafts.get({
          userId: 'me',
          id: d.id!,
          format: 'metadata',
        })
      )
    );

    const results = details.map((res) => {
      const headers = res.data.message?.payload?.headers ?? [];
      const get = (name: string) => headers.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value ?? '';
      return {
        draftId: res.data.id ?? '',
        messageId: res.data.message?.id ?? '',
        threadId: res.data.message?.threadId ?? undefined,
        subject: get('Subject'),
        to: get('To'),
        internalDate: Number(res.data.message?.internalDate ?? 0),
      };
    });

    return results.sort((a, b) => a.internalDate - b.internalDate);
  }

  async searchDrafts(query: string, maxResults = 20): Promise<Array<{ draftId: string; messageId: string; threadId?: string; subject: string; to: string; snippet: string }>> {
    const listRes = await this.gmail.users.drafts.list({
      userId: 'me',
      maxResults,
      q: query,
    });

    const drafts = listRes.data.drafts ?? [];
    if (drafts.length === 0) return [];

    const details = await Promise.all(
      drafts.map((d) =>
        this.gmail.users.drafts.get({
          userId: 'me',
          id: d.id!,
          format: 'metadata',
        })
      )
    );

    return details.map((res) => {
      const headers = res.data.message?.payload?.headers ?? [];
      const get = (name: string) => headers.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value ?? '';
      return {
        draftId: res.data.id ?? '',
        messageId: res.data.message?.id ?? '',
        threadId: res.data.message?.threadId ?? undefined,
        subject: get('Subject'),
        to: get('To'),
        snippet: res.data.message?.snippet ?? '',
      };
    });
  }

  async sendEmail(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
    attachments?: EmailAttachment[];
  }): Promise<{ messageId: string; threadId?: string }> {
    const raw = await createRawEmailMessage(input);

    const response = await this.gmail.users.messages.send({
      userId: 'me',
      requestBody: { raw },
    });

    return {
      messageId: response.data.id ?? '',
      threadId: response.data.threadId ?? undefined,
    };
  }

  private parseMessage(message: gmail_v1.Schema$Message, includeBody: boolean): ParsedEmail {
    const headers = message.payload?.headers;
    const internalDate = Number(message.internalDate ?? 0);
    const attachments = extractAttachmentsMetadata(message.payload).filter(
      (a) => !a.isInline,
    );

    return {
      id: message.id ?? '',
      threadId: message.threadId ?? '',
      snippet: message.snippet ?? '',
      from: getHeaderValue(headers, 'From'),
      to: getHeaderValue(headers, 'To'),
      cc: getHeaderValue(headers, 'Cc'),
      subject: getHeaderValue(headers, 'Subject') || '(no subject)',
      date: getHeaderValue(headers, 'Date'),
      internalDate: Number.isFinite(internalDate) ? internalDate : 0,
      messageHeaderId: getHeaderValue(headers, 'Message-ID'),
      inReplyTo: getHeaderValue(headers, 'In-Reply-To'),
      references: getHeaderValue(headers, 'References'),
      body: includeBody ? extractEmailBody(message.payload) : undefined,
      labels: message.labelIds ?? [],
      attachments,
      accountId: this.account.id,
      accountEmail: this.account.email,
    };
  }

  // ─── Drive ───────────────────────────────────────────────────────────────

  async listDriveFiles(options: {
    folderId?: string;
    query?: string;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ files: DriveFileSummary[]; nextPageToken?: string }> {
    const parts: string[] = ['trashed = false'];
    if (options.folderId?.trim()) parts.push(`'${options.folderId.trim()}' in parents`);
    if (options.query?.trim()) {
      const esc = escapeDriveQueryValue(options.query.trim());
      parts.push(`(name contains '${esc}' or fullText contains '${esc}')`);
    }

    try {
      const response = await this.drive.files.list({
        q: parts.join(' and '),
        pageSize: Math.max(1, Math.min(options.maxResults ?? 25, 500)),
        pageToken: options.pageToken,
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        orderBy: 'modifiedTime desc',
        fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink,owners(displayName,emailAddress))',
      });
      return {
        files: (response.data.files ?? []).map((f) => ({
          id: f.id ?? '',
          name: f.name ?? '(untitled)',
          mimeType: f.mimeType ?? 'application/octet-stream',
          modifiedTime: f.modifiedTime ?? undefined,
          owners: (f.owners ?? []).map((o) => o.displayName || o.emailAddress || '(unknown)'),
          webViewLink: f.webViewLink ?? undefined,
          accountId: this.account.id,
          accountEmail: this.account.email,
        })),
        nextPageToken: response.data.nextPageToken ?? undefined,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive list failed.'));
    }
  }

  async getDriveFile(fileId: string): Promise<DriveFileDetail> {
    if (!fileId?.trim()) throw new Error('file_id is required.');
    try {
      const response = await this.drive.files.get({
        fileId: fileId.trim(),
        supportsAllDrives: true,
        fields: 'id,name,mimeType,modifiedTime,createdTime,size,parents,owners(displayName,emailAddress),webViewLink,trashed,starred,shared,exportLinks',
      });
      const f = response.data;
      return {
        id: f.id ?? '',
        name: f.name ?? '(untitled)',
        mimeType: f.mimeType ?? 'application/octet-stream',
        modifiedTime: f.modifiedTime ?? undefined,
        createdTime: f.createdTime ?? undefined,
        size: f.size ? Number(f.size) : undefined,
        parents: f.parents ?? undefined,
        owners: (f.owners ?? []).map((o) => o.displayName || o.emailAddress || '(unknown)'),
        webViewLink: f.webViewLink ?? undefined,
        trashed: f.trashed ?? false,
        starred: f.starred ?? false,
        shared: f.shared ?? false,
        exportLinks: f.exportLinks ? f.exportLinks as Record<string, string> : undefined,
        accountId: this.account.id,
        accountEmail: this.account.email,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive file get failed.'));
    }
  }

  async getDriveFileContent(fileId: string): Promise<{ bytes: Buffer; contentType: string; filename: string }> {
    if (!fileId?.trim()) throw new Error('file_id is required.');
    const meta = await this.getDriveFile(fileId.trim());
    const wsExport = WORKSPACE_EXPORT_MAP[meta.mimeType];

    try {
      if (wsExport) {
        const response = await this.drive.files.export(
          { fileId: fileId.trim(), mimeType: wsExport.exportMime },
          { responseType: 'arraybuffer' },
        );
        return {
          bytes: Buffer.from(response.data as unknown as ArrayBuffer),
          contentType: wsExport.contentType,
          filename: `${meta.name}${wsExport.ext}`,
        };
      }
      const response = await this.drive.files.get(
        { fileId: fileId.trim(), alt: 'media', supportsAllDrives: true },
        { responseType: 'arraybuffer' },
      );
      return {
        bytes: Buffer.from(response.data as unknown as ArrayBuffer),
        contentType: meta.mimeType,
        filename: meta.name,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive file download failed.'));
    }
  }

  async uploadDriveFile(input: {
    localPath: string;
    name?: string;
    folderId?: string;
    mimeType?: string;
  }): Promise<DriveFileDetail> {
    const localPath = input.localPath.trim();
    if (!localPath) throw new Error('local_path is required.');
    const filename = input.name?.trim() || path.basename(localPath);
    const mimeType = input.mimeType?.trim() || inferContentType(filename);
    const requestBody: drive_v3.Schema$File = { name: filename };
    if (input.folderId?.trim()) requestBody.parents = [input.folderId.trim()];

    try {
      const response = await this.drive.files.create({
        supportsAllDrives: true,
        requestBody,
        media: { mimeType, body: createReadStream(localPath) },
        fields: 'id,name,mimeType,modifiedTime,createdTime,size,parents,owners(displayName,emailAddress),webViewLink,trashed,starred,shared',
      });
      const f = response.data;
      return {
        id: f.id ?? '',
        name: f.name ?? filename,
        mimeType: f.mimeType ?? mimeType,
        modifiedTime: f.modifiedTime ?? undefined,
        createdTime: f.createdTime ?? undefined,
        size: f.size ? Number(f.size) : undefined,
        parents: f.parents ?? undefined,
        owners: (f.owners ?? []).map((o) => o.displayName || o.emailAddress || '(unknown)'),
        webViewLink: f.webViewLink ?? undefined,
        trashed: false,
        starred: false,
        shared: false,
        accountId: this.account.id,
        accountEmail: this.account.email,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive upload failed.'));
    }
  }

  async createDriveFolder(name: string, parentId?: string): Promise<DriveFileSummary> {
    if (!name?.trim()) throw new Error('name is required.');
    const requestBody: drive_v3.Schema$File = {
      name: name.trim(),
      mimeType: 'application/vnd.google-apps.folder',
    };
    if (parentId?.trim()) requestBody.parents = [parentId.trim()];

    try {
      const response = await this.drive.files.create({
        supportsAllDrives: true,
        requestBody,
        fields: 'id,name,mimeType,modifiedTime,webViewLink,owners(displayName,emailAddress)',
      });
      const f = response.data;
      return {
        id: f.id ?? '',
        name: f.name ?? name,
        mimeType: f.mimeType ?? 'application/vnd.google-apps.folder',
        modifiedTime: f.modifiedTime ?? undefined,
        owners: (f.owners ?? []).map((o) => o.displayName || o.emailAddress || '(unknown)'),
        webViewLink: f.webViewLink ?? undefined,
        accountId: this.account.id,
        accountEmail: this.account.email,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Create folder failed.'));
    }
  }

  async updateDriveFile(
    fileId: string,
    updates: { name?: string; addParents?: string; removeParents?: string; starred?: boolean; description?: string },
  ): Promise<DriveFileSummary> {
    if (!fileId?.trim()) throw new Error('file_id is required.');
    const requestBody: drive_v3.Schema$File = {};
    if (updates.name !== undefined) requestBody.name = updates.name.trim();
    if (updates.starred !== undefined) requestBody.starred = updates.starred;
    if (updates.description !== undefined) requestBody.description = updates.description;

    try {
      const response = await this.drive.files.update({
        fileId: fileId.trim(),
        supportsAllDrives: true,
        addParents: updates.addParents?.trim() || undefined,
        removeParents: updates.removeParents?.trim() || undefined,
        requestBody,
        fields: 'id,name,mimeType,modifiedTime,webViewLink,owners(displayName,emailAddress)',
      });
      const f = response.data;
      return {
        id: f.id ?? '',
        name: f.name ?? '(untitled)',
        mimeType: f.mimeType ?? 'application/octet-stream',
        modifiedTime: f.modifiedTime ?? undefined,
        owners: (f.owners ?? []).map((o) => o.displayName || o.emailAddress || '(unknown)'),
        webViewLink: f.webViewLink ?? undefined,
        accountId: this.account.id,
        accountEmail: this.account.email,
      };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive file update failed.'));
    }
  }

  async trashDriveFile(fileId: string): Promise<void> {
    if (!fileId?.trim()) throw new Error('file_id is required.');
    try {
      await this.drive.files.update({
        fileId: fileId.trim(),
        supportsAllDrives: true,
        requestBody: { trashed: true },
      });
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive trash failed.'));
    }
  }

  async shareDriveFile(
    fileId: string,
    input: { email?: string; role: string; type: string; sendNotification?: boolean; notificationMessage?: string },
  ): Promise<{ permissionId: string }> {
    if (!fileId?.trim()) throw new Error('file_id is required.');
    try {
      const response = await this.drive.permissions.create({
        fileId: fileId.trim(),
        supportsAllDrives: true,
        sendNotificationEmail: input.sendNotification ?? true,
        emailMessage: input.notificationMessage,
        requestBody: { type: input.type, role: input.role, emailAddress: input.email },
      });
      return { permissionId: response.data.id ?? '' };
    } catch (error) {
      throw new Error(describeDriveApiError(error, 'Drive share failed.'));
    }
  }

  // ─── Sheets ──────────────────────────────────────────────────────────────

  async getSheetsMetadata(spreadsheetId: string): Promise<SpreadsheetMetadata> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const response = await this.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId.trim(),
      fields: 'spreadsheetId,spreadsheetUrl,properties/title,sheets(properties(sheetId,title,index,gridProperties))',
    });
    return {
      id: response.data.spreadsheetId ?? '',
      title: response.data.properties?.title ?? '(untitled)',
      url: response.data.spreadsheetUrl ?? `https://docs.google.com/spreadsheets/d/${response.data.spreadsheetId}`,
      sheets: (response.data.sheets ?? []).map((s) => ({
        sheetId: s.properties?.sheetId ?? 0,
        title: s.properties?.title ?? '(untitled)',
        index: s.properties?.index ?? 0,
        rowCount: s.properties?.gridProperties?.rowCount ?? 0,
        columnCount: s.properties?.gridProperties?.columnCount ?? 0,
      })),
    };
  }

  async readSheetValues(
    spreadsheetId: string,
    range: string,
    valueRenderOption = 'FORMATTED_VALUE',
  ): Promise<{ range: string; values: unknown[][] }> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    if (!range?.trim()) throw new Error('range is required.');
    const response = await this.sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId.trim(),
      range: range.trim(),
      valueRenderOption,
    });
    return {
      range: response.data.range ?? range,
      values: (response.data.values ?? []) as unknown[][],
    };
  }

  async writeSheetValues(
    spreadsheetId: string,
    range: string,
    values: unknown[][],
    valueInputOption = 'USER_ENTERED',
  ): Promise<{ updatedRange: string; updatedRows: number; updatedCells: number }> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    if (!range?.trim()) throw new Error('range is required.');
    const response = await this.sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId.trim(),
      range: range.trim(),
      valueInputOption,
      requestBody: { values },
    });
    return {
      updatedRange: response.data.updatedRange ?? range,
      updatedRows: response.data.updatedRows ?? 0,
      updatedCells: response.data.updatedCells ?? 0,
    };
  }

  async appendSheetValues(
    spreadsheetId: string,
    range: string,
    values: unknown[][],
    valueInputOption = 'USER_ENTERED',
  ): Promise<{ updatedRange: string; updatedRows: number }> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    if (!range?.trim()) throw new Error('range is required.');
    const response = await this.sheets.spreadsheets.values.append({
      spreadsheetId: spreadsheetId.trim(),
      range: range.trim(),
      valueInputOption,
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values },
    });
    return {
      updatedRange: response.data.updates?.updatedRange ?? range,
      updatedRows: response.data.updates?.updatedRows ?? 0,
    };
  }

  async createSpreadsheet(title: string): Promise<{ id: string; url: string }> {
    if (!title?.trim()) throw new Error('title is required.');
    const response = await this.sheets.spreadsheets.create({
      requestBody: { properties: { title: title.trim() } },
    });
    return {
      id: response.data.spreadsheetId ?? '',
      url: response.data.spreadsheetUrl ?? `https://docs.google.com/spreadsheets/d/${response.data.spreadsheetId}`,
    };
  }

  private async getSheetIdByTitle(spreadsheetId: string, sheetTitle: string): Promise<number> {
    const meta = await this.getSheetsMetadata(spreadsheetId);
    const sheet = meta.sheets.find((s) => s.title === sheetTitle);
    if (!sheet) throw new Error(`Sheet tab "${sheetTitle}" not found in spreadsheet.`);
    return sheet.sheetId;
  }

  private async sheetsBatchUpdate(
    spreadsheetId: string,
    requests: sheets_v4.Schema$Request[],
  ): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
    const response = await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId.trim(),
      requestBody: { requests },
    });
    return response.data;
  }

  async addSheetTab(spreadsheetId: string, title: string, index?: number): Promise<{ sheetId: number; title: string }> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    if (!title?.trim()) throw new Error('title is required.');
    const props: sheets_v4.Schema$SheetProperties = { title: title.trim() };
    if (index !== undefined) props.index = index;
    const result = await this.sheetsBatchUpdate(spreadsheetId, [{ addSheet: { properties: props } }]);
    const added = result.replies?.[0]?.addSheet?.properties;
    return { sheetId: added?.sheetId ?? 0, title: added?.title ?? title };
  }

  async renameSheetTab(spreadsheetId: string, currentTitle: string, newTitle: string): Promise<void> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, currentTitle);
    await this.sheetsBatchUpdate(spreadsheetId, [{
      updateSheetProperties: {
        properties: { sheetId, title: newTitle.trim() },
        fields: 'title',
      },
    }]);
  }

  async deleteSheetTab(spreadsheetId: string, sheetTitle: string): Promise<void> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
    await this.sheetsBatchUpdate(spreadsheetId, [{ deleteSheet: { sheetId } }]);
  }

  async insertDimension(
    spreadsheetId: string,
    sheetTitle: string,
    dimension: 'ROWS' | 'COLUMNS',
    startIndex: number,
    count: number,
  ): Promise<void> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
    await this.sheetsBatchUpdate(spreadsheetId, [{
      insertDimension: {
        range: { sheetId, dimension, startIndex, endIndex: startIndex + count },
        inheritFromBefore: startIndex > 0,
      },
    }]);
  }

  async deleteDimension(
    spreadsheetId: string,
    sheetTitle: string,
    dimension: 'ROWS' | 'COLUMNS',
    startIndex: number,
    count: number,
  ): Promise<void> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
    await this.sheetsBatchUpdate(spreadsheetId, [{
      deleteDimension: {
        range: { sheetId, dimension, startIndex, endIndex: startIndex + count },
      },
    }]);
  }

  private parseA1Range(range: string, sheetId: number): sheets_v4.Schema$GridRange {
    const colToIndex = (letters: string): number => {
      let index = 0;
      for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
      return index - 1;
    };
    const withoutSheet = range.includes('!') ? range.split('!')[1] ?? range : range;
    const [startCell, endCell] = withoutSheet.split(':');
    const parseCell = (cell: string): { row: number; col: number } => {
      const m = (cell ?? '').match(/^([A-Za-z]+)(\d+)$/);
      if (!m) return { row: 0, col: 0 };
      return { col: colToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
    };
    const s = parseCell(startCell ?? '');
    const e = endCell ? parseCell(endCell) : s;
    return { sheetId, startRowIndex: s.row, endRowIndex: e.row + 1, startColumnIndex: s.col, endColumnIndex: e.col + 1 };
  }

  private hexToColor(hex: string): sheets_v4.Schema$Color {
    const h = hex.replace('#', '');
    return {
      red: parseInt(h.substring(0, 2), 16) / 255,
      green: parseInt(h.substring(2, 4), 16) / 255,
      blue: parseInt(h.substring(4, 6), 16) / 255,
    };
  }

  async formatCells(
    spreadsheetId: string,
    sheetTitle: string,
    range: string,
    format: {
      bold?: boolean;
      italic?: boolean;
      fontSize?: number;
      backgroundColor?: string;
      textColor?: string;
      horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT';
      numberFormat?: string;
      wrapStrategy?: 'OVERFLOW_CELL' | 'LEGACY_WRAP' | 'CLIP' | 'WRAP';
    },
  ): Promise<void> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
    const gridRange = this.parseA1Range(range, sheetId);
    const cellFormat: sheets_v4.Schema$CellFormat = {};
    const fields: string[] = [];

    if (format.bold !== undefined || format.italic !== undefined || format.fontSize !== undefined || format.textColor !== undefined) {
      cellFormat.textFormat = {};
      if (format.bold !== undefined) { cellFormat.textFormat.bold = format.bold; fields.push('userEnteredFormat.textFormat.bold'); }
      if (format.italic !== undefined) { cellFormat.textFormat.italic = format.italic; fields.push('userEnteredFormat.textFormat.italic'); }
      if (format.fontSize !== undefined) { cellFormat.textFormat.fontSize = format.fontSize; fields.push('userEnteredFormat.textFormat.fontSize'); }
      if (format.textColor) { cellFormat.textFormat.foregroundColor = this.hexToColor(format.textColor); fields.push('userEnteredFormat.textFormat.foregroundColor'); }
    }
    if (format.backgroundColor) { cellFormat.backgroundColor = this.hexToColor(format.backgroundColor); fields.push('userEnteredFormat.backgroundColor'); }
    if (format.horizontalAlignment) { cellFormat.horizontalAlignment = format.horizontalAlignment; fields.push('userEnteredFormat.horizontalAlignment'); }
    if (format.numberFormat) { cellFormat.numberFormat = { type: 'NUMBER', pattern: format.numberFormat }; fields.push('userEnteredFormat.numberFormat'); }
    if (format.wrapStrategy) { cellFormat.wrapStrategy = format.wrapStrategy; fields.push('userEnteredFormat.wrapStrategy'); }

    if (fields.length === 0) return;
    await this.sheetsBatchUpdate(spreadsheetId, [{
      repeatCell: { range: gridRange, cell: { userEnteredFormat: cellFormat }, fields: fields.join(',') },
    }]);
  }

  async addChart(
    spreadsheetId: string,
    sheetTitle: string,
    options: {
      chartType: 'BAR' | 'LINE' | 'PIE' | 'COLUMN' | 'AREA' | 'SCATTER';
      dataRange: string;
      title?: string;
      anchorRow?: number;
      anchorCol?: number;
      widthPixels?: number;
      heightPixels?: number;
    },
  ): Promise<{ chartId: number }> {
    if (!spreadsheetId?.trim()) throw new Error('spreadsheet_id is required.');
    const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
    const dataRange = this.parseA1Range(options.dataRange, sheetId);

    const domainRange: sheets_v4.Schema$GridRange = { ...dataRange, endColumnIndex: dataRange.startColumnIndex! + 1 };
    const seriesRange: sheets_v4.Schema$GridRange = { ...dataRange, startColumnIndex: dataRange.startColumnIndex! + 1 };

    let spec: sheets_v4.Schema$ChartSpec;
    if (options.chartType === 'PIE') {
      spec = {
        title: options.title,
        pieChart: {
          legendPosition: 'RIGHT_LEGEND',
          domain: { sourceRange: { sources: [domainRange] } },
          series: { sourceRange: { sources: [seriesRange] } },
        },
      };
    } else {
      spec = {
        title: options.title,
        basicChart: {
          chartType: options.chartType as sheets_v4.Schema$BasicChartSpec['chartType'],
          legendPosition: 'BOTTOM_LEGEND',
          axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS' }],
          domains: [{ domain: { sourceRange: { sources: [domainRange] } } }],
          series: [{ series: { sourceRange: { sources: [seriesRange] } }, targetAxis: 'LEFT_AXIS' }],
          headerCount: 1,
        },
      };
    }

    const result = await this.sheetsBatchUpdate(spreadsheetId, [{
      addChart: {
        chart: {
          spec,
          position: {
            overlayPosition: {
              anchorCell: { sheetId, rowIndex: options.anchorRow ?? 0, columnIndex: options.anchorCol ?? 0 },
              widthPixels: options.widthPixels ?? 600,
              heightPixels: options.heightPixels ?? 400,
            },
          },
        },
      },
    }]);

    return { chartId: result.replies?.[0]?.addChart?.chart?.chartId ?? 0 };
  }

  // ─── Docs ─────────────────────────────────────────────────────────────────

  async getDocument(documentId: string): Promise<{ documentId: string; title: string; body: string; url: string }> {
    if (!documentId?.trim()) throw new Error('document_id is required.');
    const response = await this.docs.documents.get({ documentId: documentId.trim() });
    const doc = response.data;

    const parts: string[] = [];
    const extract = (elements: docs_v1.Schema$StructuralElement[]): void => {
      for (const el of elements) {
        if (el.paragraph) {
          for (const pe of el.paragraph.elements ?? []) {
            if (pe.textRun?.content) parts.push(pe.textRun.content);
          }
        } else if (el.table) {
          for (const row of el.table.tableRows ?? []) {
            for (const cell of row.tableCells ?? []) extract(cell.content ?? []);
          }
        }
      }
    };
    extract(doc.body?.content ?? []);

    return {
      documentId: doc.documentId ?? documentId,
      title: doc.title ?? '(untitled)',
      body: parts.join(''),
      url: `https://docs.google.com/document/d/${doc.documentId}/edit`,
    };
  }

  async createDocument(title: string): Promise<{ documentId: string; url: string }> {
    if (!title?.trim()) throw new Error('title is required.');
    const response = await this.docs.documents.create({ requestBody: { title: title.trim() } });
    return {
      documentId: response.data.documentId ?? '',
      url: `https://docs.google.com/document/d/${response.data.documentId}/edit`,
    };
  }

  async appendToDocument(
    documentId: string,
    text: string,
    options?: { style?: 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3'; bold?: boolean; italic?: boolean },
  ): Promise<void> {
    if (!documentId?.trim()) throw new Error('document_id is required.');
    const doc = await this.docs.documents.get({ documentId: documentId.trim() });
    const endIndex = doc.data.body?.content?.slice(-1)?.[0]?.endIndex ?? 1;
    const insertIndex = endIndex - 1;
    const insertedText = text.endsWith('\n') ? text : text + '\n';

    const requests: docs_v1.Schema$Request[] = [
      { insertText: { location: { index: insertIndex }, text: insertedText } },
    ];

    if (options?.style && options.style !== 'NORMAL_TEXT') {
      requests.push({
        updateParagraphStyle: {
          range: { startIndex: insertIndex, endIndex: insertIndex + insertedText.length },
          paragraphStyle: { namedStyleType: options.style },
          fields: 'namedStyleType',
        },
      });
    }

    const styleFields: string[] = [];
    const textStyle: docs_v1.Schema$TextStyle = {};
    if (options?.bold !== undefined) { textStyle.bold = options.bold; styleFields.push('bold'); }
    if (options?.italic !== undefined) { textStyle.italic = options.italic; styleFields.push('italic'); }
    if (styleFields.length > 0) {
      requests.push({
        updateTextStyle: {
          range: { startIndex: insertIndex, endIndex: insertIndex + text.length },
          textStyle,
          fields: styleFields.join(','),
        },
      });
    }

    await this.docs.documents.batchUpdate({ documentId: documentId.trim(), requestBody: { requests } });
  }

  async replaceInDocument(
    documentId: string,
    findText: string,
    replaceText: string,
    matchCase = false,
  ): Promise<{ occurrencesChanged: number }> {
    if (!documentId?.trim()) throw new Error('document_id is required.');
    const response = await this.docs.documents.batchUpdate({
      documentId: documentId.trim(),
      requestBody: {
        requests: [{ replaceAllText: { containsText: { text: findText, matchCase }, replaceText } }],
      },
    });
    return { occurrencesChanged: response.data.replies?.[0]?.replaceAllText?.occurrencesChanged ?? 0 };
  }

  async insertTableInDocument(documentId: string, rows: number, columns: number): Promise<void> {
    if (!documentId?.trim()) throw new Error('document_id is required.');
    const doc = await this.docs.documents.get({ documentId: documentId.trim() });
    const endIndex = doc.data.body?.content?.slice(-1)?.[0]?.endIndex ?? 1;
    await this.docs.documents.batchUpdate({
      documentId: documentId.trim(),
      requestBody: {
        requests: [{ insertTable: { rows, columns, location: { index: endIndex - 1 } } }],
      },
    });
  }

  async applyDocHeadingStyle(
    documentId: string,
    startIndex: number,
    endIndex: number,
    style: 'NORMAL_TEXT' | 'HEADING_1' | 'HEADING_2' | 'HEADING_3' | 'HEADING_4',
  ): Promise<void> {
    if (!documentId?.trim()) throw new Error('document_id is required.');
    await this.docs.documents.batchUpdate({
      documentId: documentId.trim(),
      requestBody: {
        requests: [{
          updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: { namedStyleType: style },
            fields: 'namedStyleType',
          },
        }],
      },
    });
  }

  // ─── Calendar ─────────────────────────────────────────────────────────────

  async listCalendars(): Promise<CalendarInfo[]> {
    const response = await this.calendar.calendarList.list({});
    return (response.data.items ?? []).map((cal) => ({
      id: cal.id ?? '',
      summary: cal.summary ?? '(untitled)',
      description: cal.description ?? undefined,
      primary: cal.primary ?? false,
      accessRole: cal.accessRole ?? undefined,
      backgroundColor: cal.backgroundColor ?? undefined,
      timeZone: cal.timeZone ?? undefined,
      accountId: this.account.id,
      accountEmail: this.account.email,
    }));
  }

  async listCalendarEvents(
    calendarId: string,
    options: { timeMin?: string; timeMax?: string; query?: string; maxResults?: number; singleEvents?: boolean },
  ): Promise<CalendarEvent[]> {
    const cid = (calendarId || 'primary').trim();
    const response = await this.calendar.events.list({
      calendarId: cid,
      timeMin: options.timeMin,
      timeMax: options.timeMax,
      q: options.query?.trim() || undefined,
      maxResults: Math.max(1, Math.min(options.maxResults ?? 25, 250)),
      singleEvents: options.singleEvents ?? true,
      orderBy: options.singleEvents !== false ? 'startTime' : undefined,
    });
    return (response.data.items ?? []).map((e) => this.parseCalendarEvent(e, cid));
  }

  async getCalendarEvent(calendarId: string, eventId: string): Promise<CalendarEvent> {
    if (!calendarId?.trim()) throw new Error('calendar_id is required.');
    if (!eventId?.trim()) throw new Error('event_id is required.');
    const response = await this.calendar.events.get({ calendarId: calendarId.trim(), eventId: eventId.trim() });
    return this.parseCalendarEvent(response.data, calendarId);
  }

  async createCalendarEvent(
    calendarId: string,
    input: {
      summary: string;
      description?: string;
      location?: string;
      start: { dateTime?: string; date?: string; timeZone?: string };
      end: { dateTime?: string; date?: string; timeZone?: string };
      attendees?: string[];
      recurrence?: string[];
      sendNotifications?: boolean;
    },
  ): Promise<CalendarEvent> {
    const cid = (calendarId || 'primary').trim();
    const response = await this.calendar.events.insert({
      calendarId: cid,
      sendNotifications: input.sendNotifications ?? true,
      requestBody: {
        summary: input.summary,
        description: input.description,
        location: input.location,
        start: input.start,
        end: input.end,
        attendees: input.attendees?.map((email) => ({ email })),
        recurrence: input.recurrence,
      },
    });
    return this.parseCalendarEvent(response.data, cid);
  }

  async updateCalendarEvent(
    calendarId: string,
    eventId: string,
    updates: {
      summary?: string;
      description?: string;
      location?: string;
      start?: { dateTime?: string; date?: string; timeZone?: string };
      end?: { dateTime?: string; date?: string; timeZone?: string };
      attendees?: string[];
      status?: string;
      sendNotifications?: boolean;
    },
  ): Promise<CalendarEvent> {
    if (!calendarId?.trim()) throw new Error('calendar_id is required.');
    if (!eventId?.trim()) throw new Error('event_id is required.');
    const requestBody: calendar_v3.Schema$Event = {};
    if (updates.summary !== undefined) requestBody.summary = updates.summary;
    if (updates.description !== undefined) requestBody.description = updates.description;
    if (updates.location !== undefined) requestBody.location = updates.location;
    if (updates.start !== undefined) requestBody.start = updates.start;
    if (updates.end !== undefined) requestBody.end = updates.end;
    if (updates.status !== undefined) requestBody.status = updates.status;
    if (updates.attendees !== undefined) requestBody.attendees = updates.attendees.map((email) => ({ email }));

    const response = await this.calendar.events.patch({
      calendarId: calendarId.trim(),
      eventId: eventId.trim(),
      sendNotifications: updates.sendNotifications ?? true,
      requestBody,
    });
    return this.parseCalendarEvent(response.data, calendarId);
  }

  async deleteCalendarEvent(calendarId: string, eventId: string, sendNotifications = true): Promise<void> {
    if (!calendarId?.trim()) throw new Error('calendar_id is required.');
    if (!eventId?.trim()) throw new Error('event_id is required.');
    await this.calendar.events.delete({ calendarId: calendarId.trim(), eventId: eventId.trim(), sendNotifications });
  }

  private parseCalendarEvent(event: calendar_v3.Schema$Event, calendarId: string): CalendarEvent {
    return {
      id: event.id ?? '',
      calendarId,
      summary: event.summary ?? '(no title)',
      description: event.description ?? undefined,
      location: event.location ?? undefined,
      start: {
        dateTime: event.start?.dateTime ?? undefined,
        date: event.start?.date ?? undefined,
        timeZone: event.start?.timeZone ?? undefined,
      },
      end: {
        dateTime: event.end?.dateTime ?? undefined,
        date: event.end?.date ?? undefined,
        timeZone: event.end?.timeZone ?? undefined,
      },
      attendees: (event.attendees ?? []).map((a) => ({
        email: a.email ?? '',
        displayName: a.displayName ?? undefined,
        responseStatus: a.responseStatus ?? undefined,
        self: a.self ?? false,
      })),
      organizer: event.organizer
        ? { email: event.organizer.email ?? '', displayName: event.organizer.displayName ?? undefined, self: event.organizer.self ?? false }
        : undefined,
      status: event.status ?? undefined,
      htmlLink: event.htmlLink ?? undefined,
      recurrence: event.recurrence ?? undefined,
      recurringEventId: event.recurringEventId ?? undefined,
      created: event.created ?? undefined,
      updated: event.updated ?? undefined,
      conferenceData: event.conferenceData?.entryPoints?.length
        ? { entryPoints: event.conferenceData.entryPoints.map((ep) => ({ uri: ep.uri ?? '', entryPointType: ep.entryPointType ?? '' })) }
        : undefined,
      accountId: this.account.id,
      accountEmail: this.account.email,
    };
  }
}
