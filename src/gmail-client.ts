import { OAuth2Client, type Credentials } from 'google-auth-library';
import { google, type gmail_v1 } from 'googleapis';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { AccountConfig, type AccountPaths, getAccountPaths } from './config.js';

export const GMAIL_SCOPES = [
  // Broadest Gmail OAuth scope Google allows (full mailbox access).
  'https://mail.google.com/',
] as const;

/**
 * Publicly exposed attachment descriptor. Safe to persist in `ParsedEmail`
 * and hand across process boundaries; `partId` is stable across
 * `messages.get` calls for the same message.
 *
 * Gmail's ephemeral `body.attachmentId` is NOT included — it rotates on
 * every `messages.get` and is only valid within a single fetch. To
 * download an attachment, pass the stable `partId` to `download_attachment`
 * / `GmailAccountClient.downloadByPartId`, which re-fetches and resolves
 * internally.
 */
export interface AttachmentMetadata {
  /** Stable Gmail MIME part id (e.g. `"0.1"`). Use as the download key. */
  partId: string;
  filename: string;
  mimeType: string;
  size: number;
}

/** Internal shape carrying the ephemeral Gmail `body.attachmentId`. Never export. */
interface AttachmentPart extends AttachmentMetadata {
  /** Ephemeral Gmail body.attachmentId — rotates per messages.get. Never persist. */
  attachmentId: string;
}

export interface ParsedEmail {
  id: string;
  threadId: string;
  snippet: string;
  from: string;
  to: string;
  subject: string;
  date: string;
  internalDate: number;
  body?: string;
  labels: string[];
  attachments: readonly AttachmentMetadata[];
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

interface OAuthClientOptions {
  credentials: unknown;
}

export interface EmailAttachment {
  path: string;
  filename?: string;
  contentType?: string;
}

function decodeBase64Url(value: string): string {
  return decodeBase64UrlToBuffer(value).toString('utf8');
}

function decodeBase64UrlToBuffer(value: string): Buffer {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
  return Buffer.from(padded, 'base64');
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

/**
 * Walks a Gmail `messages.get(format: 'full')` payload and collects every
 * real attachment as an internal `AttachmentPart` (includes the ephemeral
 * `body.attachmentId`).
 *
 * **Exported for unit tests only.** Production callers should go through
 * `listAttachments` / `downloadByPartId` on `GmailAccountClient`, which
 * return the public `AttachmentMetadata` shape.
 *
 * Filters:
 * - requires `filename` and `body.attachmentId` (skips body parts)
 * - skips parts marked `Content-Disposition: inline` (embedded logos,
 *   tracking pixels, styled icons)
 *
 * Normalization:
 * - root payload's empty `partId` becomes `"0"`
 * - non-finite / negative `body.size` becomes `0`
 */
export function extractAttachmentParts(
  payload?: gmail_v1.Schema$MessagePart
): AttachmentPart[] {
  if (!payload) return [];

  const results: AttachmentPart[] = [];
  // Seed with the root payload itself — some messages (bare PDF, some forwards)
  // carry filename + body.attachmentId on the root and have no child parts.
  const stack: gmail_v1.Schema$MessagePart[] = [payload];
  while (stack.length > 0) {
    const part = stack.shift();
    if (!part) continue;

    if (part.parts?.length) {
      stack.push(...part.parts);
    }

    const filename = part.filename?.trim();
    const attachmentId = part.body?.attachmentId;
    if (!filename || !attachmentId) continue;

    // Only surface real attachments. Content-Disposition: inline marks embedded
    // body content (logos, tracking pixels, styled icons) that a user doesn't
    // think of as an "attachment" even though it has a filename + attachmentId.
    const disposition = getHeaderValue(part.headers, 'Content-Disposition').toLowerCase();
    if (disposition.startsWith('inline')) continue;

    // Gmail leaves `partId` empty only on the root payload. Substitute `"0"`
    // so the bare-payload case (root carries the attachment, no children)
    // still has a stable download key; real child ids are never shadowed
    // because this branch only fires when the root itself is the attachment.
    const partId = part.partId || '0';

    // Clamp size: malformed Gmail responses could produce NaN or a negative
    // number, which would silently bypass the MAX_ATTACHMENT_BYTES cap
    // downstream (NaN > maxBytes is false). Coerce both cases to 0.
    const rawSize = Number(part.body?.size ?? 0);
    const size = Number.isFinite(rawSize) && rawSize >= 0 ? Math.floor(rawSize) : 0;

    results.push({
      partId,
      attachmentId,
      filename,
      mimeType: part.mimeType ?? 'application/octet-stream',
      size,
    });
  }
  return results;
}

function toAttachmentMetadata(part: AttachmentPart): AttachmentMetadata {
  return {
    partId: part.partId,
    filename: part.filename,
    mimeType: part.mimeType,
    size: part.size,
  };
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
    scope: [...GMAIL_SCOPES],
    prompt: 'consent',
    include_granted_scopes: true,
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

export class GmailAccountClient {
  readonly account: AccountConfig;
  readonly paths: AccountPaths;
  private readonly gmail: gmail_v1.Gmail;

  private constructor(account: AccountConfig, paths: AccountPaths, gmail: gmail_v1.Gmail) {
    this.account = account;
    this.paths = paths;
    this.gmail = gmail;
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
            `[gmail-multi-inbox-mcp] Failed to persist refreshed token for account ${account.id}:`,
            error
          );
        });
    });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
    return new GmailAccountClient(account, paths, gmail);
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

  private async fetchMessages(
    query: string,
    maxResults: number,
    includeBody: boolean
  ): Promise<ParsedEmail[]> {
    const boundedMax = Math.max(1, Math.min(maxResults, 100));

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

  async createDraft(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
    attachments?: EmailAttachment[];
  }): Promise<{ draftId: string; threadId?: string }> {
    const raw = await createRawEmailMessage(input);

    const response = await this.gmail.users.drafts.create({
      userId: 'me',
      requestBody: {
        message: { raw },
      },
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

  async listAttachments(messageId: string): Promise<AttachmentMetadata[]> {
    const trimmed = messageId.trim();
    if (!trimmed) throw new Error('message_id is required.');

    const response = await this.gmail.users.messages.get({
      userId: 'me',
      id: trimmed,
      format: 'full',
    });
    return extractAttachmentParts(response.data.payload).map(toAttachmentMetadata);
  }

  /**
   * Resolve a stable `partId` to a fresh `body.attachmentId` and fetch the
   * decoded bytes. Encapsulates the rotating-id round trip so callers never
   * need to hold an ephemeral `attachmentId`. Enforces `maxBytes` against
   * the metadata reported by Gmail before pulling the payload.
   *
   * Returns a discriminated union so expected failure modes (not found,
   * too large) never throw — they surface to the MCP host as clean text
   * results instead of handler crashes.
   */
  async downloadByPartId(
    messageId: string,
    partId: string,
    maxBytes: number
  ): Promise<
    | { kind: 'ok'; data: Buffer; metadata: AttachmentMetadata }
    | { kind: 'not_found' }
    | { kind: 'too_large'; metadata: AttachmentMetadata }
  > {
    const trimmedMessageId = messageId.trim();
    const trimmedPartId = partId.trim();
    if (!trimmedMessageId) throw new Error('message_id is required.');
    if (!trimmedPartId) throw new Error('part_id is required.');

    const messageResponse = await this.gmail.users.messages.get({
      userId: 'me',
      id: trimmedMessageId,
      format: 'full',
    });

    const parts = extractAttachmentParts(messageResponse.data.payload);
    const part = parts.find((p) => p.partId === trimmedPartId);
    if (!part) return { kind: 'not_found' };

    const metadata = toAttachmentMetadata(part);
    if (part.size > maxBytes) return { kind: 'too_large', metadata };

    const attachmentResponse = await this.gmail.users.messages.attachments.get({
      userId: 'me',
      messageId: trimmedMessageId,
      id: part.attachmentId,
    });

    const raw = attachmentResponse.data.data;
    if (!raw) {
      throw new Error(
        `Gmail returned no data for attachment on message ${trimmedMessageId}.`
      );
    }

    const data = decodeBase64UrlToBuffer(raw);
    // Buffer.from(base64) silently drops invalid characters, so a truncated
    // or corrupted payload produces a short buffer with no error. Compare
    // the decoded length to Gmail's reported metadata size and fail loudly
    // on any mismatch rather than returning silently-corrupt bytes.
    if (metadata.size > 0 && data.length !== metadata.size) {
      throw new Error(
        `Attachment decode produced ${data.length} bytes on message ${trimmedMessageId}, expected ${metadata.size}. The base64 payload may be truncated or corrupted.`
      );
    }

    return { kind: 'ok', data, metadata };
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

    return {
      id: message.id ?? '',
      threadId: message.threadId ?? '',
      snippet: message.snippet ?? '',
      from: getHeaderValue(headers, 'From'),
      to: getHeaderValue(headers, 'To'),
      subject: getHeaderValue(headers, 'Subject') || '(no subject)',
      date: getHeaderValue(headers, 'Date'),
      internalDate: Number.isFinite(internalDate) ? internalDate : 0,
      body: includeBody ? extractEmailBody(message.payload) : undefined,
      labels: message.labelIds ?? [],
      attachments: extractAttachmentParts(message.payload).map(toAttachmentMetadata),
      accountId: this.account.id,
      accountEmail: this.account.email,
    };
  }
}
