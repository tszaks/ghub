import { OAuth2Client, type Credentials } from 'google-auth-library';
import { google, type gmail_v1 } from 'googleapis';
import { promises as fs } from 'node:fs';
import { AccountConfig, type AccountPaths, getAccountPaths } from './config.js';

export const GMAIL_SCOPES = [
  // Broadest Gmail OAuth scope Google allows (full mailbox access).
  'https://mail.google.com/',
] as const;

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

function decodeBase64Url(value: string): string {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
  return Buffer.from(padded, 'base64').toString('utf8');
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

function buildRawEmailMessage(input: {
  to: string;
  subject: string;
  body: string;
  cc?: string;
  bcc?: string;
  html?: boolean;
}): string {
  const to = normalizeOutgoingAddressList(input.to);
  if (!to) {
    throw new Error('Recipient "to" is required.');
  }

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

  return Buffer.from(lines.join('\r\n'))
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
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
          format: includeBody ? 'full' : 'metadata',
          metadataHeaders: includeBody ? undefined : ['From', 'To', 'Subject', 'Date'],
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
  }): Promise<{ draftId: string; threadId?: string }> {
    const raw = buildRawEmailMessage(input);

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

  async sendEmail(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
  }): Promise<{ messageId: string; threadId?: string }> {
    const raw = buildRawEmailMessage(input);

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
      accountId: this.account.id,
      accountEmail: this.account.email,
    };
  }
}
