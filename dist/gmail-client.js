import { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import { promises as fs, createReadStream } from 'node:fs';
import path from 'node:path';
import { getAccountPaths } from './config.js';
export const GMAIL_SCOPES = [
    'https://mail.google.com/',
    'https://www.googleapis.com/auth/gmail.settings.basic',
];
export const DRIVE_METADATA_SCOPE = 'https://www.googleapis.com/auth/drive.metadata.readonly';
export const DRIVE_FULL_SCOPE = 'https://www.googleapis.com/auth/drive';
export const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
export const DOCS_SCOPE = 'https://www.googleapis.com/auth/documents';
export const CALENDAR_SCOPE = 'https://www.googleapis.com/auth/calendar';
export const GOOGLE_ACCOUNT_SCOPES = [
    ...GMAIL_SCOPES,
    DRIVE_FULL_SCOPE,
    SHEETS_SCOPE,
    DOCS_SCOPE,
    CALENDAR_SCOPE,
];
function decodeBase64Url(value) {
    const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
    return Buffer.from(padded, 'base64').toString('utf8');
}
function decodeBase64UrlBuffer(value) {
    const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
    return Buffer.from(padded, 'base64');
}
function isInlineDisposition(headers) {
    if (!headers)
        return false;
    const disposition = headers.find((h) => h.name?.toLowerCase() === 'content-disposition');
    return Boolean(disposition?.value?.trim().toLowerCase().startsWith('inline'));
}
function extractAttachmentsMetadata(payload) {
    if (!payload)
        return [];
    const out = [];
    const consider = (part) => {
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
    const stack = payload.parts ? [...payload.parts] : [];
    while (stack.length > 0) {
        const part = stack.shift();
        if (!part)
            continue;
        consider(part);
        if (part.parts?.length)
            stack.push(...part.parts);
    }
    return out;
}
function findAttachmentPart(payload, attachmentId) {
    if (!payload)
        return null;
    if (payload.body?.attachmentId === attachmentId)
        return payload;
    const stack = payload.parts ? [...payload.parts] : [];
    while (stack.length > 0) {
        const part = stack.shift();
        if (!part)
            continue;
        if (part.body?.attachmentId === attachmentId)
            return part;
        if (part.parts?.length)
            stack.push(...part.parts);
    }
    return null;
}
function stripHtmlTags(input) {
    return input.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}
function getHeaderValue(headers, name) {
    if (!headers)
        return '';
    const found = headers.find((header) => header.name?.toLowerCase() === name.toLowerCase());
    return found?.value?.trim() ?? '';
}
function extractEmailBody(payload) {
    if (!payload)
        return '';
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
        if (!part)
            continue;
        if (part.parts?.length) {
            stack.push(...part.parts);
        }
        if (!part.body?.data)
            continue;
        if (part.mimeType === 'text/plain' && !textPlain) {
            textPlain = decodeBase64Url(part.body.data);
        }
        else if (part.mimeType === 'text/html' && !textHtml) {
            textHtml = decodeBase64Url(part.body.data);
        }
    }
    if (textPlain)
        return textPlain;
    if (textHtml)
        return stripHtmlTags(textHtml);
    return '';
}
function normalizeOutgoingAddressList(value) {
    if (!value || value.trim() === '')
        return null;
    return value
        .split(',')
        .map((item) => item.trim())
        .filter(Boolean)
        .join(', ');
}
function encodeBase64Url(value) {
    return Buffer.from(value)
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/g, '');
}
function wrapBase64(value) {
    return value.replace(/.{1,76}/g, '$&\r\n').trimEnd();
}
function inferContentType(filename) {
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
function sanitizeHeaderValue(value) {
    return value.replace(/[\r\n"]/g, ' ').trim();
}
async function buildRawEmailMessage(input) {
    const to = normalizeOutgoingAddressList(input.to);
    if (!to) {
        throw new Error('Recipient "to" is required.');
    }
    const attachments = (input.attachments ?? []).filter((attachment) => attachment.path.trim() !== '');
    if (attachments.length === 0) {
        const lines = [
            `To: ${to}`,
            `Subject: ${input.subject}`,
            'MIME-Version: 1.0',
            `Content-Type: text/${input.html ? 'html' : 'plain'}; charset=utf-8`,
        ];
        const cc = normalizeOutgoingAddressList(input.cc);
        if (cc)
            lines.push(`Cc: ${cc}`);
        const bcc = normalizeOutgoingAddressList(input.bcc);
        if (bcc)
            lines.push(`Bcc: ${bcc}`);
        if (input.inReplyTo)
            lines.push(`In-Reply-To: ${input.inReplyTo}`);
        if (input.references)
            lines.push(`References: ${input.references}`);
        lines.push('', input.body);
        return encodeBase64Url(lines.join('\r\n'));
    }
    const lines = [
        `To: ${to}`,
        `Subject: ${input.subject}`,
        'MIME-Version: 1.0',
    ];
    const cc = normalizeOutgoingAddressList(input.cc);
    if (cc)
        lines.push(`Cc: ${cc}`);
    const bcc = normalizeOutgoingAddressList(input.bcc);
    if (bcc)
        lines.push(`Bcc: ${bcc}`);
    const boundary = `gmail-multi-inbox-mcp-${Date.now().toString(36)}-${Math.random()
        .toString(36)
        .slice(2, 10)}`;
    lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`, '');
    lines.push(`--${boundary}`, `Content-Type: text/${input.html ? 'html' : 'plain'}; charset=utf-8`, 'Content-Transfer-Encoding: base64', '', wrapBase64(Buffer.from(input.body, 'utf8').toString('base64')));
    for (const attachment of attachments) {
        const filePath = attachment.path.trim();
        const fileBuffer = await fs.readFile(filePath);
        const filename = sanitizeHeaderValue(attachment.filename?.trim() || path.basename(filePath));
        const contentType = attachment.contentType?.trim() || inferContentType(filename);
        lines.push(`--${boundary}`, `Content-Type: ${contentType}; name="${filename}"`, 'Content-Transfer-Encoding: base64', `Content-Disposition: attachment; filename="${filename}"`, '', wrapBase64(fileBuffer.toString('base64')));
    }
    lines.push(`--${boundary}--`);
    return encodeBase64Url(lines.join('\r\n'));
}
function normalizeAttachments(attachments) {
    return (attachments ?? [])
        .map((attachment) => ({
        path: attachment.path.trim(),
        filename: attachment.filename?.trim() || undefined,
        contentType: attachment.contentType?.trim() || undefined,
    }))
        .filter((attachment) => attachment.path !== '');
}
async function createRawEmailMessage(input) {
    return buildRawEmailMessage({
        ...input,
        attachments: normalizeAttachments(input.attachments),
    });
}
export function createOAuthClientFromCredentials(options) {
    if (!options.credentials || typeof options.credentials !== 'object') {
        throw new Error('Invalid credentials content.');
    }
    const credentialsObject = options.credentials;
    const source = credentialsObject.installed ?? credentialsObject.web;
    if (!source?.client_id || !source.client_secret) {
        throw new Error('Credentials must include client_id and client_secret under "installed" or "web".');
    }
    const redirectUri = source.redirect_uris?.[0] ?? 'http://localhost';
    return new OAuth2Client(source.client_id, source.client_secret, redirectUri);
}
export async function readCredentialsFile(credentialsPath) {
    const raw = await fs.readFile(credentialsPath, 'utf8');
    return JSON.parse(raw);
}
export async function buildOAuthClientFromCredentialsFile(credentialsPath) {
    const credentials = await readCredentialsFile(credentialsPath);
    return createOAuthClientFromCredentials({ credentials });
}
export function generateAuthUrlFromCredentials(credentials) {
    const oauth2Client = createOAuthClientFromCredentials({ credentials });
    const authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: [...GOOGLE_ACCOUNT_SCOPES],
        prompt: 'consent',
    });
    return { oauth2Client, authUrl };
}
export async function exchangeCodeForToken(credentials, authorizationCode) {
    const oauth2Client = createOAuthClientFromCredentials({ credentials });
    const { tokens } = await oauth2Client.getToken(authorizationCode);
    return tokens;
}
function sanitizeMessageIds(messageIds) {
    return Array.from(new Set(messageIds
        .map((messageId) => messageId.trim())
        .filter((messageId) => messageId.length > 0)));
}
function escapeDriveQueryValue(value) {
    return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}
export function buildDriveSearchQuery(query) {
    const normalizedQuery = query.trim();
    if (!normalizedQuery) {
        throw new Error('Drive search query is required.');
    }
    const escapedQuery = escapeDriveQueryValue(normalizedQuery);
    return `trashed = false and (name contains '${escapedQuery}' or fullText contains '${escapedQuery}')`;
}
export function describeDriveApiError(error, fallback) {
    const googleError = error;
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
    if (googleError.code === 403 &&
        /insufficient.*scope|insufficient.*permission/i.test(message)) {
        return [
            'Drive access is not granted for this account yet.',
            'Re-run `begin_account_auth` and `finish_account_auth` to grant Google Drive metadata access.',
        ].join(' ');
    }
    return message;
}
const WORKSPACE_EXPORT_MAP = {
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
    account;
    paths;
    gmail;
    drive;
    sheets;
    docs;
    calendar;
    constructor(account, paths, gmail, drive, sheets, docs, calendar) {
        this.account = account;
        this.paths = paths;
        this.gmail = gmail;
        this.drive = drive;
        this.sheets = sheets;
        this.docs = docs;
        this.calendar = calendar;
    }
    static async create(configRoot, account) {
        const paths = getAccountPaths(configRoot, account);
        const oauth2Client = await buildOAuthClientFromCredentialsFile(paths.credentialsPath);
        let cachedTokens;
        try {
            const rawToken = await fs.readFile(paths.tokenPath, 'utf8');
            cachedTokens = JSON.parse(rawToken);
        }
        catch (error) {
            throw new Error(`Token file missing or invalid for account "${account.id}" at ${paths.tokenPath}: ${error.message}`);
        }
        oauth2Client.setCredentials(cachedTokens);
        oauth2Client.on('tokens', (incomingTokens) => {
            cachedTokens = { ...cachedTokens, ...incomingTokens };
            void fs
                .writeFile(paths.tokenPath, `${JSON.stringify(cachedTokens, null, 2)}\n`, 'utf8')
                .catch((error) => {
                console.error(`[ghub] Failed to persist refreshed token for account ${account.id}:`, error);
            });
        });
        const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
        const drive = google.drive({ version: 'v3', auth: oauth2Client });
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
        const docs = google.docs({ version: 'v1', auth: oauth2Client });
        const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
        return new GmailAccountClient(account, paths, gmail, drive, sheets, docs, calendar);
    }
    async getProfileEmail() {
        const profile = await this.gmail.users.getProfile({ userId: 'me' });
        if (!profile.data.emailAddress) {
            throw new Error(`Gmail profile did not return an email address for account "${this.account.id}".`);
        }
        return profile.data.emailAddress;
    }
    async readEmails(query, maxResults, includeBody) {
        return this.fetchMessages(query, maxResults, includeBody);
    }
    async searchEmails(query, maxResults) {
        if (!query || query.trim() === '') {
            throw new Error('Search query is required.');
        }
        return this.fetchMessages(query, maxResults, false);
    }
    async searchDriveFiles(query, maxResults) {
        const boundedMax = Math.max(1, Math.min(maxResults, 500));
        try {
            const response = await this.drive.files.list({
                q: buildDriveSearchQuery(query),
                pageSize: boundedMax,
                includeItemsFromAllDrives: true,
                supportsAllDrives: true,
                orderBy: 'modifiedTime desc',
                fields: 'files(id,name,mimeType,modifiedTime,webViewLink,owners(displayName,emailAddress))',
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
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive search failed.'));
        }
    }
    async fetchMessages(query, maxResults, includeBody) {
        const boundedMax = Math.max(1, Math.min(maxResults, 500));
        const listResponse = await this.gmail.users.messages.list({
            userId: 'me',
            q: query.trim() === '' ? undefined : query,
            maxResults: boundedMax,
        });
        const messageIds = (listResponse.data.messages ?? [])
            .map((message) => message.id)
            .filter((id) => Boolean(id));
        if (messageIds.length === 0) {
            return [];
        }
        const fullMessages = await Promise.all(messageIds.map((messageId) => this.gmail.users.messages.get({
            userId: 'me',
            id: messageId,
            format: 'full',
        })));
        return fullMessages
            .map((response) => this.parseMessage(response.data, includeBody))
            .sort((a, b) => b.internalDate - a.internalDate);
    }
    async listAttachments(messageId) {
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
    async getAttachment(messageId, attachmentId) {
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
            throw new Error(`Attachment ${attachmentId} not found on message ${messageId}.`);
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
    async getThread(threadId) {
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
    async getLabels() {
        const labelsResponse = await this.gmail.users.labels.list({ userId: 'me' });
        return (labelsResponse.data.labels ?? []).map((label) => ({
            id: label.id ?? '',
            name: label.name ?? '(unnamed)',
            type: label.type ?? undefined,
            messagesTotal: label.messagesTotal ?? undefined,
        }));
    }
    async markAsRead(messageIds) {
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
    async addLabels(messageIds, labelIds) {
        const ids = sanitizeMessageIds(messageIds);
        const labels = labelIds.map((labelId) => labelId.trim()).filter(Boolean);
        if (ids.length === 0)
            throw new Error('message_ids must include at least one value.');
        if (labels.length === 0)
            throw new Error('label_ids must include at least one value.');
        await this.gmail.users.messages.batchModify({
            userId: 'me',
            requestBody: {
                ids,
                addLabelIds: labels,
            },
        });
        return ids.length;
    }
    async removeLabels(messageIds, labelIds) {
        const ids = sanitizeMessageIds(messageIds);
        const labels = labelIds.map((labelId) => labelId.trim()).filter(Boolean);
        if (ids.length === 0)
            throw new Error('message_ids must include at least one value.');
        if (labels.length === 0)
            throw new Error('label_ids must include at least one value.');
        await this.gmail.users.messages.batchModify({
            userId: 'me',
            requestBody: {
                ids,
                removeLabelIds: labels,
            },
        });
        return ids.length;
    }
    async archiveEmails(messageIds) {
        const ids = sanitizeMessageIds(messageIds);
        if (ids.length === 0)
            throw new Error('message_ids must include at least one value.');
        await this.gmail.users.messages.batchModify({
            userId: 'me',
            requestBody: {
                ids,
                removeLabelIds: ['INBOX'],
            },
        });
        return ids.length;
    }
    async trashEmails(messageIds) {
        const ids = sanitizeMessageIds(messageIds);
        if (ids.length === 0)
            throw new Error('message_ids must include at least one value.');
        await Promise.all(ids.map((messageId) => this.gmail.users.messages.trash({
            userId: 'me',
            id: messageId,
        })));
        return ids.length;
    }
    async createLabel(name, labelListVisibility = 'labelShow', messageListVisibility = 'show') {
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
    async deleteLabel(labelId) {
        if (!labelId || labelId.trim() === '') {
            throw new Error('label_id is required.');
        }
        await this.gmail.users.labels.delete({
            userId: 'me',
            id: labelId,
        });
    }
    async createFilter(criteria, action) {
        const response = await this.gmail.users.settings.filters.create({
            userId: 'me',
            requestBody: { criteria, action },
        });
        return response.data;
    }
    async listFilters() {
        const response = await this.gmail.users.settings.filters.list({ userId: 'me' });
        return response.data.filter ?? [];
    }
    async deleteFilter(filterId) {
        if (!filterId || filterId.trim() === '') {
            throw new Error('filter_id is required.');
        }
        await this.gmail.users.settings.filters.delete({
            userId: 'me',
            id: filterId.trim(),
        });
    }
    async createBlockFilter(sender, action) {
        const trimmed = sender.trim();
        if (!trimmed)
            throw new Error('sender is required.');
        const criteria = { from: trimmed };
        const filterAction = action === 'archive'
            ? { removeLabelIds: ['INBOX'] }
            : action === 'spam'
                ? { addLabelIds: ['SPAM'], removeLabelIds: ['INBOX'] }
                : { addLabelIds: ['TRASH'], removeLabelIds: ['INBOX', 'UNREAD'] };
        return this.createFilter(criteria, filterAction);
    }
    async modifyThread(threadId, modifications) {
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
    async getThreadSubject(threadId) {
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
    async getMessageHeaders(messageId, headerNames) {
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
        const result = {};
        for (const name of names) {
            result[name] = getHeaderValue(headers, name);
        }
        return result;
    }
    async createDraft(input) {
        const raw = await createRawEmailMessage(input);
        const message = { raw };
        if (input.threadId)
            message.threadId = input.threadId;
        const response = await this.gmail.users.drafts.create({
            userId: 'me',
            requestBody: { message },
        });
        return {
            draftId: response.data.id ?? '',
            threadId: response.data.message?.threadId ?? undefined,
        };
    }
    async deleteDrafts(draftIds) {
        const ids = draftIds.map((id) => id.trim()).filter(Boolean);
        if (ids.length === 0)
            throw new Error('draft_ids must include at least one value.');
        await Promise.all(ids.map((draftId) => this.gmail.users.drafts.delete({
            userId: 'me',
            id: draftId,
        })));
        return ids.length;
    }
    async sendDraft(draftId) {
        const response = await this.gmail.users.drafts.send({
            userId: 'me',
            requestBody: { id: draftId },
        });
        return {
            messageId: response.data.id ?? '',
            threadId: response.data.threadId ?? undefined,
        };
    }
    async listDrafts(maxResults = 20) {
        const listRes = await this.gmail.users.drafts.list({
            userId: 'me',
            maxResults,
        });
        const drafts = listRes.data.drafts ?? [];
        if (drafts.length === 0)
            return [];
        const details = await Promise.all(drafts.map((d) => this.gmail.users.drafts.get({
            userId: 'me',
            id: d.id,
            format: 'metadata',
        })));
        const results = details.map((res) => {
            const headers = res.data.message?.payload?.headers ?? [];
            const get = (name) => headers.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value ?? '';
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
    async searchDrafts(query, maxResults = 20) {
        const listRes = await this.gmail.users.drafts.list({
            userId: 'me',
            maxResults,
            q: query,
        });
        const drafts = listRes.data.drafts ?? [];
        if (drafts.length === 0)
            return [];
        const details = await Promise.all(drafts.map((d) => this.gmail.users.drafts.get({
            userId: 'me',
            id: d.id,
            format: 'metadata',
        })));
        return details.map((res) => {
            const headers = res.data.message?.payload?.headers ?? [];
            const get = (name) => headers.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value ?? '';
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
    async sendEmail(input) {
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
    parseMessage(message, includeBody) {
        const headers = message.payload?.headers;
        const internalDate = Number(message.internalDate ?? 0);
        const attachments = extractAttachmentsMetadata(message.payload).filter((a) => !a.isInline);
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
    async listDriveFiles(options) {
        const parts = ['trashed = false'];
        if (options.folderId?.trim())
            parts.push(`'${options.folderId.trim()}' in parents`);
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
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive list failed.'));
        }
    }
    async getDriveFile(fileId) {
        if (!fileId?.trim())
            throw new Error('file_id is required.');
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
                exportLinks: f.exportLinks ? f.exportLinks : undefined,
                accountId: this.account.id,
                accountEmail: this.account.email,
            };
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive file get failed.'));
        }
    }
    async getDriveFileContent(fileId) {
        if (!fileId?.trim())
            throw new Error('file_id is required.');
        const meta = await this.getDriveFile(fileId.trim());
        const wsExport = WORKSPACE_EXPORT_MAP[meta.mimeType];
        try {
            if (wsExport) {
                const response = await this.drive.files.export({ fileId: fileId.trim(), mimeType: wsExport.exportMime }, { responseType: 'arraybuffer' });
                return {
                    bytes: Buffer.from(response.data),
                    contentType: wsExport.contentType,
                    filename: `${meta.name}${wsExport.ext}`,
                };
            }
            const response = await this.drive.files.get({ fileId: fileId.trim(), alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
            return {
                bytes: Buffer.from(response.data),
                contentType: meta.mimeType,
                filename: meta.name,
            };
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive file download failed.'));
        }
    }
    async uploadDriveFile(input) {
        const localPath = input.localPath.trim();
        if (!localPath)
            throw new Error('local_path is required.');
        const filename = input.name?.trim() || path.basename(localPath);
        const mimeType = input.mimeType?.trim() || inferContentType(filename);
        const requestBody = { name: filename };
        if (input.folderId?.trim())
            requestBody.parents = [input.folderId.trim()];
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
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive upload failed.'));
        }
    }
    async createDriveFolder(name, parentId) {
        if (!name?.trim())
            throw new Error('name is required.');
        const requestBody = {
            name: name.trim(),
            mimeType: 'application/vnd.google-apps.folder',
        };
        if (parentId?.trim())
            requestBody.parents = [parentId.trim()];
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
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Create folder failed.'));
        }
    }
    async updateDriveFile(fileId, updates) {
        if (!fileId?.trim())
            throw new Error('file_id is required.');
        const requestBody = {};
        if (updates.name !== undefined)
            requestBody.name = updates.name.trim();
        if (updates.starred !== undefined)
            requestBody.starred = updates.starred;
        if (updates.description !== undefined)
            requestBody.description = updates.description;
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
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive file update failed.'));
        }
    }
    async trashDriveFile(fileId) {
        if (!fileId?.trim())
            throw new Error('file_id is required.');
        try {
            await this.drive.files.update({
                fileId: fileId.trim(),
                supportsAllDrives: true,
                requestBody: { trashed: true },
            });
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive trash failed.'));
        }
    }
    async shareDriveFile(fileId, input) {
        if (!fileId?.trim())
            throw new Error('file_id is required.');
        try {
            const response = await this.drive.permissions.create({
                fileId: fileId.trim(),
                supportsAllDrives: true,
                sendNotificationEmail: input.sendNotification ?? true,
                emailMessage: input.notificationMessage,
                requestBody: { type: input.type, role: input.role, emailAddress: input.email },
            });
            return { permissionId: response.data.id ?? '' };
        }
        catch (error) {
            throw new Error(describeDriveApiError(error, 'Drive share failed.'));
        }
    }
    // ─── Sheets ──────────────────────────────────────────────────────────────
    async getSheetsMetadata(spreadsheetId) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
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
    async readSheetValues(spreadsheetId, range, valueRenderOption = 'FORMATTED_VALUE') {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        if (!range?.trim())
            throw new Error('range is required.');
        const response = await this.sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId.trim(),
            range: range.trim(),
            valueRenderOption,
        });
        return {
            range: response.data.range ?? range,
            values: (response.data.values ?? []),
        };
    }
    async writeSheetValues(spreadsheetId, range, values, valueInputOption = 'USER_ENTERED') {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        if (!range?.trim())
            throw new Error('range is required.');
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
    async appendSheetValues(spreadsheetId, range, values, valueInputOption = 'USER_ENTERED') {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        if (!range?.trim())
            throw new Error('range is required.');
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
    async createSpreadsheet(title) {
        if (!title?.trim())
            throw new Error('title is required.');
        const response = await this.sheets.spreadsheets.create({
            requestBody: { properties: { title: title.trim() } },
        });
        return {
            id: response.data.spreadsheetId ?? '',
            url: response.data.spreadsheetUrl ?? `https://docs.google.com/spreadsheets/d/${response.data.spreadsheetId}`,
        };
    }
    async getSheetIdByTitle(spreadsheetId, sheetTitle) {
        const meta = await this.getSheetsMetadata(spreadsheetId);
        const sheet = meta.sheets.find((s) => s.title === sheetTitle);
        if (!sheet)
            throw new Error(`Sheet tab "${sheetTitle}" not found in spreadsheet.`);
        return sheet.sheetId;
    }
    async sheetsBatchUpdate(spreadsheetId, requests) {
        const response = await this.sheets.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId.trim(),
            requestBody: { requests },
        });
        return response.data;
    }
    async addSheetTab(spreadsheetId, title, index) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        if (!title?.trim())
            throw new Error('title is required.');
        const props = { title: title.trim() };
        if (index !== undefined)
            props.index = index;
        const result = await this.sheetsBatchUpdate(spreadsheetId, [{ addSheet: { properties: props } }]);
        const added = result.replies?.[0]?.addSheet?.properties;
        return { sheetId: added?.sheetId ?? 0, title: added?.title ?? title };
    }
    async renameSheetTab(spreadsheetId, currentTitle, newTitle) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, currentTitle);
        await this.sheetsBatchUpdate(spreadsheetId, [{
                updateSheetProperties: {
                    properties: { sheetId, title: newTitle.trim() },
                    fields: 'title',
                },
            }]);
    }
    async deleteSheetTab(spreadsheetId, sheetTitle) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
        await this.sheetsBatchUpdate(spreadsheetId, [{ deleteSheet: { sheetId } }]);
    }
    async insertDimension(spreadsheetId, sheetTitle, dimension, startIndex, count) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
        await this.sheetsBatchUpdate(spreadsheetId, [{
                insertDimension: {
                    range: { sheetId, dimension, startIndex, endIndex: startIndex + count },
                    inheritFromBefore: startIndex > 0,
                },
            }]);
    }
    async deleteDimension(spreadsheetId, sheetTitle, dimension, startIndex, count) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
        await this.sheetsBatchUpdate(spreadsheetId, [{
                deleteDimension: {
                    range: { sheetId, dimension, startIndex, endIndex: startIndex + count },
                },
            }]);
    }
    parseA1Range(range, sheetId) {
        const colToIndex = (letters) => {
            let index = 0;
            for (const ch of letters.toUpperCase())
                index = index * 26 + (ch.charCodeAt(0) - 64);
            return index - 1;
        };
        const withoutSheet = range.includes('!') ? range.split('!')[1] ?? range : range;
        const [startCell, endCell] = withoutSheet.split(':');
        const parseCell = (cell) => {
            const m = (cell ?? '').match(/^([A-Za-z]+)(\d+)$/);
            if (!m)
                return { row: 0, col: 0 };
            return { col: colToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
        };
        const s = parseCell(startCell ?? '');
        const e = endCell ? parseCell(endCell) : s;
        return { sheetId, startRowIndex: s.row, endRowIndex: e.row + 1, startColumnIndex: s.col, endColumnIndex: e.col + 1 };
    }
    hexToColor(hex) {
        const h = hex.replace('#', '');
        return {
            red: parseInt(h.substring(0, 2), 16) / 255,
            green: parseInt(h.substring(2, 4), 16) / 255,
            blue: parseInt(h.substring(4, 6), 16) / 255,
        };
    }
    async formatCells(spreadsheetId, sheetTitle, range, format) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
        const gridRange = this.parseA1Range(range, sheetId);
        const cellFormat = {};
        const fields = [];
        if (format.bold !== undefined || format.italic !== undefined || format.fontSize !== undefined || format.textColor !== undefined) {
            cellFormat.textFormat = {};
            if (format.bold !== undefined) {
                cellFormat.textFormat.bold = format.bold;
                fields.push('userEnteredFormat.textFormat.bold');
            }
            if (format.italic !== undefined) {
                cellFormat.textFormat.italic = format.italic;
                fields.push('userEnteredFormat.textFormat.italic');
            }
            if (format.fontSize !== undefined) {
                cellFormat.textFormat.fontSize = format.fontSize;
                fields.push('userEnteredFormat.textFormat.fontSize');
            }
            if (format.textColor) {
                cellFormat.textFormat.foregroundColor = this.hexToColor(format.textColor);
                fields.push('userEnteredFormat.textFormat.foregroundColor');
            }
        }
        if (format.backgroundColor) {
            cellFormat.backgroundColor = this.hexToColor(format.backgroundColor);
            fields.push('userEnteredFormat.backgroundColor');
        }
        if (format.horizontalAlignment) {
            cellFormat.horizontalAlignment = format.horizontalAlignment;
            fields.push('userEnteredFormat.horizontalAlignment');
        }
        if (format.numberFormat) {
            cellFormat.numberFormat = { type: 'NUMBER', pattern: format.numberFormat };
            fields.push('userEnteredFormat.numberFormat');
        }
        if (format.wrapStrategy) {
            cellFormat.wrapStrategy = format.wrapStrategy;
            fields.push('userEnteredFormat.wrapStrategy');
        }
        if (fields.length === 0)
            return;
        await this.sheetsBatchUpdate(spreadsheetId, [{
                repeatCell: { range: gridRange, cell: { userEnteredFormat: cellFormat }, fields: fields.join(',') },
            }]);
    }
    async addChart(spreadsheetId, sheetTitle, options) {
        if (!spreadsheetId?.trim())
            throw new Error('spreadsheet_id is required.');
        const sheetId = await this.getSheetIdByTitle(spreadsheetId, sheetTitle);
        const dataRange = this.parseA1Range(options.dataRange, sheetId);
        const domainRange = { ...dataRange, endColumnIndex: dataRange.startColumnIndex + 1 };
        const seriesRange = { ...dataRange, startColumnIndex: dataRange.startColumnIndex + 1 };
        let spec;
        if (options.chartType === 'PIE') {
            spec = {
                title: options.title,
                pieChart: {
                    legendPosition: 'RIGHT_LEGEND',
                    domain: { sourceRange: { sources: [domainRange] } },
                    series: { sourceRange: { sources: [seriesRange] } },
                },
            };
        }
        else {
            spec = {
                title: options.title,
                basicChart: {
                    chartType: options.chartType,
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
    async getDocument(documentId) {
        if (!documentId?.trim())
            throw new Error('document_id is required.');
        const response = await this.docs.documents.get({ documentId: documentId.trim() });
        const doc = response.data;
        const parts = [];
        const extract = (elements) => {
            for (const el of elements) {
                if (el.paragraph) {
                    for (const pe of el.paragraph.elements ?? []) {
                        if (pe.textRun?.content)
                            parts.push(pe.textRun.content);
                    }
                }
                else if (el.table) {
                    for (const row of el.table.tableRows ?? []) {
                        for (const cell of row.tableCells ?? [])
                            extract(cell.content ?? []);
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
    async createDocument(title) {
        if (!title?.trim())
            throw new Error('title is required.');
        const response = await this.docs.documents.create({ requestBody: { title: title.trim() } });
        return {
            documentId: response.data.documentId ?? '',
            url: `https://docs.google.com/document/d/${response.data.documentId}/edit`,
        };
    }
    async appendToDocument(documentId, text, options) {
        if (!documentId?.trim())
            throw new Error('document_id is required.');
        const doc = await this.docs.documents.get({ documentId: documentId.trim() });
        const endIndex = doc.data.body?.content?.slice(-1)?.[0]?.endIndex ?? 1;
        const insertIndex = endIndex - 1;
        const insertedText = text.endsWith('\n') ? text : text + '\n';
        const requests = [
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
        const styleFields = [];
        const textStyle = {};
        if (options?.bold !== undefined) {
            textStyle.bold = options.bold;
            styleFields.push('bold');
        }
        if (options?.italic !== undefined) {
            textStyle.italic = options.italic;
            styleFields.push('italic');
        }
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
    async replaceInDocument(documentId, findText, replaceText, matchCase = false) {
        if (!documentId?.trim())
            throw new Error('document_id is required.');
        const response = await this.docs.documents.batchUpdate({
            documentId: documentId.trim(),
            requestBody: {
                requests: [{ replaceAllText: { containsText: { text: findText, matchCase }, replaceText } }],
            },
        });
        return { occurrencesChanged: response.data.replies?.[0]?.replaceAllText?.occurrencesChanged ?? 0 };
    }
    async insertTableInDocument(documentId, rows, columns) {
        if (!documentId?.trim())
            throw new Error('document_id is required.');
        const doc = await this.docs.documents.get({ documentId: documentId.trim() });
        const endIndex = doc.data.body?.content?.slice(-1)?.[0]?.endIndex ?? 1;
        await this.docs.documents.batchUpdate({
            documentId: documentId.trim(),
            requestBody: {
                requests: [{ insertTable: { rows, columns, location: { index: endIndex - 1 } } }],
            },
        });
    }
    async applyDocHeadingStyle(documentId, startIndex, endIndex, style) {
        if (!documentId?.trim())
            throw new Error('document_id is required.');
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
    async listCalendars() {
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
    async listCalendarEvents(calendarId, options) {
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
    async getCalendarEvent(calendarId, eventId) {
        if (!calendarId?.trim())
            throw new Error('calendar_id is required.');
        if (!eventId?.trim())
            throw new Error('event_id is required.');
        const response = await this.calendar.events.get({ calendarId: calendarId.trim(), eventId: eventId.trim() });
        return this.parseCalendarEvent(response.data, calendarId);
    }
    async createCalendarEvent(calendarId, input) {
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
    async updateCalendarEvent(calendarId, eventId, updates) {
        if (!calendarId?.trim())
            throw new Error('calendar_id is required.');
        if (!eventId?.trim())
            throw new Error('event_id is required.');
        const requestBody = {};
        if (updates.summary !== undefined)
            requestBody.summary = updates.summary;
        if (updates.description !== undefined)
            requestBody.description = updates.description;
        if (updates.location !== undefined)
            requestBody.location = updates.location;
        if (updates.start !== undefined)
            requestBody.start = updates.start;
        if (updates.end !== undefined)
            requestBody.end = updates.end;
        if (updates.status !== undefined)
            requestBody.status = updates.status;
        if (updates.attendees !== undefined)
            requestBody.attendees = updates.attendees.map((email) => ({ email }));
        const response = await this.calendar.events.patch({
            calendarId: calendarId.trim(),
            eventId: eventId.trim(),
            sendNotifications: updates.sendNotifications ?? true,
            requestBody,
        });
        return this.parseCalendarEvent(response.data, calendarId);
    }
    async deleteCalendarEvent(calendarId, eventId, sendNotifications = true) {
        if (!calendarId?.trim())
            throw new Error('calendar_id is required.');
        if (!eventId?.trim())
            throw new Error('event_id is required.');
        await this.calendar.events.delete({ calendarId: calendarId.trim(), eventId: eventId.trim(), sendNotifications });
    }
    parseCalendarEvent(event, calendarId) {
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
//# sourceMappingURL=gmail-client.js.map