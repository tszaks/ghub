#!/usr/bin/env node

import { promises as fs } from 'node:fs';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
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
  type ParsedEmail,
} from './gmail-client.js';

const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

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

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
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

  if (email.attachments.length > 0) {
    const summary = email.attachments
      .map((a) => `${a.filename} (${formatBytes(a.size)}, part=${a.partId})`)
      .join(', ');
    lines.push(`**Attachments**: ${summary}`);
  }

  if (includeBody) {
    const body = (email.body || '').trim();
    if (body) {
      const trimmed = body.length > 600 ? `${body.slice(0, 600)}...` : body;
      lines.push(`**Body**:\n${trimmed}`);
    }
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

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
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
          name: 'list_attachments',
          description:
            'List attachments on a specific Gmail message. Returns each attachment with filename, mime type, size, and stable MIME part id for use with download_attachment.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_id: { type: 'string', description: 'Gmail message id.' },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'download_attachment',
          description:
            'Download a Gmail attachment as an MCP embedded resource. The host client (Claude Code / Claude Desktop) handles saving or rendering — the server does not write to disk. Enforces a 25 MB cap.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              message_id: { type: 'string', description: 'Gmail message id.' },
              part_id: {
                type: 'string',
                description:
                  'Stable Gmail MIME part id (e.g. "0.1") as shown by list_attachments or the Attachments line in read_emails. Note: Gmail rotates the underlying attachment id per request, so the server resolves part_id to a fresh attachment id internally.',
              },
            },
            required: ['account', 'message_id', 'part_id'],
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
          case 'get_email_thread':
            return await this.handleGetThread(args);
          case 'list_attachments':
            return await this.handleListAttachments(args);
          case 'download_attachment':
            return await this.handleDownloadAttachment(args);
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

  private async handleListAttachments(rawArgs: Record<string, unknown>): Promise<CallToolResult> {
    const account_id = valueToString(rawArgs.account).trim();
    const message_id = valueToString(rawArgs.message_id).trim();

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, account_id);
    const client = await this.getClientForAccount(account);

    const attachments = await client.listAttachments(message_id);

    if (attachments.length === 0) {
      return textResult(`No attachments on message ${message_id}.`);
    }

    const lines = [
      `# Attachments on message ${message_id}`,
      '',
      `**Account**: ${account.id} (${account.email})`,
      `**Count**: ${attachments.length}`,
      '',
    ];
    attachments.forEach((att, index) => {
      lines.push(
        `#${index + 1} ${att.filename} · ${att.mimeType} · ${formatBytes(att.size)} · part=${att.partId}`
      );
    });
    return textResult(lines.join('\n'));
  }

  private async handleDownloadAttachment(
    rawArgs: Record<string, unknown>
  ): Promise<CallToolResult> {
    const account_id = valueToString(rawArgs.account).trim();
    const message_id = valueToString(rawArgs.message_id).trim();
    const part_id = valueToString(rawArgs.part_id).trim();

    const config = await this.loadConfig();
    const account = resolveWriteAccount(config, account_id);
    const client = await this.getClientForAccount(account);

    const result = await client.downloadByPartId(message_id, part_id, MAX_ATTACHMENT_BYTES);

    switch (result.kind) {
      case 'not_found':
        return textResult(
          `No attachment with part_id=${part_id} on message ${message_id} in account ${account.id}.`
        );
      case 'too_large':
        return textResult(
          [
            `Attachment ${result.metadata.filename} is ${formatBytes(result.metadata.size)}, which exceeds the ${formatBytes(MAX_ATTACHMENT_BYTES)} cap.`,
            `Gmail caps outbound attachments at 25 MB; larger payloads cannot be safely returned over MCP stdio.`,
          ].join('\n')
        );
      case 'ok': {
        const { data, metadata } = result;
        const text = [
          `Downloaded ${metadata.filename} (${formatBytes(data.length)}, ${metadata.mimeType})`,
          `Account: ${account.id} (${account.email})`,
          `Message ID: ${message_id}`,
          `Part ID: ${part_id}`,
        ].join('\n');

        return {
          content: [
            { type: 'text', text },
            {
              type: 'resource',
              resource: {
                uri: `gmail-attachment://${account.id}/${message_id}/${part_id}`,
                mimeType: metadata.mimeType,
                blob: data.toString('base64'),
              },
            },
          ],
        };
      }
      default: {
        // Exhaustiveness guard: if a new union variant is added to
        // downloadByPartId without a handler here, this line fails to
        // typecheck, turning a potential silent failure into a compile error.
        const _exhaustive: never = result;
        throw new Error(
          `Unhandled downloadByPartId result kind: ${JSON.stringify(_exhaustive)}`
        );
      }
    }
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
        '# Gmail OAuth Started',
        '',
        `**Account ID**: ${args.account_id}`,
        `**Email**: ${args.email}`,
        `**Credentials File**: ${defaultPaths.credentialsPath}`,
        '',
        'Open this URL and approve access:',
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
        '✅ Gmail OAuth completed.',
        `Account ID: ${args.account_id}`,
        `Email: ${profileEmail}`,
        `Token File: ${paths.tokenPath}`,
        '',
        'This account is now enabled for aggregate reads/search and explicit write/admin tools.',
      ].join('\n')
    );
  }

  async run(): Promise<void> {
    await ensureConfigLayout(this.configRoot);
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error(`[gmail-multi-inbox-mcp] Running on stdio. Config root: ${this.configRoot}`);
  }
}

const server = new GmailMultiInboxServer();
server.run().catch((error) => {
  console.error('[gmail-multi-inbox-mcp] Fatal error:', error);
  process.exit(1);
});
