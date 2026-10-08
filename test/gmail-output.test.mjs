import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ts from 'typescript';
import { GmailMultiInboxServer } from '../dist/app.js';

const primary = { id: 'work', email: 'work@example.test', enabled: true };
const secondary = { id: 'personal', email: 'personal@example.test', enabled: true };
const secret = 'access_token=DO-NOT-EXPOSE';
const email = (id, internalDate, account = primary) => ({
  id, threadId: 'thread-1', snippet: 'A short preview', from: 'sender@example.test',
  to: account.email, cc: '', subject: 'Project update', date: 'Mon, 01 Jan 2024 00:00:00 +0000',
  internalDate, messageHeaderId: '<mail@example.test>', inReplyTo: '', references: '',
  labels: ['INBOX'], attachments: [], accountId: account.id, accountEmail: account.email,
});
const draft = { draftId: 'draft-1', messageId: 'message-1', threadId: 'thread-1', subject: 'Draft', to: 'to@example.test', internalDate: 123, snippet: 'Draft text' };
const label = { id: 'label-1', name: 'Review', type: 'user', messagesTotal: 2 };
const filter = { id: 'filter-1', criteria: { from: 'sender@example.test' }, action: { addLabelIds: ['TRASH'] } };

function mockApp(methods = {}, accounts = [primary], forAccount) {
  const app = new GmailMultiInboxServer('/tmp/ghub-tests-never-read-accounts');
  app.loadConfig = async () => ({ defaultAccount: null, accounts });
  app.getClientForAccount = async (account) => new Proxy(forAccount ? forAccount(account) : methods, {
    get(target, prop) {
      if (prop === 'then') return undefined;
      if (prop in target) return target[prop];
      throw new Error(`Unexpected client method: ${String(prop)}`);
    },
  });
  return app;
}

function structured(result) {
  assert.ok(result.structuredContent && typeof result.structuredContent === 'object');
  assert.equal(result.content.length, 1);
  assert.equal(result.content[0].type, 'text');
  assert.equal(typeof result.content[0].text, 'string');
  return result.structuredContent;
}

for (const tool of ['read_emails', 'search_emails']) {
  test(`${tool} returns structured native records and an explicit-account continuation`, async () => {
    const calls = [];
    const item = email('message-1', 100);
    const app = mockApp({ readEmailPage: async (...args) => { calls.push(args); return { emails: [item], nextPageToken: 'page-2', resultSizeEstimate: 15 }; } });
    const result = await app.callTool(tool, { account: 'work', query: 'label:inbox', include_body: true, max_results: 2, page_token: 'page-1' });
    const data = structured(result);
    assert.deepEqual(data.emails, [item]);
    assert.deepEqual(calls, [['label:inbox', 2, tool === 'read_emails', 'page-1']]);
    assert.equal(data.account, 'work');
    assert.equal(data.count, 1);
    assert.equal(data.nextPageToken, 'page-2');
    assert.equal(data.resultSizeEstimate, 15);
    assert.deepEqual(data.errors, []);
    assert.deepEqual(data.successfulAccounts, ['work']);
    assert.match(result.content[0].text, /Gmail ID.*message-1/);
  });

  test(`${tool} aggregates accounts without lossy cross-account continuation tokens`, async () => {
    const app = mockApp({}, [primary, secondary], (account) => ({ readEmailPage: async () => ({ emails: [email(account.id + '-old', 10, account), email(account.id + '-new', account.id === 'work' ? 30 : 20, account)], nextPageToken: account.id + '-next', resultSizeEstimate: 50 }) }));
    const data = structured(await app.callTool(tool, { query: 'newer:2020/01/01', max_results: 2 }));
    assert.deepEqual(data.emails.map((item) => item.id), ['work-new', 'personal-new']);
    assert.equal(data.count, 2);
    assert.equal(data.totalFound, 4);
    assert.equal('nextPageToken' in data, false);
    assert.equal('nextPageTokens' in data, false);
    assert.deepEqual(data.pagination, { mode: 'per_account', recovery: 'Repeat with an explicit account to page all results.' });
    assert.ok(!JSON.stringify(data).includes('-next'));
  });

  test(`${tool} has structured empty and sanitized partial-account results`, async () => {
    const app = mockApp({}, [primary, secondary], (account) => ({ readEmailPage: async () => {
      if (account.id === 'personal') throw new Error(secret);
      return { emails: [], resultSizeEstimate: 0 };
    } }));
    const data = structured(await app.callTool(tool, { query: 'no-matches' }));
    assert.deepEqual(data.emails, []);
    assert.equal(data.count, 0);
    assert.deepEqual(data.successfulAccounts, ['work'], 'an empty successful account is still a success');
    assert.deepEqual(data.errors, [{
      account: 'personal', code: 'INTERNAL_ERROR', message: 'The operation failed unexpectedly.',
      exitCode: 1, recovery: 'Check configuration and report the command name and error code without credentials.', retryable: false,
    }]);
    assert.ok(!JSON.stringify(data).includes(secret));
  });

  test(`${tool} preserves sanitized authentication and rate-limit classifications`, async (t) => {
    const cases = [
      ['missing token file', new Error(`Token file missing: ${secret}`), 'AUTH_REQUIRED', 3, false, true],
      ['HTTP 401', Object.assign(new Error(secret), { response: { status: 401, data: { secret }, config: { headers: { Authorization: secret } } } }), 'AUTH_REQUIRED', 3, false, false],
      ['HTTP 429', Object.assign(new Error(secret), { response: { status: 429, data: { secret } } }), 'UPSTREAM_UNAVAILABLE', 5, true, false],
    ];
    for (const [name, failure, code, exitCode, retryable, failClientCreation] of cases) {
      await t.test(name, async () => {
        const app = mockApp({ readEmailPage: async () => { throw failure; } });
        if (failClientCreation) app.getClientForAccount = async () => { throw failure; };
        const data = structured(await app.callTool(tool, { account: 'work', query: 'query' }));
        assert.deepEqual(data.successfulAccounts, []);
        assert.deepEqual(data.emails, []);
        assert.equal(data.errors.length, 1);
        assert.equal(data.errors[0].account, 'work');
        assert.equal(data.errors[0].code, code);
        assert.equal(data.errors[0].exitCode, exitCode);
        assert.equal(data.errors[0].retryable, retryable);
        assert.equal(typeof data.errors[0].recovery, 'string');
        assert.ok(!JSON.stringify(data).includes(secret));
      });
    }
  });

  test(`${tool} identifies all-account failure independently of the email count`, async () => {
    const app = mockApp({ readEmailPage: async () => { throw Object.assign(new Error(secret), { response: { status: 401 } }); } }, [primary, secondary]);
    const data = structured(await app.callTool(tool, { query: 'query' }));
    assert.deepEqual(data.successfulAccounts, []);
    assert.deepEqual(data.errors.map((error) => error.account), ['work', 'personal']);
    assert.equal(data.errors.every((error) => error.code === 'AUTH_REQUIRED' && error.exitCode === 3), true);
    assert.ok(!JSON.stringify(data).includes(secret));
  });

  test(`${tool} rejects a page token without account before contacting any client`, async () => {
    await assert.rejects(mockApp().callTool(tool, { query: 'query', page_token: 'page-2' }), /page_token requires an explicit account/);
    const schema = mockApp().listTools().find((item) => item.name === tool).inputSchema;
    assert.equal(schema.properties.page_token.type, 'string');
  });
}

test('Gmail record and mutation tools retain machine-readable IDs and actual succeeded counts', async (t) => {
  const message = email('message-1', 100);
  const cases = [
    ['get_email_thread', { thread_id: 'thread-1' }, { getThread: async () => ({ threadId: 'thread-1', messages: [message] }) }, { threadId: 'thread-1', messages: [message], count: 1 }],
    ['get_labels', {}, { getLabels: async () => [label] }, { labels: [label], count: 1 }],
    ['mark_as_read', { message_ids: ['m1', 'm2'] }, { markAsRead: async () => 2 }, { messageIds: ['m1', 'm2'], succeededCount: 2 }],
    ['add_labels', { message_ids: ['m1'], label_ids: ['label-1'] }, { addLabels: async () => 1 }, { messageIds: ['m1'], labelIds: ['label-1'], succeededCount: 1 }],
    ['remove_labels', { message_ids: ['m1'], label_ids: ['label-1'] }, { removeLabels: async () => 1 }, { messageIds: ['m1'], labelIds: ['label-1'], succeededCount: 1 }],
    ['archive_emails', { message_ids: ['m1'] }, { archiveEmails: async () => 1 }, { messageIds: ['m1'], succeededCount: 1 }],
    ['trash_emails', { message_ids: ['m1'] }, { trashEmails: async () => 1 }, { messageIds: ['m1'], succeededCount: 1 }],
    ['create_label', { name: 'Review' }, { createLabel: async () => label }, { label, succeededCount: 1 }],
    ['delete_label', { label_id: 'label-1' }, { deleteLabel: async () => {} }, { labelId: 'label-1', succeededCount: 1 }],
    ['list_blocked_senders', {}, { listFilters: async () => [filter] }, { filters: [filter], count: 1 }],
    ['list_blocked_senders', {}, { listFilters: async () => [] }, { filters: [], count: 0 }],
    ['unblock_sender', { filter_id: 'filter-1' }, { deleteFilter: async () => {} }, { filterIds: ['filter-1'], succeededCount: 1 }],
    ['unblock_sender', { sender: 'sender@example.test' }, { listFilters: async () => [filter], deleteFilter: async () => {} }, { sender: 'sender@example.test', filterIds: ['filter-1'], succeededCount: 1 }],
    ['unblock_sender', { sender: 'unknown@example.test' }, { listFilters: async () => [filter] }, { sender: 'unknown@example.test', filterIds: [], succeededCount: 0 }],
    ['create_draft', { to: 'to@example.test', subject: 'Subject', body: 'Body' }, { createDraft: async () => ({ draftId: 'draft-1', threadId: 'thread-1' }) }, { draftId: 'draft-1', threadId: 'thread-1', attachmentCount: 0, succeededCount: 1 }],
    ['delete_drafts', { draft_ids: ['draft-1', 'draft-2'] }, { deleteDrafts: async () => 2 }, { draftIds: ['draft-1', 'draft-2'], succeededCount: 2 }],
    ['send_draft', { draft_id: 'draft-1' }, { sendDraft: async () => ({ messageId: 'message-1', threadId: 'thread-1' }) }, { draftId: 'draft-1', messageId: 'message-1', threadId: 'thread-1', succeededCount: 1 }],
    ['list_drafts', {}, { listDrafts: async () => [draft] }, { drafts: [draft], count: 1 }],
    ['list_drafts', {}, { listDrafts: async () => [] }, { drafts: [], count: 0 }],
    ['search_drafts', { query: 'subject:Draft' }, { searchDrafts: async () => [draft] }, { query: 'subject:Draft', drafts: [draft], count: 1 }],
    ['search_drafts', { query: 'missing' }, { searchDrafts: async () => [] }, { query: 'missing', drafts: [], count: 0 }],
    ['send_email', { to: 'to@example.test', subject: 'Subject', body: 'Body' }, { sendEmail: async () => ({ messageId: 'message-1', threadId: 'thread-1' }) }, { messageId: 'message-1', threadId: 'thread-1', attachmentCount: 0, succeededCount: 1 }],
  ];
  for (const [tool, args, methods, expected] of cases) {
    await t.test(`${tool}: ${Object.keys(expected).join(', ')}`, async () => {
      const result = await mockApp(methods).callTool(tool, { account: 'work', ...args });
      assert.deepEqual(structured(result), { account: 'work', ...expected });
    });
  }
});

test('blocking reports the created filter and sanitizes a partial sweep warning', async () => {
  const app = mockApp({ createBlockFilter: async () => filter, searchEmails: async () => { throw new Error(secret); } });
  const data = structured(await app.callTool('block_sender', { account: 'work', sender: 'sender@example.test' }));
  assert.deepEqual(data.filter, filter);
  assert.equal(data.succeededCount, 1);
  assert.equal(data.trashedCount, 0);
  assert.equal(data.warnings[0].code, 'ACCOUNT_ERROR');
  assert.equal(data.errors[0].code, 'SWEEP_FAILED');
  assert.ok(!JSON.stringify(data).includes(secret));
  const archived = structured(await mockApp({ createBlockFilter: async () => filter }).callTool('block_sender', { account: 'work', sender: 'sender@example.test', action: 'archive' }));
  assert.equal(archived.warnings[0].code, 'SWEEP_SKIPPED');
});

test('muting exposes thread-only, filter-created and generic-subject warning outcomes', async () => {
  const methods = { modifyThread: async () => {}, getThreadSubject: async () => 'Quarterly planning discussion', createFilter: async () => filter };
  const data = structured(await mockApp(methods).callTool('mute_thread', { account: 'work', thread_id: 'thread-1' }));
  assert.equal(data.thread_archived, true);
  assert.equal(data.filter_created, true);
  assert.equal(data.filter_id, 'filter-1');
  const threadOnly = structured(await mockApp({ modifyThread: methods.modifyThread }).callTool('mute_thread', { account: 'work', thread_id: 'thread-1', scope: 'thread_only' }));
  assert.equal(threadOnly.filter_created, false);
  const warned = structured(await mockApp({ ...methods, getThreadSubject: async () => '' }).callTool('mute_thread', { account: 'work', thread_id: 'thread-1' }));
  assert.equal(warned.filter_created, false);
  assert.equal(warned.warnings.length, 1);
});

test('every unsubscribe outcome is structured without leaking exception details', async (t) => {
  const originalFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = originalFetch; });
  const oneClick = { 'List-Unsubscribe': '<https://example.test/unsubscribe>', 'List-Unsubscribe-Post': 'List-Unsubscribe=One-Click' };
  const mailto = { 'List-Unsubscribe': '<mailto:leave@example.test?subject=unsubscribe>' };
  const cases = [
    ['missing header', {}, {}, null, {}, 'unavailable'],
    ['one-click dry run', oneClick, { dry_run: true }, null, {}, 'dry_run'],
    ['one-click success', oneClick, {}, { status: 204, statusText: 'No Content' }, {}, 'success'],
    ['one-click non-2xx', oneClick, {}, { status: 503, statusText: 'Unavailable' }, {}, 'failed'],
    ['one-click fetch error', oneClick, {}, new Error(secret), {}, 'failed'],
    ['mailto dry run', mailto, { dry_run: true }, null, {}, 'dry_run'],
    ['mailto sent', mailto, {}, null, { sendEmail: async () => ({ messageId: 'sent-1', threadId: 'sent-thread' }) }, 'sent'],
    ['mailto failure', mailto, {}, null, { sendEmail: async () => { throw new Error(secret); } }, 'failed'],
    ['manual URL', { 'List-Unsubscribe': oneClick['List-Unsubscribe'] }, {}, null, {}, 'manual_confirmation_required'],
    ['unparseable header', { 'List-Unsubscribe': 'invalid' }, {}, null, {}, 'unavailable'],
  ];
  for (const [name, headers, args, response, methods, status] of cases) {
    await t.test(name, async () => {
      globalThis.fetch = async (url, init) => {
        assert.equal(url, 'https://example.test/unsubscribe');
        assert.equal(init.method, 'POST');
        assert.equal(init.body, 'List-Unsubscribe=One-Click');
        if (response instanceof Error) throw response;
        assert.ok(response, 'Unexpected network request');
        return response;
      };
      const data = structured(await mockApp({ getMessageHeaders: async () => headers, ...methods }).callTool('unsubscribe_from_email', { account: 'work', message_id: 'message-1', ...args }));
      assert.equal(data.account, 'work');
      assert.equal(data.messageId, 'message-1');
      assert.equal(data.status, status);
      assert.equal(data.succeededCount, ['success', 'sent'].includes(status) ? 1 : 0);
      if (status === 'failed') assert.equal(data.errors[0].code, 'ACCOUNT_ERROR');
      if (status === 'sent') assert.equal(data.sentMessageId, 'sent-1');
      assert.ok(!JSON.stringify(data).includes(secret));
    });
  }
});

test('attachment outputs contain saved records, empty arrays, and sanitized partial failures', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ghub-gmail-output-'));
  const previous = process.env.MCP_ATTACHMENTS_DIR;
  process.env.MCP_ATTACHMENTS_DIR = directory;
  t.after(async () => {
    if (previous === undefined) delete process.env.MCP_ATTACHMENTS_DIR;
    else process.env.MCP_ATTACHMENTS_DIR = previous;
    await fs.rm(directory, { recursive: true, force: true });
  });
  const metadata = { id: 'attachment-1', filename: 'notes.txt', contentType: 'text/plain', sizeBytes: 5, isInline: false };
  const fetched = { bytes: Buffer.from('hello'), metadata };
  const single = structured(await mockApp({ getAttachment: async () => fetched }).callTool('get_attachment', { account: 'work', email_id: 'message-1', attachment_id: metadata.id }));
  assert.equal(single.attachment.id, metadata.id);
  assert.equal(single.attachment.text, 'hello');
  assert.equal(await fs.readFile(single.attachment.savedPath, 'utf8'), 'hello');
  assert.equal(single.messageId, 'message-1');
  const partial = structured(await mockApp({ fetchAllAttachments: async () => [fetched, { metadata: { ...metadata, id: 'failed-attachment' }, error: secret }] }).callTool('get_all_attachments', { account: 'work', email_id: 'message-1' }));
  assert.equal(partial.count, 1);
  assert.equal(partial.succeededCount, 1);
  assert.equal(partial.attachments[0].text, 'hello');
  assert.equal(partial.errors[0].attachmentId, 'failed-attachment');
  assert.equal(partial.errors[0].code, 'ACCOUNT_ERROR');
  assert.ok(!JSON.stringify(partial).includes(secret));
  const empty = structured(await mockApp({ fetchAllAttachments: async () => [] }).callTool('get_all_attachments', { account: 'work', email_id: 'message-1' }));
  assert.deepEqual(empty.attachments, []);
  assert.deepEqual(empty.errors, []);
  assert.equal(empty.count, 0);
});

test('every Gmail handler return provides structured content, including uncommon branches', async () => {
  const source = await fs.readFile(new URL('../src/app.ts', import.meta.url), 'utf8');
  const ast = ts.createSourceFile('app.ts', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const names = new Set(['handleReadEmails', 'handleSearchEmails', 'handleGetThread', 'handleGetLabels', 'handleMarkAsRead', 'handleAddLabels', 'handleRemoveLabels', 'handleArchiveEmails', 'handleTrashEmails', 'handleCreateLabel', 'handleDeleteLabel', 'handleBlockSender', 'handleListBlockedSenders', 'handleUnblockSender', 'handleUnsubscribeFromEmail', 'handleMuteThread', 'handleCreateDraft', 'handleDeleteDrafts', 'handleSendDraft', 'handleListDrafts', 'handleSearchDrafts', 'handleSendEmail', 'handleGetAttachment', 'handleGetAllAttachments']);
  let count = 0;
  function visit(node) {
    if (ts.isMethodDeclaration(node) && names.has(node.name.getText(ast))) {
      function inspect(child) {
        if (ts.isCallExpression(child) && child.expression.getText(ast) === 'textResult') {
          assert.equal(child.arguments.length, 2, `${node.name.getText(ast)} must return structured content`);
          count++;
        }
        ts.forEachChild(child, inspect);
      }
      ts.forEachChild(node, inspect);
    } else ts.forEachChild(node, visit);
  }
  visit(ast);
  assert.equal(count, 37);
});
