import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { runCli } from '../dist/cli.js';
import { GmailMultiInboxServer } from '../dist/app.js';
import { safeError, project, validateInput } from '../dist/cli-contract.js';

const registry = new GmailMultiInboxServer().listTools();
function harness(result = { structuredContent: { account: 'work', messageId: 'sent-id', succeededCount: 1 } }) {
  const calls = [];
  const writes = [];
  const io = {
    stdout: (text) => writes.push(text),
    readInput: async (file) => { assert.equal(file, '-'); return '{"account":"work"}'; },
    createApp: async () => ({ listTools: () => registry, callTool: async (...args) => { calls.push(args); if (result instanceof Error) throw result; return result; } }),
  };
  return { calls, writes, io, value: () => { assert.equal(writes.length, 1); assert.ok(writes[0].endsWith('\n')); return JSON.parse(writes[0]); } };
}

test('help and full 60-command discovery are JSON with no API calls', async () => {
  for (const args of [['--help'], ['tools']]) {
    const h = harness(); assert.equal(await runCli(args, h.io), 0); const value = h.value();
    assert.equal(value.ok, true); assert.equal(h.calls.length, 0);
    if (args[0] === 'tools') assert.deepEqual(value.data.tools.map((x) => x.name), registry.map((x) => x.name));
  }
});

test('every tool has input/output discovery and a validating dry-run example', async () => {
  for (const tool of registry) {
    const h = harness(); assert.equal(await runCli(['schema', tool.name], h.io), 0);
    const schema = h.value().data;
    assert.deepEqual(schema.inputSchema.required, tool.inputSchema.required);
    assert.equal(schema.outputSchema.properties.data.type, 'object');
    const example = schema.example.argv.slice(1);
    const dry = harness(); assert.equal(await runCli([...example, '--dry-run'], dry.io), 0, tool.name + ':' + dry.writes);
    assert.equal(dry.calls.length, 0);
  }
});

const invalid = [
  ['no-such-command'], ['call', 'no_such_tool'], ['call'], ['tools', 'unexpected'],
  ['call', 'list_accounts', '--bad'], ['call', 'list_accounts', '--json'],
  ['call', 'list_accounts', '--json', '{bad'], ['call', 'list_accounts', '--json', '[]'],
  ['call', 'list_accounts', '--json', 'null'], ['call', 'list_accounts', '--json', '{"secret":"NEVER_ECHO"}'],
  ['call', 'search_emails', '--json', '{"query":1}'],
  ['call', 'search_emails', '--json', '{"query":"q","max_results":0}'],
  ['call', 'search_emails', '--json', '{"query":"q","max_results":101}'],
  ['call', 'search_emails', '--json', '{"query":"q","max_results":1.5}'],
  ['call', 'read_emails', '--json', '{"page_token":"page"}'],
  ['call', 'list_drive_files', '--json', '{"page_token":"page"}'],
  ['call', 'send_email', '--json', '{"to":"x@y.test","subject":"s","body":"b"}'],
  ['call', 'trash_emails', '--json', '{"account":"work","message_ids":[]}'],
  ['call', 'trash_emails', '--json', '{"account":"work","message_ids":[1]}'],
  ['call', 'trash_emails', '--json', '{"account":"../work","message_ids":["m"]}'],
  ['call', 'create_event', '--json', '{"account":"work","summary":"s"}'],
  ['call', 'unblock_sender', '--json', '{"account":"work"}'],
  ['call', 'docs_apply_style', '--json', '{"account":"work","document_id":"d","start_index":4,"end_index":2,"style":"NORMAL_TEXT"}'],
  ['call', 'sheets_insert_dimension', '--json', '{"account":"work","spreadsheet_id":"s","sheet_title":"s","dimension":"ROWS","start_index":-1}'],
  ['call', 'list_accounts', '--limit', '0'], ['call', 'list_accounts', '--limit', 'NaN'],
  ['call', 'list_accounts', '--limit', '2', '--limit', '3'],
  ['call', 'list_accounts', '--json', '{}', '--input', '-'],
  ['call', 'list_accounts', '--fields', '__proto__.x'],
  ['call', 'list_accounts', '--json', '{"__proto__":{}}'],
  ['call', 'list_accounts', '--json', '{"toString":"must-be-unknown"}'],
  ['call', 'list_accounts', '--json', '{"valueOf":"must-be-unknown"}'],
];
for (const [index, args] of invalid.entries()) test(`invalid invocation ${index} fails before a tool call`, async () => {
  const h = harness(); assert.equal(await runCli(args, h.io), 2); assert.equal(h.calls.length, 0);
  assert.equal(h.value().ok, false); assert.ok(!h.writes[0].includes('NEVER_ECHO'));
});

test('explicit stdin, no implicit input, config argument and strict unknown flags', async () => {
  const h = harness({ structuredContent: { labels: [], count: 0 } });
  assert.equal(await runCli(['call', 'get_labels', '--input', '-', '--config-dir', '/tmp/isolated'], h.io), 0);
  assert.deepEqual(h.calls, [['get_labels', { account: 'work' }]]);
  const noInput = harness({ structuredContent: { accounts: [] } });
  noInput.io.readInput = async () => { throw new Error('must not read stdin'); };
  assert.equal(await runCli(['call', 'list_accounts'], noInput.io), 0);
});

test('field projection preserves arrays and pagination/error recovery metadata', async () => {
  const h = harness({ structuredContent: { accounts: ['work'], emails: [{ id: 'a', subject: 's', body: 'hidden' }], nextPageToken: 'cursor', count: 1 } });
  assert.equal(await runCli(['call', 'read_emails', '--json', '{"account":"work"}', '--fields', 'emails.id,emails.subject'], h.io), 0);
  assert.deepEqual(h.value().data, { emails: [{ id: 'a', subject: 's' }] });
  assert.equal(h.value().meta.nextPageToken, 'cursor');
  assert.equal(h.calls[0][1].max_results, 20);
});

test('bounds output arrays, strings and bytes with explicit truncation metadata', async () => {
  const emails = Array.from({ length: 100 }, (_, id) => ({ id: String(id), body: 'ü'.repeat(10000) }));
  const h = harness({ structuredContent: { emails, count: 100 } });
  assert.equal(await runCli(['call', 'read_emails', '--limit', '3', '--max-string-length', '100', '--max-output-bytes', '2048'], h.io), 0);
  const value = h.value();
  assert.ok(Buffer.byteLength(h.writes[0]) <= 2049);
  assert.ok(value.meta.truncation.length > 0);
  assert.ok(value.data.emails.length <= 3);
  assert.ok(value.data.emails[0].body.length <= 100);
});

test('oversized completed mutation never becomes a retryable failure', async () => {
  const data = Object.fromEntries(Array.from({ length: 500 }, (_, i) => [`field${i}`, 'x'.repeat(100)]));
  const h = harness({ structuredContent: data });
  assert.equal(await runCli(['call', 'send_draft', '--json', '{"account":"work","draft_id":"d"}', '--max-output-bytes', '1024'], h.io), 0);
  assert.equal(h.value().meta.operationCompleted, true);
  assert.equal(h.value().meta.resultOmitted, true);
  assert.match(h.value().meta.recovery, /Do not repeat/);
});

test('mixed successful-empty account and failed account remains partial', async () => {
  const h = harness({ structuredContent: { accounts: ['work', 'home'], successfulAccounts: ['work'], emails: [], count: 0, errors: [{ account: 'home', code: 'AUTH_REQUIRED', message: 'Needs auth', exitCode: 3 }] } });
  assert.equal(await runCli(['call', 'read_emails', '--fields', 'emails.id'], h.io), 6);
  assert.equal(h.value().error.code, 'PARTIAL_FAILURE');
  assert.equal(h.value().meta.errors[0].account, 'home');
});

test('all failed aggregation is nonzero, single-account errors preserve exit classification', async () => {
  const base = { accounts: ['work', 'home'], successfulAccounts: [], emails: [], count: 0, errors: [{ account: 'work', code: 'AUTH_REQUIRED', message: 'Needs auth', exitCode: 3, retryable: false }] };
  const h = harness({ structuredContent: base }); assert.equal(await runCli(['call', 'read_emails'], h.io), 4);
  for (const [code, exitCode] of [['AUTH_REQUIRED', 3], ['UPSTREAM_UNAVAILABLE', 5]]) {
    const single = harness({ structuredContent: { ...base, accounts: ['work'], errors: [{ ...base.errors[0], code, exitCode }] } });
    assert.equal(await runCli(['call', 'read_emails', '--json', '{"account":"work"}'], single.io), exitCode);
    assert.equal(single.value().error.code, code);
  }
});

for (const [status, code, exit] of [[401, 'AUTH_REQUIRED', 3], [403, 'PERMISSION_DENIED', 4], [404, 'NOT_FOUND', 4], [429, 'UPSTREAM_UNAVAILABLE', 5], [503, 'UPSTREAM_UNAVAILABLE', 5], [400, 'API_ERROR', 4]]) {
  test(`provider ${status} is sanitized and exits ${exit}`, async () => {
    const error = Object.assign(new Error('SENSITIVE_TOKEN'), { response: { status, config: { Authorization: 'SECRET' } } });
    const h = harness(error); assert.equal(await runCli(['call', 'get_labels', '--json', '{"account":"work"}'], h.io), exit);
    assert.equal(h.value().error.code, code); assert.doesNotMatch(h.writes[0], /SENSITIVE_TOKEN|SECRET/);
  });
}

test('mutations are never auto-retried and uncertain failures have reconciliation advice', async () => {
  const error = Object.assign(new Error('network failed'), { code: 'ETIMEDOUT' });
  const h = harness(error);
  assert.equal(await runCli(['call', 'send_draft', '--json', '{"account":"work","draft_id":"d"}'], h.io), 5);
  assert.equal(h.calls.length, 1); assert.equal(h.value().error.retryable, false);
  assert.match(h.value().error.recovery, /may have taken effect/);
});

test('unstructured or MCP error results cannot be reported as CLI success', async () => {
  for (const result of [{ content: [{ type: 'text', text: 'secret' }] }, { isError: true, structuredContent: { secret: 'hidden' } }]) {
    const h = harness(result); assert.equal(await runCli(['call', 'list_accounts'], h.io), 1);
    assert.doesNotMatch(h.writes[0], /secret|hidden/);
  }
});

test('real executable reads protected JSON file/stdin and emits one JSON line only', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ghub-cli-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'input.json'); await fs.writeFile(file, '{}', { mode: 0o600 });
  for (const input of [file, '-']) {
    const child = spawnSync(process.execPath, ['dist/index.js', 'call', 'list_accounts', '--config-dir', directory, '--input', input], { input: '{}', encoding: 'utf8', timeout: 10000 });
    assert.equal(child.status, 0, child.stderr); assert.equal(child.stderr, '');
    assert.equal(child.stdout.trim().split('\n').length, 1); assert.deepEqual(JSON.parse(child.stdout).data.accounts, []);
  }
  const badFile = spawnSync(process.execPath, ['dist/index.js', 'call', 'list_accounts', '--input', directory], { encoding: 'utf8', timeout: 10000 });
  assert.equal(badFile.status, 2); assert.equal(JSON.parse(badFile.stdout).error.code, 'INVALID_ARGUMENT');
});

test('projection has no prototype pollution and does not flatten arrays', () => {
  assert.deepEqual(project({ items: [{ child: { id: 'x', name: 'n' } }] }, [['items', 'child', 'id']]), { items: [{ child: { id: 'x' } }] });
  assert.throws(() => validateInput(JSON.parse('{"constructor":{}}'), { type: 'object', properties: {}, additionalProperties: false }));
  assert.equal(safeError(Object.assign(new Error('x'), { code: 'ECONNRESET' }), true).retryable, false);
});


test('known permission status takes priority over OAuth-related error prose', async () => {
  const error = Object.assign(new Error('Enable the Google Drive API for this disabled OAuth client'), { code: 403 });
  const h = harness(error);
  assert.equal(await runCli(['call', 'get_labels', '--json', '{"account":"work"}'], h.io), 4);
  assert.equal(h.value().error.code, 'PERMISSION_DENIED');
  assert.match(h.value().error.recovery, /API enablement/);
});


test('output truncation never hands an unsafe continuation past omitted records', async () => {
  const h = harness({ structuredContent: { account: 'work', emails: [{ id: '1' }, { id: '2' }], count: 2, nextPageToken: 'skip-two' } });
  assert.equal(await runCli(['call', 'read_emails', '--json', '{"account":"work","max_results":2}', '--limit', '1'], h.io), 0);
  const result = h.value();
  assert.equal(result.meta.paginationBlocked, true);
  assert.equal(result.meta.nextPageToken, undefined);
  assert.equal(result.data.nextPageToken, undefined);
  assert.match(result.meta.recovery, /Replay/);
});

test('completed block filter with failed requested sweep is partial and preserves the new filter ID', async () => {
  const h = harness({ structuredContent: { account: 'work', filter: { id: 'f1' }, succeededCount: 1, errors: [{ code: 'SWEEP_FAILED' }] } });
  assert.equal(await runCli(['call', 'block_sender', '--json', '{"account":"work","sender":"x@y.test"}'], h.io), 6);
  assert.equal(h.value().data.filter.id, 'f1');
  assert.match(h.value().error.recovery, /some changes may have succeeded/);
});
