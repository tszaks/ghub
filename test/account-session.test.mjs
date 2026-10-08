import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OAuth2Client } from 'google-auth-library';
import { AccountSession } from '../dist/account-session.js';
import { GmailAccountClient, GOOGLE_ACCOUNT_SCOPES } from '../dist/gmail-client.js';
import { GmailMultiInboxServer } from '../dist/app.js';
import { getDefaultAccountPaths, saveAccountsConfig } from '../dist/config.js';

// Synthetic OAuth fixtures only. Tests never contact Google or read real accounts.
const credentials = {
  installed: {
    client_id: 'fixture-client-id',
    client_secret: 'fixture-client-secret',
    redirect_uris: ['http://localhost'],
  },
};

async function createSession(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ghub-account-session-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  return new AccountSession(root);
}

async function beginFixture(session, accountId = 'work') {
  return session.beginAuth({
    account_id: accountId,
    email: `${accountId}@example.test`,
    display_name: 'Example inbox',
    credentials_json: credentials,
  });
}

test('empty accounts are returned as structured, path-free data', async (t) => {
  const session = await createSession(t);
  assert.deepEqual(await session.listAccounts(), { defaultAccount: null, accounts: [] });
  assert.deepEqual(await session.loadConfig(), { defaultAccount: null, accounts: [] });
});

test('listAccounts reports file presence without reading or returning secrets and paths', async (t) => {
  const session = await createSession(t);
  const records = [
    { id: 'ready', email: 'ready@example.test', displayName: 'Ready', enabled: true },
    { id: 'missing', email: 'missing@example.test', enabled: true },
    { id: 'disabled', email: 'disabled@example.test', enabled: false },
  ];
  await saveAccountsConfig(session.configRoot, { defaultAccount: 'ready', accounts: records });
  for (const record of [records[0], records[2]]) {
    const paths = getDefaultAccountPaths(session.configRoot, record.id);
    await fs.mkdir(paths.accountDir, { recursive: true });
    // Invalid JSON is intentional: listing must not parse either file.
    await fs.writeFile(paths.credentialsPath, 'unread-credential-fixture');
    await fs.writeFile(paths.tokenPath, 'unread-token-fixture');
  }

  const result = await session.listAccounts();
  assert.equal(result.defaultAccount, 'ready');
  assert.deepEqual(result.accounts.map(({ status }) => status), ['ready', 'needs-auth-files', 'disabled']);
  assert.deepEqual(result.accounts.map(({ ready }) => ready), [true, false, false]);
  assert.deepEqual(result.accounts.map(({ hasTokenFile }) => hasTokenFile), [true, false, true]);
  const serialized = JSON.stringify(result);
  assert.ok(!serialized.includes(session.configRoot));
  assert.doesNotMatch(serialized, /credentialPath|tokenPath|unread-credential|unread-token/);
});

test('beginAuth preserves scopes and registers a disabled account without exchanging a code', async (t) => {
  const session = await createSession(t);
  const exchange = t.mock.method(OAuth2Client.prototype, 'getToken', async () => {
    throw new Error('beginAuth must not exchange a code');
  });
  const result = await beginFixture(session);
  assert.deepEqual(Object.keys(result).sort(), ['accountId', 'authUrl', 'email']);
  assert.equal(result.accountId, 'work');
  assert.equal(result.email, 'work@example.test');
  const url = new URL(result.authUrl);
  assert.equal(url.origin, 'https://accounts.google.com');
  assert.deepEqual(new Set(url.searchParams.get('scope').split(' ')), new Set(GOOGLE_ACCOUNT_SCOPES));
  assert.equal(url.searchParams.get('access_type'), 'offline');
  assert.equal(url.searchParams.get('prompt'), 'consent');
  assert.ok(!result.authUrl.includes(credentials.installed.client_secret));
  assert.equal(exchange.mock.callCount(), 0);
  const config = await session.loadConfig();
  assert.equal(config.defaultAccount, null);
  assert.equal(config.accounts[0].enabled, false);
  assert.equal(config.accounts[0].displayName, 'Example inbox');
  const paths = getDefaultAccountPaths(session.configRoot, 'work');
  assert.deepEqual(JSON.parse(await fs.readFile(paths.credentialsPath, 'utf8')), credentials);
  await assert.rejects(fs.access(paths.tokenPath), { code: 'ENOENT' });
});

test('beginAuth accepts JSON text, a credentials file, and existing stored credentials', async (t) => {
  const session = await createSession(t);
  await session.beginAuth({
    account_id: 'json', email: 'json@example.test', credentials_json: JSON.stringify(credentials),
  });
  const source = path.join(session.configRoot, 'synthetic-client.json');
  await fs.writeFile(source, JSON.stringify(credentials));
  await session.beginAuth({ account_id: 'file', email: 'file@example.test', credentials_path: source });
  await session.beginAuth({ account_id: 'file', email: 'updated@example.test' });
  const config = await session.loadConfig();
  assert.equal(config.accounts.length, 2);
  assert.equal(config.accounts[1].email, 'updated@example.test');
});

test('beginAuth validates required fields, ids, and credentials before persisting invalid content', async (t) => {
  const session = await createSession(t);
  await assert.rejects(session.beginAuth({ account_id: '', email: 'work@example.test' }), /account_id is required/);
  await assert.rejects(session.beginAuth({ account_id: 'work', email: '' }), /email is required/);
  await assert.rejects(session.beginAuth({ account_id: '../escape', email: 'work@example.test' }), /Invalid account id/);
  await assert.rejects(session.beginAuth({ account_id: 'work', email: 'work@example.test', credentials_json: 42 }), /JSON string or object/);
  await beginFixture(session);
  const paths = getDefaultAccountPaths(session.configRoot, 'work');
  await assert.rejects(session.beginAuth({ account_id: 'work', email: 'work@example.test', credentials_json: {} }), /client_id and client_secret/);
  assert.deepEqual(JSON.parse(await fs.readFile(paths.credentialsPath, 'utf8')), credentials);
});

test('getClientForAccount delegates to the existing shared Gmail client', async (t) => {
  const session = await createSession(t);
  const account = { id: 'work', email: 'work@example.test', enabled: true };
  const expected = { syntheticClient: true };
  const create = t.mock.method(GmailAccountClient, 'create', async () => expected);
  assert.equal(await session.getClientForAccount(account), expected);
  assert.deepEqual(create.mock.calls[0].arguments, [session.configRoot, account]);
});

test('finishAuth stores mocked tokens, verifies profile, and sets the initial default', async (t) => {
  const session = await createSession(t);
  await beginFixture(session);
  const tokens = { access_token: 'synthetic-access', refresh_token: 'synthetic-refresh' };
  const exchange = t.mock.method(OAuth2Client.prototype, 'getToken', async () => ({ tokens }));
  const create = t.mock.method(GmailAccountClient, 'create', async () => ({
    getProfileEmail: async () => 'verified@example.test',
  }));

  const result = await session.finishAuth({ account_id: 'work', authorization_code: 'synthetic-code' });
  assert.deepEqual(result, { accountId: 'work', email: 'verified@example.test', enabled: true });
  assert.deepEqual(exchange.mock.calls[0].arguments, ['synthetic-code']);
  assert.equal(create.mock.callCount(), 1);
  const config = await session.loadConfig();
  assert.equal(config.defaultAccount, 'work');
  assert.equal(config.accounts[0].enabled, true);
  assert.equal(config.accounts[0].email, 'verified@example.test');
  assert.deepEqual(JSON.parse(await fs.readFile(config.accounts[0].tokenPath, 'utf8')), tokens);
  assert.doesNotMatch(JSON.stringify(result), /synthetic-access|synthetic-refresh|tokenPath|credentialPath/);
});

test('finishAuth preserves an existing default and falls back if profile lookup fails', async (t) => {
  const session = await createSession(t);
  await beginFixture(session, 'personal');
  await beginFixture(session, 'work');
  const config = await session.loadConfig();
  config.defaultAccount = 'personal';
  await saveAccountsConfig(session.configRoot, config);
  t.mock.method(OAuth2Client.prototype, 'getToken', async () => ({ tokens: { refresh_token: 'synthetic-refresh' } }));
  t.mock.method(GmailAccountClient, 'create', async () => ({
    getProfileEmail: async () => { throw new Error('Synthetic profile failure'); },
  }));
  assert.deepEqual(await session.finishAuth({ account_id: 'work', authorization_code: 'synthetic-code' }), {
    accountId: 'work', email: 'work@example.test', enabled: true,
  });
  assert.equal((await session.loadConfig()).defaultAccount, 'personal');
});

test('finishAuth rejects missing tokens and failed exchanges without enabling the account', async (t) => {
  const session = await createSession(t);
  await beginFixture(session);
  const exchange = t.mock.method(OAuth2Client.prototype, 'getToken', async () => ({ tokens: {} }));
  await assert.rejects(session.finishAuth({ account_id: 'work', authorization_code: 'synthetic-code' }), /no token payload/);
  exchange.mock.mockImplementation(async () => { throw new Error('Synthetic exchange rejection'); });
  await assert.rejects(session.finishAuth({ account_id: 'work', authorization_code: 'synthetic-code' }), /Synthetic exchange rejection/);
  const config = await session.loadConfig();
  assert.equal(config.accounts[0].enabled, false);
  assert.equal(config.defaultAccount, null);
  await assert.rejects(fs.access(getDefaultAccountPaths(session.configRoot, 'work').tokenPath), { code: 'ENOENT' });
});

test('finishAuth validates account and authorization code before exchanging', async (t) => {
  const session = await createSession(t);
  const exchange = t.mock.method(OAuth2Client.prototype, 'getToken', async () => {
    throw new Error('Invalid arguments must not exchange a code');
  });
  await assert.rejects(session.finishAuth({ account_id: '', authorization_code: 'synthetic-code' }), /account_id is required/);
  await assert.rejects(session.finishAuth({ account_id: 'work', authorization_code: '' }), /authorization_code is required/);
  await assert.rejects(session.finishAuth({ account_id: '../escape', authorization_code: 'synthetic-code' }), /Invalid account id/);
  await assert.rejects(session.finishAuth({ account_id: 'unknown', authorization_code: 'synthetic-code' }), /Unknown account/);
  assert.equal(exchange.mock.callCount(), 0);
});

test('MCP account and auth handlers preserve text and add the same structured results', async (t) => {
  const session = await createSession(t);
  const app = new GmailMultiInboxServer(session.configRoot);
  const empty = await app.callTool('list_accounts', {});
  assert.match(empty.content[0].text, /No accounts configured yet/);
  assert.deepEqual(empty.structuredContent, { defaultAccount: null, accounts: [] });

  const started = await app.callTool('begin_account_auth', {
    account_id: 'work', email: 'work@example.test', credentials_json: credentials,
  });
  assert.match(started.content[0].text, /Google Account OAuth Started/);
  assert.match(started.content[0].text, /Gmail, Google Drive, Sheets, Docs, and Calendar/);
  assert.equal(started.structuredContent.accountId, 'work');
  assert.equal(started.structuredContent.authUrl, started.content[0].text.split('\n').find((line) => line.startsWith('https://')));

  t.mock.method(OAuth2Client.prototype, 'getToken', async () => ({ tokens: { access_token: 'synthetic-access' } }));
  t.mock.method(GmailAccountClient, 'create', async () => ({ getProfileEmail: async () => 'work@example.test' }));
  const finished = await app.callTool('finish_account_auth', { account_id: 'work', authorization_code: 'synthetic-code' });
  assert.match(finished.content[0].text, /Google account OAuth completed/);
  assert.deepEqual(finished.structuredContent, { accountId: 'work', email: 'work@example.test', enabled: true });

  const listed = await app.callTool('list_accounts', {});
  assert.match(listed.content[0].text, /work \(default\)/);
  assert.deepEqual(listed.structuredContent, await session.listAccounts());
  assert.ok(!JSON.stringify(listed.structuredContent).includes(session.configRoot));
});
