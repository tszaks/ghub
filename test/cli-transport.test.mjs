import test from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import {
  configureCliTransport, getCliRequestTimeout, createOAuthClientFromCredentials,
  GmailAccountClient,
} from '../dist/gmail-client.js';
import { safeError } from '../dist/core-errors.js';

// All credentials and tokens in these tests are invented. Every HTTP request is
// intercepted by a Gaxios adapter; no live OAuth grants or Google calls are made.
const credentials = { installed: {
  client_id: 'offline-test-client', client_secret: 'offline-test-secret',
  redirect_uris: ['http://localhost'],
} };
function response(config, data = {}, status = 200) {
  return { config, data, status, statusText: String(status), headers: {}, request: { responseURL: String(config.url) } };
}
function cliClient(t, adapter, timeoutMs = 1234) {
  configureCliTransport({ timeoutMs });
  t.after(() => configureCliTransport());
  const client = createOAuthClientFromCredentials({ credentials });
  assert.ok(client.gaxios);
  client.gaxios.defaults.adapter = adapter;
  return client;
}
function assertPolicy(options, timeout = 1234) {
  assert.equal(options.timeout, timeout);
  assert.equal(options.retry, false);
  assert.equal(options.retryConfig.retry, 0);
}

test('CLI transport is opt-in and leaves unconfigured MCP OAuth clients unchanged', (t) => {
  t.after(() => configureCliTransport());
  configureCliTransport();
  const before = createOAuthClientFromCredentials({ credentials });
  assert.equal(before.constructor, OAuth2Client);
  assert.equal(getCliRequestTimeout(), undefined);
  assert.equal(before.gaxios.defaults.timeout, undefined);
  configureCliTransport({ timeoutMs: 9876 });
  assert.equal(getCliRequestTimeout(), 9876);
  const configured = createOAuthClientFromCredentials({ credentials });
  assert.equal(configured.gaxios.defaults.timeout, 9876);
  assert.equal(configured.gaxios.defaults.retry, false);
  configureCliTransport();
  const after = createOAuthClientFromCredentials({ credentials });
  assert.equal(after.constructor, OAuth2Client);
  assert.equal(after.gaxios.defaults.timeout, undefined);
  assert.equal(configured.gaxios.defaults.timeout, 9876);
  assert.throws(() => configureCliTransport({ timeoutMs: 0 }), /positive integer/);
});

test('authorization-code exchange obeys CLI timeout and cannot retry 503', async (t) => {
  const requests = [];
  const oauth = cliClient(t, async (options) => {
    requests.push(options);
    assertPolicy(options);
    return response(options, { error: 'offline-provider-error' }, 503);
  });
  await assert.rejects(oauth.getToken('offline-code'), (error) => error.response.status === 503);
  assert.equal(requests.length, 1);
  assert.equal(new URL(requests[0].url).pathname, '/token');
});

test('expired-token refresh and subsequent API request both obey CLI timeout', async (t) => {
  const requests = [];
  const oauth = cliClient(t, async (options) => {
    requests.push(options);
    assertPolicy(options, 5678);
    if (new URL(options.url).pathname === '/token') {
      return response(options, { access_token: 'refreshed-offline-token', expires_in: 3600, token_type: 'Bearer' });
    }
    assert.equal(options.headers.Authorization, 'Bearer refreshed-offline-token');
    return response(options, { id: 'sent-offline' });
  }, 5678);
  oauth.setCredentials({ access_token: 'expired-offline-token', refresh_token: 'offline-refresh', expiry_date: 1 });
  const gmail = google.gmail({ version: 'v1', auth: oauth });
  const result = await gmail.users.messages.send({ userId: 'me', requestBody: { raw: 'offline-content' } });
  assert.equal(result.data.id, 'sent-offline');
  assert.equal(requests.length, 2);
  assert.equal(new URL(requests[0].url).pathname, '/token');
  assert.match(new URL(requests[1].url).pathname, /messages\/send$/);
});

test('refresh failure is bounded and prevents the actual mutation', async (t) => {
  const requests = [];
  const oauth = cliClient(t, async (options) => {
    requests.push(options);
    assertPolicy(options);
    return response(options, { error: 'offline-provider-error' }, 503);
  });
  oauth.setCredentials({ access_token: 'expired-offline-token', refresh_token: 'offline-refresh', expiry_date: 1 });
  await assert.rejects(google.gmail({ version: 'v1', auth: oauth }).users.messages.send({ userId: 'me', requestBody: { raw: 'offline-content' } }));
  assert.equal(requests.length, 1);
  assert.equal(new URL(requests[0].url).pathname, '/token');
});

for (const status of [401, 403, 429, 503]) {
  test(`cached tokens without expiry cannot replay a Gmail mutation after ${status}`, async (t) => {
    const requests = [];
    const oauth = cliClient(t, async (options) => {
      requests.push(options);
      assertPolicy(options);
      // A regression would refresh and send again; return success on subsequent
      // requests so a retry bug fails the test promptly, without a retry loop.
      return requests.length === 1
        ? response(options, { error: { message: 'offline failure', code: status } }, status)
        : response(options, { access_token: 'unexpected-refresh', expires_in: 3600, id: 'unexpected-send' });
    });
    oauth.setCredentials({ access_token: 'offline-cached', refresh_token: 'offline-refresh' });
    oauth.forceRefreshOnFailure = true;
    const gmail = google.gmail({ version: 'v1', auth: oauth });
    await assert.rejects(gmail.users.messages.send({ userId: 'me', requestBody: { raw: 'offline-content' } }), (error) => error.response.status === status);
    assert.equal(requests.length, 1);
    assert.match(new URL(requests[0].url).pathname, /messages\/send$/);
  });
}

test('explicit retry options cannot override CLI policy, including callback API calls', async (t) => {
  let requests = 0;
  const oauth = cliClient(t, async (options) => {
    requests++;
    assertPolicy(options);
    return response(options, { error: { message: 'offline failure', code: 503 } }, requests === 1 ? 503 : 200);
  });
  oauth.setCredentials({ access_token: 'offline-cached', refresh_token: 'offline-refresh' });
  await new Promise((resolve, reject) => {
    oauth.request({
      url: 'https://www.googleapis.com/offline-test', method: 'POST',
      timeout: 999999, retry: true,
      retryConfig: { retry: 5, shouldRetry: () => true },
    }, (error, result) => {
      try {
        assert.ok(error);
        assert.equal(result.status, 503);
        resolve();
      } catch (failure) { reject(failure); }
    });
  });
  assert.equal(requests, 1);
});

test('CLI request keeps authenticated and caller headers', async (t) => {
  const oauth = cliClient(t, async (options) => {
    assertPolicy(options);
    assert.equal(options.headers.Authorization, 'Bearer offline-cached');
    assert.equal(options.headers['X-Goog-Api-Key'], 'offline-key');
    assert.equal(options.headers['x-goog-user-project'], 'offline-project');
    assert.equal(options.headers['x-test-header'], 'keep');
    return response(options, { ok: true });
  });
  oauth.setCredentials({ access_token: 'offline-cached' });
  oauth.apiKey = 'offline-key';
  oauth.quotaProjectId = 'offline-project';
  const result = await oauth.request({ url: 'https://www.googleapis.com/offline-test', headers: { 'x-test-header': 'keep' } });
  assert.deepEqual(result.data, { ok: true });
});

const driveOperations = [
  ['search', (client) => client.searchDriveFiles('test', 10)],
  ['list', (client) => client.listDriveFiles({})],
  ['get', (client) => client.getDriveFile('offline-file')],
  ['download', (client) => client.getDriveFileContent('offline-file')],
  ['upload', (client) => client.uploadDriveFile({ localPath: fileURLToPath(import.meta.url) })],
  ['create folder', (client) => client.createDriveFolder('offline-folder')],
  ['update', (client) => client.updateDriveFile('offline-file', { name: 'offline-renamed' })],
  ['trash', (client) => client.trashDriveFile('offline-file')],
  ['share', (client) => client.shareDriveFile('offline-file', { email: 'offline@example.invalid', role: 'reader', type: 'user' })],
];
for (const [status, expectedCode, expectedExit] of [[403, 'PERMISSION_DENIED', 4], [404, 'NOT_FOUND', 4], [503, 'UPSTREAM_UNAVAILABLE', 5]]) {
  for (const [operation, invoke] of driveOperations) {
    test(`Drive ${operation} preserves ${status} classification and cause without leaking provider data`, async () => {
      const original = Object.assign(new Error('SECRET_PROVIDER_MESSAGE'), {
        code: status,
        response: { status, data: { error: { message: 'SECRET_PROVIDER_MESSAGE' } }, config: { headers: { Authorization: 'SECRET_BEARER' } } },
      });
      const fail = async (params) => { params?.media?.body?.destroy(); throw original; };
      const drive = { files: { list: fail, create: fail, update: fail, get: async (params) => {
        if (operation === 'download' && params.alt !== 'media') {
          return { data: { id: 'offline-file', name: 'offline.txt', mimeType: 'text/plain' } };
        }
        return fail(params);
      } }, permissions: { create: fail } };
      const client = new GmailAccountClient({ id: 'offline', email: 'offline@example.invalid' }, {}, {}, drive, {}, {}, {});
      await assert.rejects(invoke(client), (wrapped) => {
        assert.equal(wrapped.cause, original);
        assert.equal(wrapped.code, status);
        assert.equal(wrapped.status, status);
        assert.deepEqual(wrapped.response, { status });
        const safe = safeError(wrapped, true);
        assert.equal(safe.code, expectedCode);
        assert.equal(safe.exitCode, expectedExit);
        assert.equal(safe.retryable, false);
        assert.ok(!JSON.stringify(safe).includes('SECRET'));
        assert.ok(!safe.message.includes('SECRET'));
        return true;
      });
    });
  }
}

test('Drive wrapper preserves string transport error codes and status-only errors', async () => {
  for (const original of [Object.assign(new Error('offline timeout'), { code: 'ETIMEDOUT' }), Object.assign(new Error('offline unavailable'), { status: 503 })]) {
    const drive = { files: { list: async () => { throw original; } } };
    const client = new GmailAccountClient({ id: 'offline', email: 'offline@example.invalid' }, {}, {}, drive, {}, {}, {});
    await assert.rejects(client.listDriveFiles({}), (wrapped) => {
      assert.equal(wrapped.cause, original);
      assert.equal(safeError(wrapped).code, 'UPSTREAM_UNAVAILABLE');
      return true;
    });
  }
});
