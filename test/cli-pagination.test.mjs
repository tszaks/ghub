import test from 'node:test';
import assert from 'node:assert/strict';
import { GmailAccountClient } from '../dist/gmail-client.js';

test('Gmail native pages pass tokens and retain provider continuation without hidden pagination', async () => {
  const lists = []; let inFlight = 0; let maxInFlight = 0;
  const account = { id: 'work', email: 'work@example.com', enabled: true };
  const gmail = { users: { messages: {
    list: async (args) => { lists.push(args); return { data: { messages: Array.from({ length: 23 }, (_, i) => ({ id: `m${i}` })), nextPageToken: 'next', resultSizeEstimate: 321 } }; },
    get: async ({ id }) => {
      inFlight++; maxInFlight = Math.max(maxInFlight, inFlight); await new Promise((r) => setTimeout(r, 1)); inFlight--;
      return { data: { id, threadId: 't', internalDate: '1', payload: { headers: [], body: {} } } };
    },
  } } };
  const client = new GmailAccountClient(account, {}, gmail, {}, {}, {}, {});
  const page = await client.readEmailPage('is:unread', 23, false, 'prior');
  assert.equal(page.emails.length, 23); assert.equal(page.nextPageToken, 'next'); assert.equal(page.resultSizeEstimate, 321);
  assert.deepEqual(lists, [{ userId: 'me', q: 'is:unread', maxResults: 23, pageToken: 'prior' }]);
  assert.ok(maxInFlight <= 10); assert.ok(page.emails.every((email) => email.accountId === 'work'));
});

test('empty Gmail page preserves continuation and does not make detail requests', async () => {
  const gmail = { users: { messages: {
    list: async () => ({ data: { messages: [], nextPageToken: 'empty-next', resultSizeEstimate: 2 } }),
    get: async () => { throw new Error('unexpected get'); },
  } } };
  const client = new GmailAccountClient({ id: 'work', email: 'work@example.com' }, {}, gmail, {}, {}, {}, {});
  assert.deepEqual(await client.readEmailPage('', 20, false), { emails: [], nextPageToken: 'empty-next', resultSizeEstimate: 2 });
});
