// Fake-backed tests for GmailAccountClient.downloadByPartId — the method
// that resolves a stable part_id to a fresh Gmail body.attachmentId and
// pulls the decoded bytes. Uses __forTests to inject a fake `gmail` client
// so the test runs offline with no real OAuth.
//
// These tests specifically guard the silent-failure gap that iteration 2
// of the finalize-branch review accidentally re-opened and iteration 4
// closed: a malformed upstream size clamped to 0 must NOT silently return
// arbitrary bytes. See commit ab27b81 for the full history.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { GmailAccountClient } from '../dist/gmail-client.js';

// Minimal fake Gmail API client. Each test builds one with the message and
// attachment shapes it needs; downloadByPartId only touches
// `users.messages.get` and `users.messages.attachments.get`.
const makeFakeGmail = ({ messagePayload, attachmentData }) => ({
  users: {
    messages: {
      get: async () => ({ data: { payload: messagePayload } }),
      attachments: {
        get: async () => ({ data: { data: attachmentData } }),
      },
    },
  },
});

const base64url = (bytes) =>
  Buffer.from(bytes)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');

const attachmentPart = ({ filename, attachmentId, size, partId = '1' }) => ({
  partId,
  filename,
  mimeType: 'application/octet-stream',
  body: { attachmentId, size },
  headers: [{ name: 'Content-Disposition', value: 'attachment' }],
});

const makeClient = (fakeGmail) =>
  GmailAccountClient.__forTests(
    { id: 'test', email: 'test@example.com', enabled: true, credentialPath: '', tokenPath: '' },
    { accountDir: '', credentialsPath: '', tokenPath: '' },
    fakeGmail
  );

const MAX = 25 * 1024 * 1024;

test('downloadByPartId: legitimate zero-byte attachment returns ok', async () => {
  // Gmail delivers a legitimate zero-byte attachment as `data: ''` on the
  // attachments.get response; the pre-iteration-4 `if (!raw)` check used to
  // treat this as "no data" and throw. The new `raw == null` check lets
  // '' through to the unconditional length check, where 0 === 0 passes.
  const client = makeClient(
    makeFakeGmail({
      messagePayload: {
        parts: [attachmentPart({ filename: 'empty.txt', attachmentId: 'ATT_EMPTY', size: 0 })],
      },
      attachmentData: '',
    })
  );

  const result = await client.downloadByPartId('msg1', '1', MAX);

  assert.equal(result.kind, 'ok');
  assert.equal(result.data.length, 0);
  assert.equal(result.metadata.size, 0);
  assert.equal(result.metadata.filename, 'empty.txt');
});

test('downloadByPartId: malformed upstream size clamped to 0 + non-empty payload throws', async () => {
  // This is the iteration-3-finding regression fence. A malformed Gmail
  // response reports size = -1 (or "NaN"), extractAttachmentParts clamps
  // it to 0, and without the unconditional length check downloadByPartId
  // used to silently return whatever bytes attachments.get gave back.
  // With the fix, the unconditional data.length !== metadata.size check
  // fires and surfaces the upstream corruption as a thrown error.
  //
  // If a future maintainer re-adds `metadata.size > 0 &&` to the length
  // check, this test is the line of defense that catches it.
  const client = makeClient(
    makeFakeGmail({
      messagePayload: {
        parts: [attachmentPart({ filename: 'bogus.bin', attachmentId: 'ATT_BOGUS', size: -1 })],
      },
      attachmentData: base64url(new Uint8Array(100).fill(42)),
    })
  );

  await assert.rejects(
    () => client.downloadByPartId('msg1', '1', MAX),
    (err) => {
      assert.match(err.message, /produced 100 bytes/);
      assert.match(err.message, /expected 0/);
      return true;
    }
  );
});

test('downloadByPartId: truncated base64 payload throws with expected-vs-actual', async () => {
  // Buffer.from(base64) silently drops invalid characters, so a truncated
  // attachment used to return a short-but-valid Buffer with no error. The
  // length check now fires on mismatch.
  const client = makeClient(
    makeFakeGmail({
      messagePayload: {
        parts: [attachmentPart({ filename: 'doc.pdf', attachmentId: 'ATT_DOC', size: 100 })],
      },
      attachmentData: base64url(new Uint8Array(50).fill(0xab)),
    })
  );

  await assert.rejects(
    () => client.downloadByPartId('msg1', '1', MAX),
    (err) => {
      assert.match(err.message, /produced 50 bytes/);
      assert.match(err.message, /expected 100/);
      return true;
    }
  );
});

test('downloadByPartId: missing data field throws "no data field"', async () => {
  // Gmail genuinely omitting the data field (as opposed to delivering an
  // empty string) should still throw — this pins the `raw == null` branch
  // distinct from the `raw === ''` branch.
  const client = makeClient(
    makeFakeGmail({
      messagePayload: {
        parts: [attachmentPart({ filename: 'doc.pdf', attachmentId: 'ATT_MISSING', size: 100 })],
      },
      attachmentData: undefined,
    })
  );

  await assert.rejects(
    () => client.downloadByPartId('msg1', '1', MAX),
    /no data field/
  );
});

test('downloadByPartId: unknown part_id returns kind=not_found', async () => {
  const client = makeClient(
    makeFakeGmail({
      messagePayload: {
        parts: [attachmentPart({ filename: 'a.pdf', attachmentId: 'ATT_A', size: 10 })],
      },
      attachmentData: base64url(new Uint8Array(10)),
    })
  );

  const result = await client.downloadByPartId('msg1', '99', MAX);
  assert.equal(result.kind, 'not_found');
});

test('downloadByPartId: oversized attachment returns kind=too_large before fetching bytes', async () => {
  // metadata.size reports 30 MB, maxBytes is 25 MB — the too_large branch
  // must fire from the metadata walk, before attachments.get is called.
  // We prove attachments.get is not called by having it throw if invoked.
  const client = makeClient({
    users: {
      messages: {
        get: async () => ({
          data: {
            payload: {
              parts: [
                attachmentPart({
                  filename: 'huge.bin',
                  attachmentId: 'ATT_HUGE',
                  size: 30 * 1024 * 1024,
                }),
              ],
            },
          },
        }),
        attachments: {
          get: async () => {
            throw new Error('attachments.get should not be called for oversized attachment');
          },
        },
      },
    },
  });

  const result = await client.downloadByPartId('msg1', '1', MAX);
  assert.equal(result.kind, 'too_large');
  assert.equal(result.metadata.filename, 'huge.bin');
  assert.equal(result.metadata.size, 30 * 1024 * 1024);
});
