import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import path from 'node:path';
import { createRawEmailMessage } from '../dist/gmail-client.js';
import { assertKnownOutgoingEmailArgs, valueToAttachmentArray } from '../dist/outgoing-email.js';
import { parseRaw, tempFile } from './helpers/mime.mjs';

test('attachments given as plain path strings are kept', () => {
  assert.deepEqual(valueToAttachmentArray(['/tmp/a.pdf']), [{ path: '/tmp/a.pdf' }]);
});

test('attachments given as a JSON-encoded array are kept', () => {
  assert.deepEqual(valueToAttachmentArray(JSON.stringify([{ path: '/tmp/a.pdf', filename: 'b.pdf' }])), [
    { path: '/tmp/a.pdf', filename: 'b.pdf' },
  ]);
});

test('an attachment entry without a path is an error, not silently dropped', () => {
  assert.throws(() => valueToAttachmentArray([{ file_path: '/tmp/a.pdf' }]), /path/);
  assert.throws(() => valueToAttachmentArray([42]), /attachment/i);
  assert.throws(() => valueToAttachmentArray(42), /attachments/);
});

test('no attachments is still fine', () => {
  assert.deepEqual(valueToAttachmentArray(undefined), []);
  assert.deepEqual(valueToAttachmentArray(null), []);
  assert.deepEqual(valueToAttachmentArray([]), []);
});

test('a ~ path is expanded to the home directory', async () => {
  const file = await tempFile('home.txt', 'x');
  const home = path.dirname(file);
  const saved = process.env.HOME;
  process.env.HOME = home;
  try {
    const msg = parseRaw(
      await createRawEmailMessage({
        to: 'someone@example.com',
        subject: 'Hi',
        body: 'b',
        attachments: [{ path: '~/home.txt' }],
      })
    );
    assert.equal(msg.parts.length, 2);
    assert.deepEqual(msg.parts[1].body, Buffer.from('x'));
  } finally {
    process.env.HOME = saved;
  }
});

test('a path that cannot be attached fails the whole message with a clear error', async () => {
  const build = (p) =>
    createRawEmailMessage({ to: 'someone@example.com', subject: 'Hi', body: 'b', attachments: [{ path: p }] });
  await assert.rejects(build('/nonexistent/dir/missing.pdf'), /Cannot attach "\/nonexistent\/dir\/missing\.pdf"/);
  await assert.rejects(build('relative/file.pdf'), /absolute/);
  await assert.rejects(build(os.tmpdir()), /not a regular file/);
});

test('an argument the tool does not define is rejected instead of ignored', () => {
  assert.throws(
    () => assertKnownOutgoingEmailArgs({ account: 'a', to: 'b', subject: 'c', body: 'd', file_path: '/tmp/a.pdf' }),
    /Unknown argument\(s\): file_path/
  );
  assert.doesNotThrow(() =>
    assertKnownOutgoingEmailArgs({ account: 'a', to: 'b', subject: 'c', body: 'd', attachments: [], thread_id: 't' })
  );
});
