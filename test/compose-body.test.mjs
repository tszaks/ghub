import test from 'node:test';
import assert from 'node:assert/strict';
import { createRawEmailMessage } from '../dist/gmail-client.js';
import { BODY, assertCrlfOnly, parseRaw, tempFile } from './helpers/mime.mjs';

test('plain body is sent with CRLF line endings', async () => {
  const msg = parseRaw(await createRawEmailMessage({ to: 'someone@example.com', subject: 'Hi', body: BODY }));
  assertCrlfOnly(msg.body);
});

test('a body that already uses CRLF is not doubled', async () => {
  const msg = parseRaw(
    await createRawEmailMessage({ to: 'someone@example.com', subject: 'Hi', body: 'one\r\ntwo\rthree\n' })
  );
  assert.equal(msg.body.toString('utf8'), 'one\r\ntwo\r\nthree\r\n');
});

test('body is sent with CRLF line endings when attachments are present', async () => {
  const file = await tempFile('report.pdf', Buffer.from('%PDF-1.4 fake'));
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Hi',
      body: BODY,
      attachments: [{ path: file }],
    })
  );
  assert.equal(msg.parts.length, 2);
  assert.match(msg.parts[0].headers['content-type'], /^text\/plain/);
  assertCrlfOnly(msg.parts[0].body);
});

test('every attachment is present with its filename and exact bytes', async () => {
  const pdf = Buffer.from([0x25, 0x50, 0x44, 0x46, 0x00, 0xff, 0x0a, 0x0d]);
  const first = await tempFile('invoice.pdf', pdf);
  const second = await tempFile('notes.txt', 'a\nb\n');
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Hi',
      body: BODY,
      attachments: [{ path: first }, { path: second }],
    })
  );
  const files = msg.parts.slice(1);
  assert.equal(files.length, 2);
  assert.match(files[0].headers['content-disposition'], /filename="invoice\.pdf"/);
  assert.match(files[0].headers['content-type'], /^application\/pdf/);
  assert.deepEqual(files[0].body, pdf);
  assert.match(files[1].headers['content-disposition'], /filename="notes\.txt"/);
  assert.deepEqual(files[1].body, Buffer.from('a\nb\n'), 'attachment bytes are not altered');
});

test('a non-ASCII attachment filename is RFC 2047 encoded', async () => {
  const file = await tempFile('notes.txt', 'x');
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Hi',
      body: 'b',
      attachments: [{ path: file, filename: 'Σημειώσεις.txt' }],
    })
  );
  const encoded = `=?UTF-8?B?${Buffer.from('Σημειώσεις.txt').toString('base64')}?=`;
  assert.ok(msg.parts[1].headers['content-disposition'].includes(`filename="${encoded}"`));
  assert.ok(msg.parts[1].headers['content-type'].includes(`name="${encoded}"`));
});
