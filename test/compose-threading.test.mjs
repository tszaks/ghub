import test from 'node:test';
import assert from 'node:assert/strict';
import { createRawEmailMessage } from '../dist/gmail-client.js';
import { parseRaw, tempFile } from './helpers/mime.mjs';

const PARENT = '<parent-1@mail.example.com>';
const CHAIN = '<root-0@mail.example.com> <parent-1@mail.example.com>';

test('threading headers are written on a plain reply', async () => {
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Re: Hi',
      body: 'b',
      inReplyTo: PARENT,
      references: CHAIN,
    })
  );
  assert.equal(msg.headers['in-reply-to'], PARENT);
  assert.equal(msg.headers['references'], CHAIN);
});

test('threading headers are kept on a reply with attachments', async () => {
  const file = await tempFile('a.pdf', 'x');
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Re: Hi',
      body: 'b',
      attachments: [{ path: file }],
      inReplyTo: PARENT,
      references: CHAIN,
    })
  );
  assert.equal(msg.headers['in-reply-to'], PARENT);
  assert.equal(msg.headers['references'], CHAIN);
});

test('References falls back to In-Reply-To when only that is given', async () => {
  const msg = parseRaw(
    await createRawEmailMessage({ to: 'someone@example.com', subject: 'Re: Hi', body: 'b', inReplyTo: PARENT })
  );
  assert.equal(msg.headers['references'], PARENT);
});

test('a line break in a threading value cannot add a header', async () => {
  const msg = parseRaw(
    await createRawEmailMessage({
      to: 'someone@example.com',
      subject: 'Re: Hi',
      body: 'b',
      inReplyTo: `${PARENT}\r\nBcc: other@example.com`,
    })
  );
  assert.equal(msg.headers['bcc'], undefined);
});
