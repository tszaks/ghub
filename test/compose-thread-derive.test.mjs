import test from 'node:test';
import assert from 'node:assert/strict';
import { threadingHeadersFromThread } from '../dist/gmail-client.js';

function message(id, labels, references) {
  const headers = [{ name: 'Message-ID', value: id }];
  if (references) headers.push({ name: 'References', value: references });
  return { labelIds: labels, payload: { headers } };
}

test('a reply is threaded under the last message of the thread', () => {
  const thread = [
    message('<a@mail.example.com>', ['INBOX']),
    message('<b@mail.example.com>', ['INBOX'], '<a@mail.example.com>'),
  ];
  assert.deepEqual(threadingHeadersFromThread(thread), {
    inReplyTo: '<b@mail.example.com>',
    references: '<a@mail.example.com> <b@mail.example.com>',
  });
});

test('drafts in the thread are skipped', () => {
  const thread = [
    message('<a@mail.example.com>', ['INBOX']),
    message('<draft@mail.example.com>', ['DRAFT'], '<a@mail.example.com>'),
  ];
  assert.deepEqual(threadingHeadersFromThread(thread), {
    inReplyTo: '<a@mail.example.com>',
    references: '<a@mail.example.com>',
  });
});

test('a thread with no usable Message-ID gives no headers', () => {
  assert.deepEqual(threadingHeadersFromThread([]), {});
  assert.deepEqual(threadingHeadersFromThread([{ labelIds: ['INBOX'], payload: { headers: [] } }]), {});
});
