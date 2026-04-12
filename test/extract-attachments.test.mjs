// Fixture tests for extractAttachmentParts — the MIME-tree walker that
// drives list_attachments, read_emails attachment summaries, and the
// download-by-part-id round trip.
//
// Run:
//   bun run build           # compile ts -> dist/
//   bun run test            # node --test test/*.test.mjs
//
// These tests import from `dist/gmail-client.js` so they exercise the
// actual compiled output the MCP server runs.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { extractAttachmentParts } from '../dist/gmail-client.js';

// Helper: build a `Schema$MessagePart`-shaped fixture with sensible defaults.
const part = (overrides = {}) => ({
  partId: '',
  mimeType: 'application/octet-stream',
  filename: '',
  headers: [],
  body: {},
  parts: undefined,
  ...overrides,
});

test('returns [] when payload is undefined', () => {
  assert.deepEqual(extractAttachmentParts(undefined), []);
});

test('returns [] when payload has no attachments and no parts', () => {
  const payload = part({ mimeType: 'text/plain', body: { data: 'aGk=', size: 2 } });
  assert.deepEqual(extractAttachmentParts(payload), []);
});

test('bare PDF at the root payload surfaces as partId "0"', () => {
  const payload = part({
    partId: '',
    filename: 'report.pdf',
    mimeType: 'application/pdf',
    body: { attachmentId: 'ATT_BARE', size: 12345 },
    headers: [
      { name: 'Content-Type', value: 'application/pdf; name="report.pdf"' },
      { name: 'Content-Disposition', value: 'attachment; filename="report.pdf"' },
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].partId, '0');
  assert.equal(result[0].filename, 'report.pdf');
  assert.equal(result[0].mimeType, 'application/pdf');
  assert.equal(result[0].size, 12345);
  assert.equal(result[0].attachmentId, 'ATT_BARE');
});

test('multipart/mixed: text body + one PDF attachment returns only the PDF', () => {
  const payload = part({
    mimeType: 'multipart/mixed',
    parts: [
      part({ partId: '0', mimeType: 'text/plain', body: { data: 'aGVsbG8=', size: 5 } }),
      part({
        partId: '1',
        filename: 'invoice.pdf',
        mimeType: 'application/pdf',
        body: { attachmentId: 'ATT_INV', size: 2048 },
        headers: [
          { name: 'Content-Disposition', value: 'attachment; filename="invoice.pdf"' },
        ],
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].partId, '1');
  assert.equal(result[0].filename, 'invoice.pdf');
  assert.equal(result[0].attachmentId, 'ATT_INV');
});

test('Content-Disposition: inline parts are filtered out', () => {
  const payload = part({
    mimeType: 'multipart/related',
    parts: [
      part({
        partId: '0',
        filename: 'logo.png',
        mimeType: 'image/png',
        body: { attachmentId: 'ATT_LOGO', size: 1500 },
        headers: [
          { name: 'Content-Disposition', value: 'inline; filename="logo.png"' },
          { name: 'Content-ID', value: '<logo@company.example>' },
        ],
      }),
    ],
  });
  assert.deepEqual(extractAttachmentParts(payload), []);
});

test('inline-disposition check is case-insensitive', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: 'pixel.gif',
        mimeType: 'image/gif',
        body: { attachmentId: 'ATT_PIXEL', size: 43 },
        headers: [{ name: 'Content-Disposition', value: 'INLINE' }],
      }),
    ],
  });
  assert.deepEqual(extractAttachmentParts(payload), []);
});

test('attachment with explicit Content-Disposition: attachment is included', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: 'receipt.pdf',
        mimeType: 'application/pdf',
        body: { attachmentId: 'ATT_RECEIPT', size: 8192 },
        headers: [
          { name: 'Content-Disposition', value: 'attachment; filename="receipt.pdf"' },
        ],
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].partId, '0');
});

test('attachment with NO Content-Disposition header at all is included (forwards / simple clients)', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: 'forward.txt',
        mimeType: 'text/plain',
        body: { attachmentId: 'ATT_FWD', size: 128 },
        headers: [
          { name: 'Content-Type', value: 'text/plain; name="forward.txt"' },
        ],
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].filename, 'forward.txt');
});

test('nested multipart/alternative inside multipart/mixed — deep attachment is found', () => {
  const payload = part({
    mimeType: 'multipart/mixed',
    parts: [
      part({
        partId: '0',
        mimeType: 'multipart/alternative',
        parts: [
          part({ partId: '0.0', mimeType: 'text/plain', body: { data: 'aGVsbG8=', size: 5 } }),
          part({ partId: '0.1', mimeType: 'text/html', body: { data: 'PHA+aGVsbG88L3A+', size: 12 } }),
        ],
      }),
      part({
        partId: '1',
        filename: 'deep.pdf',
        mimeType: 'application/pdf',
        body: { attachmentId: 'ATT_DEEP', size: 4096 },
        headers: [{ name: 'Content-Disposition', value: 'attachment' }],
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].partId, '1');
  assert.equal(result[0].attachmentId, 'ATT_DEEP');
});

test('part with filename but no body.attachmentId is skipped (body part, not attachment)', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: '',
        mimeType: 'text/plain',
        body: { data: 'aGVsbG8=', size: 5 },
      }),
    ],
  });
  assert.deepEqual(extractAttachmentParts(payload), []);
});

test('mimeType defaults to application/octet-stream when missing', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: 'blob',
        mimeType: undefined,
        body: { attachmentId: 'ATT_BLOB', size: 64 },
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].mimeType, 'application/octet-stream');
});

test('size defaults to 0 when body.size is missing', () => {
  const payload = part({
    parts: [
      part({
        partId: '0',
        filename: 'mystery.bin',
        body: { attachmentId: 'ATT_MYSTERY' },
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 1);
  assert.equal(result[0].size, 0);
});

test('two real attachments in one message are both returned in order', () => {
  const payload = part({
    mimeType: 'multipart/mixed',
    parts: [
      part({
        partId: '0',
        filename: 'first.pdf',
        mimeType: 'application/pdf',
        body: { attachmentId: 'ATT_FIRST', size: 100 },
        headers: [{ name: 'Content-Disposition', value: 'attachment' }],
      }),
      part({
        partId: '1',
        filename: 'second.png',
        mimeType: 'image/png',
        body: { attachmentId: 'ATT_SECOND', size: 200 },
        headers: [{ name: 'Content-Disposition', value: 'attachment' }],
      }),
    ],
  });
  const result = extractAttachmentParts(payload);
  assert.equal(result.length, 2);
  assert.equal(result[0].partId, '0');
  assert.equal(result[0].filename, 'first.pdf');
  assert.equal(result[1].partId, '1');
  assert.equal(result[1].filename, 'second.png');
});
