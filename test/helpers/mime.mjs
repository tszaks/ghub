import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

// Parse the base64url "raw" string Gmail would receive into headers and decoded parts.
export function parseRaw(raw) {
  const bytes = Buffer.from(raw.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
  return parsePart(bytes.toString('latin1'));
}

function parsePart(text) {
  const split = text.indexOf('\r\n\r\n');
  assert.ok(split > 0, 'headers end with a blank CRLF line');
  const headers = {};
  for (const line of text.slice(0, split).split('\r\n')) {
    const colon = line.indexOf(':');
    headers[line.slice(0, colon).toLowerCase()] = line.slice(colon + 1).trim();
  }
  const rawBody = text.slice(split + 4);
  const contentType = headers['content-type'] ?? 'text/plain';
  const boundary = /boundary="([^"]+)"/.exec(contentType)?.[1];
  if (boundary) {
    const chunks = rawBody.split(`--${boundary}`);
    assert.ok(chunks.at(-1).startsWith('--'), 'multipart has a closing boundary');
    const parts = chunks.slice(1, -1).map((chunk) => parsePart(chunk.replace(/^\r\n/, '').replace(/\r\n$/, '')));
    return { headers, parts };
  }
  const body =
    (headers['content-transfer-encoding'] ?? '').toLowerCase() === 'base64'
      ? Buffer.from(rawBody.replace(/\r\n/g, ''), 'base64')
      : Buffer.from(rawBody, 'latin1');
  return { headers, body };
}

export function assertCrlfOnly(buffer) {
  const text = buffer.toString('utf8');
  assert.ok(text.includes('\r\n'), 'body carries CRLF line endings');
  assert.equal((text.match(/(?<!\r)\n/g) ?? []).length, 0, 'body has no bare LF');
}

export const BODY = 'Hello,\n\nLine one.\nLine two.\n\nThanks';

export async function tempFile(name, contents) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'compose-test-'));
  const file = path.join(dir, name);
  await fs.writeFile(file, contents);
  return file;
}
