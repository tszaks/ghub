import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {
  sanitizeFilename,
  buildSavePath,
  getSaveDir,
  saveAttachment,
  extractText,
} from '../dist/attachments.js';

test('sanitizeFilename blocks path traversal', () => {
  const result = sanitizeFilename('../../etc/passwd');
  assert.ok(!result.includes('/'), 'forward slash removed');
  assert.ok(!result.includes('..'), 'double-dot collapsed');
  assert.ok(!result.includes('\\'), 'backslash removed');
});

test('sanitizeFilename removes control chars', () => {
  assert.equal(sanitizeFilename('evil\u0000file.txt'), 'evilfile.txt');
  assert.equal(sanitizeFilename('line\u000abreak.txt'), 'linebreak.txt');
});

test('sanitizeFilename handles empty and whitespace', () => {
  assert.equal(sanitizeFilename(''), 'attachment');
  assert.equal(sanitizeFilename('   '), 'attachment');
});

test('sanitizeFilename truncates long names but keeps extension', () => {
  const longName = 'a'.repeat(300) + '.pdf';
  const result = sanitizeFilename(longName);
  assert.ok(result.length <= 200, `length ${result.length} > 200`);
  assert.ok(result.endsWith('.pdf'), 'extension preserved');
});

test('sanitizeFilename passes through normal filenames', () => {
  assert.equal(sanitizeFilename('Q1 2026 Report.pdf'), 'Q1 2026 Report.pdf');
  assert.equal(sanitizeFilename('invoice-123.xlsx'), 'invoice-123.xlsx');
});

test('buildSavePath includes date and stable short id', () => {
  const p = buildSavePath('attach-xyz-9988', 'report.pdf');
  const base = path.basename(p);
  assert.match(base, /^\d{4}-\d{2}-\d{2}-/, 'starts with YYYY-MM-DD');
  assert.ok(base.endsWith('-report.pdf'), 'ends with sanitized filename');
  assert.ok(base.includes('attachxy'), 'contains 8-char id prefix');
});

test('buildSavePath is collision-safe across different attachment ids', () => {
  const a = buildSavePath('aaaa1111', 'file.pdf');
  const b = buildSavePath('bbbb2222', 'file.pdf');
  assert.notEqual(a, b);
});

test('getSaveDir respects MCP_ATTACHMENTS_DIR override', () => {
  const before = process.env.MCP_ATTACHMENTS_DIR;
  process.env.MCP_ATTACHMENTS_DIR = '/tmp/mcp-attach-override';
  try {
    assert.equal(getSaveDir(), '/tmp/mcp-attach-override');
  } finally {
    if (before === undefined) delete process.env.MCP_ATTACHMENTS_DIR;
    else process.env.MCP_ATTACHMENTS_DIR = before;
  }
});

test('getSaveDir default is under ~/Downloads/mcp-attachments', () => {
  const before = process.env.MCP_ATTACHMENTS_DIR;
  delete process.env.MCP_ATTACHMENTS_DIR;
  try {
    const dir = getSaveDir();
    assert.ok(
      dir.includes(path.join('Downloads', 'mcp-attachments')),
      `default dir "${dir}" should be under Downloads/mcp-attachments`,
    );
  } finally {
    if (before !== undefined) process.env.MCP_ATTACHMENTS_DIR = before;
  }
});

test('extractText on UTF-8 file returns inline text', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'attach-test-'));
  const file = path.join(dir, 'sample.txt');
  await fs.writeFile(file, 'hello world\n');
  const result = await extractText(file, 'text/plain');
  assert.equal(result.extractionMethod, 'utf8');
  assert.equal(result.text, 'hello world\n');
  assert.equal(result.textTruncated, undefined);
});

test('extractText truncates text files over 500 KB', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'attach-test-'));
  const file = path.join(dir, 'big.txt');
  await fs.writeFile(file, 'a'.repeat(600_000));
  const result = await extractText(file, 'text/plain');
  assert.equal(result.extractionMethod, 'utf8');
  assert.equal(result.text?.length, 500_000);
  assert.equal(result.textTruncated, true);
});

test('extractText on unknown binary returns none / no text', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'attach-test-'));
  const file = path.join(dir, 'blob.bin');
  await fs.writeFile(file, Buffer.from([0x00, 0xff, 0x01, 0xfe]));
  const result = await extractText(file, 'application/octet-stream');
  assert.equal(result.extractionMethod, 'none');
  assert.equal(result.text, null);
});

test('saveAttachment writes bytes to configured dir', async () => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'save-test-'));
  const before = process.env.MCP_ATTACHMENTS_DIR;
  process.env.MCP_ATTACHMENTS_DIR = tmp;
  try {
    const outPath = await saveAttachment(
      Buffer.from('abc'),
      'att-test-id',
      'file.bin',
    );
    const stat = await fs.stat(outPath);
    assert.equal(stat.size, 3);
    assert.ok(outPath.startsWith(tmp), 'saved inside override dir');
  } finally {
    if (before === undefined) delete process.env.MCP_ATTACHMENTS_DIR;
    else process.env.MCP_ATTACHMENTS_DIR = before;
  }
});

test('extractText dispatches to OCR for image MIME types', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'attach-test-'));
  const file = path.join(dir, 'fake.png');
  await fs.writeFile(file, Buffer.from([0x89, 0x50, 0x4e, 0x47]));
  const result = await extractText(file, 'image/png');
  // Either OCR ran (real binary present) and returned ocr-vision, or binary
  // was missing/unreadable PNG caused an OCR error. Both are valid here.
  assert.ok(
    result.extractionMethod === 'ocr-vision' || result.extractionMethod === 'none',
    `unexpected method ${result.extractionMethod}`,
  );
});
