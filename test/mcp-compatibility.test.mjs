import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

for (const args of [[], ['mcp']]) test(`MCP handshake and all 60 tools work with ${args.length ? 'explicit' : 'legacy bare'} invocation`, async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ghub-mcp-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const client = new Client({ name: 'offline-test', version: '1' });
  const transport = new StdioClientTransport({ command: process.execPath, args: ['dist/index.js', ...args], env: { ...process.env, GMAILMCPCONFIG_DIR: directory, MCP_TRANSPORT: 'stdio' }, stderr: 'pipe' });
  try {
    await client.connect(transport);
    const tools = await client.listTools(); assert.equal(tools.tools.length, 60);
    const result = await client.callTool({ name: 'list_accounts', arguments: {} });
    assert.deepEqual(result.structuredContent.accounts, []);
    assert.match(result.content[0].text, /No accounts configured/);
    const failure = await client.callTool({ name: 'get_labels', arguments: { account: 'missing' } });
    assert.equal(failure.isError, true); assert.match(failure.content[0].text, /Error executing/);
  } finally { await client.close(); }
});

test('existing SSE transport serves the same shared tools', async (t) => {
  const { createServer } = await import('node:net');
  const { spawn } = await import('node:child_process');
  const { once } = await import('node:events');
  const { SSEClientTransport } = await import('@modelcontextprotocol/sdk/client/sse.js');
  const reservation = createServer();
  await new Promise((resolve) => reservation.listen(0, '127.0.0.1', resolve));
  const port = reservation.address().port;
  await new Promise((resolve) => reservation.close(resolve));
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'ghub-sse-'));
  const child = spawn(process.execPath, ['dist/index.js'], { env: { ...process.env, GMAILMCPCONFIG_DIR: directory, MCP_TRANSPORT: 'sse', HOST: '127.0.0.1', PORT: String(port) }, stdio: ['ignore', 'pipe', 'pipe'] });
  t.after(async () => {
    if (child.exitCode === null) { child.kill('SIGTERM'); await once(child, 'exit'); }
    await fs.rm(directory, { recursive: true, force: true });
  });
  await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error('SSE startup timed out')), 10000);
    child.stderr.on('data', (chunk) => {
      if (chunk.toString().includes('Running on SSE')) { clearTimeout(timeout); resolve(); }
    });
    child.once('exit', (code) => { clearTimeout(timeout); reject(new Error(`SSE process exited ${code}`)); });
  });
  const response = await fetch(`http://127.0.0.1:${port}/`);
  assert.equal(await response.text(), 'ghub SSE server');
  const client = new Client({ name: 'offline-sse-test', version: '1' });
  try {
    await client.connect(new SSEClientTransport(new URL(`http://127.0.0.1:${port}/sse`)));
    assert.equal((await client.listTools()).tools.length, 60);
    const accounts = await client.callTool({ name: 'list_accounts', arguments: {} });
    assert.deepEqual(accounts.structuredContent.accounts, []);
  } finally { await client.close(); }
});
