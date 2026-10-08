import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { GmailMultiInboxServer } from '../dist/app.js';

const account = { id: 'work', email: 'work@example.com', enabled: true };
const otherAccount = { id: 'personal', email: 'personal@example.com', enabled: true };
const file = {
  id: 'file-1', name: 'Project notes', mimeType: 'text/plain', owners: ['work@example.com'],
  accountId: account.id, accountEmail: account.email, modifiedTime: '2026-10-08T12:00:00Z',
};
const calendar = {
  id: 'primary', summary: 'Work', primary: true, accountId: account.id, accountEmail: account.email,
};
const event = {
  id: 'event-1', calendarId: 'primary', summary: 'Planning',
  start: { dateTime: '2026-10-08T14:00:00Z' }, end: { dateTime: '2026-10-08T15:00:00Z' },
  attendees: [{ email: 'colleague@example.com', responseStatus: 'accepted' }],
  accountId: account.id, accountEmail: account.email,
};
const spreadsheet = {
  id: 'sheet-1', title: 'Planning', url: 'https://docs.google.com/spreadsheets/d/sheet-1',
  sheets: [{ sheetId: 0, title: 'Sheet1', index: 0, rowCount: 100, columnCount: 20 }],
};
const document = {
  documentId: 'doc-1', title: 'Notes', body: `${'Full content '.repeat(200)}tail`,
  url: 'https://docs.google.com/document/d/doc-1',
};
const sheetValues = [[0, false, 'literal **Markdown**', null], [12.5, true, 'next', '']];

function fakeClient(overrides = {}) {
  return {
    searchDriveFiles: async () => [file],
    listDriveFiles: async () => ({ files: [file], nextPageToken: 'page-2' }),
    getDriveFile: async () => file,
    getDriveFileContent: async () => ({ bytes: Buffer.from('Extracted file body'), contentType: 'text/plain', filename: 'notes.txt' }),
    uploadDriveFile: async () => file,
    createDriveFolder: async () => ({ ...file, id: 'folder-1', mimeType: 'application/vnd.google-apps.folder' }),
    updateDriveFile: async () => file,
    trashDriveFile: async () => undefined,
    shareDriveFile: async () => ({ permissionId: 'permission-1' }),
    getSheetsMetadata: async () => spreadsheet,
    readSheetValues: async () => ({ range: 'Sheet1!A1:D2', values: sheetValues }),
    writeSheetValues: async () => ({ updatedRange: 'Sheet1!A1:D2', updatedRows: 2, updatedCells: 8 }),
    appendSheetValues: async () => ({ updatedRange: 'Sheet1!A3:D4', updatedRows: 2 }),
    createSpreadsheet: async () => ({ id: spreadsheet.id, url: spreadsheet.url }),
    addSheetTab: async () => ({ sheetId: 1, title: 'Archive' }),
    renameSheetTab: async () => undefined,
    deleteSheetTab: async () => undefined,
    formatCells: async () => undefined,
    addChart: async () => ({ chartId: 45 }),
    insertDimension: async () => undefined,
    deleteDimension: async () => undefined,
    getDocument: async () => document,
    createDocument: async () => ({ documentId: document.documentId, url: document.url }),
    appendToDocument: async () => undefined,
    replaceInDocument: async () => ({ occurrencesChanged: 3 }),
    insertTableInDocument: async () => undefined,
    applyDocHeadingStyle: async () => undefined,
    listCalendars: async () => [calendar],
    listCalendarEvents: async () => [event],
    getCalendarEvent: async () => event,
    createCalendarEvent: async () => event,
    updateCalendarEvent: async () => event,
    deleteCalendarEvent: async () => undefined,
    ...overrides,
  };
}

function createApp(client = fakeClient(), accounts = [account]) {
  const app = new GmailMultiInboxServer();
  app.loadConfig = async () => ({ accounts, defaultAccount: account.id });
  app.getClientForAccount = async () => client;
  return app;
}

const args = {
  account: 'work', file_id: 'file-1', query: 'Project', name: 'Archive', local_path: '/unused/fixture',
  role: 'reader', type: 'user', email: 'colleague@example.com', spreadsheet_id: 'sheet-1',
  range: 'Sheet1!A1:D2', values: sheetValues, title: 'Archive', current_title: 'Sheet1',
  new_title: 'Renamed', sheet_title: 'Sheet1', chart_type: 'BAR', data_range: 'A1:D2',
  dimension: 'ROWS', start_index: 0, count: 2, document_id: 'doc-1', text: 'Append this',
  find: 'old', replace_with: 'new', rows: 2, columns: 3, end_index: 10, style: 'HEADING_1',
  calendar_id: 'primary', event_id: 'event-1', summary: 'Planning',
  start_date_time: event.start.dateTime, end_date_time: event.end.dateTime,
};

const cases = [
  ['search_drive_files', (data) => assert.deepEqual(data.files, [file])],
  ['list_drive_files', (data) => {
    assert.deepEqual(data.files, [file]);
    assert.equal(data.nextPageToken, 'page-2');
    assert.equal(data.nextPageTokens, undefined);
    assert.equal(data.pagination, undefined);
  }],
  ['get_drive_file', (data) => assert.deepEqual(data.file, file)],
  ['upload_drive_file', (data) => assert.equal(data.fileId, 'file-1')],
  ['create_drive_folder', (data) => assert.equal(data.folderId, 'folder-1')],
  ['update_drive_file', (data) => assert.equal(data.updated, true)],
  ['trash_drive_file', (data) => assert.equal(data.trashed, true)],
  ['share_drive_file', (data) => assert.equal(data.permissionId, 'permission-1')],
  ['sheets_get', (data) => assert.deepEqual(data.sheets, spreadsheet.sheets)],
  ['sheets_read', (data) => assert.deepEqual(data.values, sheetValues)],
  ['sheets_write', (data) => { assert.equal(data.updatedCells, 8); assert.equal(data.updatedRows, 2); }],
  ['sheets_append', (data) => { assert.equal(data.updatedRange, 'Sheet1!A3:D4'); assert.equal(data.updatedRows, 2); }],
  ['sheets_create', (data) => assert.equal(data.spreadsheetId, 'sheet-1')],
  ['sheets_add_tab', (data) => assert.equal(data.sheetId, 1)],
  ['sheets_rename_tab', (data) => { assert.equal(data.previousTitle, 'Sheet1'); assert.equal(data.title, 'Renamed'); }],
  ['sheets_delete_tab', (data) => assert.equal(data.deleted, true)],
  ['sheets_format', (data) => { assert.equal(data.range, args.range); assert.equal(data.updated, true); }],
  ['sheets_add_chart', (data) => assert.equal(data.chartId, 45)],
  ['sheets_insert_dimension', (data) => { assert.equal(data.count, 2); assert.equal(data.updated, true); }],
  ['sheets_delete_dimension', (data) => { assert.equal(data.count, 2); assert.equal(data.deleted, true); }],
  ['docs_get', (data) => assert.equal(data.body, document.body)],
  ['docs_create', (data) => assert.equal(data.documentId, 'doc-1')],
  ['docs_append', (data) => assert.equal(data.updated, true)],
  ['docs_replace_text', (data) => assert.equal(data.occurrencesChanged, 3)],
  ['docs_insert_table', (data) => { assert.equal(data.rows, 2); assert.equal(data.columns, 3); }],
  ['docs_apply_style', (data) => { assert.equal(data.startIndex, 0); assert.equal(data.endIndex, 10); assert.equal(data.style, 'HEADING_1'); }],
  ['list_calendars', (data) => assert.deepEqual(data.calendars, [calendar])],
  ['list_events', (data) => assert.deepEqual(data.events, [event])],
  ['get_event', (data) => assert.deepEqual(data.event, event)],
  ['create_event', (data) => { assert.equal(data.eventId, 'event-1'); assert.equal(data.created, true); }],
  ['update_event', (data) => { assert.equal(data.eventId, 'event-1'); assert.equal(data.updated, true); }],
  ['delete_event', (data) => { assert.equal(data.eventId, 'event-1'); assert.equal(data.deleted, true); }],
];

for (const [name, verify] of cases) {
  test(`${name} exposes native structured output and retains MCP text`, async () => {
    const result = await createApp().callTool(name, args);
    assert.ok(result.structuredContent, name);
    assert.equal(result.content[0].type, 'text');
    assert.ok(result.content[0].text.length > 0);
    assert.equal(result.isError, undefined);
    if (['search_drive_files', 'list_drive_files', 'list_calendars', 'list_events'].includes(name)) {
      assert.deepEqual(result.structuredContent.accounts, ['work']);
      assert.deepEqual(result.structuredContent.successfulAccounts, ['work']);
      assert.deepEqual(result.structuredContent.errors, []);
      assert.equal(result.structuredContent.count, 1);
    } else {
      assert.equal(result.structuredContent.account, 'work');
    }
    verify(result.structuredContent);
  });
}

test('workspace fixtures cover every registered Workspace command', () => {
  const workspaceNames = createApp().listTools().map((tool) => tool.name)
    .filter((name) => /drive|^sheets_|^docs_|calendar|_event$|_events$/.test(name));
  assert.deepEqual(workspaceNames.sort(), [...cases.map(([name]) => name), 'get_drive_file_content'].sort());
});

test('get_drive_file_content includes extracted text, metadata, and saved file', async (t) => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'ghub-workspace-output-'));
  const previous = process.env.MCP_ATTACHMENTS_DIR;
  process.env.MCP_ATTACHMENTS_DIR = directory;
  t.after(async () => {
    if (previous === undefined) delete process.env.MCP_ATTACHMENTS_DIR;
    else process.env.MCP_ATTACHMENTS_DIR = previous;
    await rm(directory, { recursive: true, force: true });
  });
  const result = await createApp().callTool('get_drive_file_content', args);
  const data = result.structuredContent;
  assert.equal(data.account, 'work');
  assert.equal(data.fileId, 'file-1');
  assert.equal(data.text, 'Extracted file body');
  assert.equal(data.contentType, 'text/plain');
  assert.equal(data.filename, 'notes.txt');
  assert.equal(data.extractionMethod, 'utf8');
  assert.equal(data.sizeBytes, Buffer.byteLength('Extracted file body'));
  assert.equal(await readFile(data.savedPath, 'utf8'), data.text);
});

test('get_drive_file_content exposes only a generic structured extraction error', async (t) => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'ghub-workspace-error-'));
  const previous = process.env.MCP_ATTACHMENTS_DIR;
  process.env.MCP_ATTACHMENTS_DIR = directory;
  t.after(async () => {
    if (previous === undefined) delete process.env.MCP_ATTACHMENTS_DIR;
    else process.env.MCP_ATTACHMENTS_DIR = previous;
    await rm(directory, { recursive: true, force: true });
  });
  const app = createApp(fakeClient({ getDriveFileContent: async () => ({
    bytes: Buffer.from('invalid document'),
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    filename: 'invalid.docx',
  }) }));
  const result = await app.callTool('get_drive_file_content', args);
  assert.equal(result.structuredContent.extractionError, 'Text extraction failed or was unavailable for this file.');
  assert.equal(result.structuredContent.text, null);
  assert.match(result.content[0].text, /Extraction Error.*DOCX parse failed/);
  assert.doesNotMatch(result.structuredContent.extractionError, /DOCX parse failed|zip|invalid document/);
});

const aggregateCases = [
  ['search_drive_files', 'searchDriveFiles', 'files'],
  ['list_drive_files', 'listDriveFiles', 'files'],
  ['list_calendars', 'listCalendars', 'calendars'],
  ['list_events', 'listCalendarEvents', 'events'],
];

for (const [name, method, key] of aggregateCases) {
  test(`${name} preserves successes and reports safe per-account errors`, async () => {
    const app = createApp(fakeClient(), [account, otherAccount]);
    app.getClientForAccount = async (selected) => selected.id === 'work' ? fakeClient() : fakeClient({
      [method]: async () => { throw new Error('Bearer secret-token; refresh_token=secret; upstream response body'); },
    });
    const result = await app.callTool(name, { query: 'Project' });
    const data = result.structuredContent;
    assert.equal(data[key].length, 1);
    assert.equal(data.count, 1);
    assert.deepEqual(data.accounts, ['work', 'personal']);
    assert.deepEqual(data.successfulAccounts, ['work']);
    assert.equal(data.errors.length, 1);
    assert.equal(data.errors[0].account, 'personal');
    assert.equal(data.errors[0].code, 'INTERNAL_ERROR');
    assert.equal(data.errors[0].exitCode, 1);
    assert.equal(data.errors[0].retryable, false);
    assert.equal(typeof data.errors[0].message, 'string');
    assert.equal(typeof data.errors[0].recovery, 'string');
    assert.doesNotMatch(JSON.stringify(data), /secret|Bearer|upstream response body/);
    // Existing MCP presentation is retained; the CLI consumes only structuredContent.
    assert.match(result.content[0].text, /Account Errors/);
  });

  test(`${name} exposes all-account failures rather than a silent empty result`, async () => {
    const app = createApp(fakeClient({ [method]: async () => { throw new Error('request failed'); } }));
    const { structuredContent: data } = await app.callTool(name, { query: 'Project' });
    assert.deepEqual(data[key], []);
    assert.equal(data.count, 0);
    assert.equal(data.errors.length, 1);
    assert.deepEqual(data.successfulAccounts, []);
    assert.equal(data.errors[0].account, 'work');
  });

  test(`${name} exposes successful empty collections`, async () => {
    const app = createApp(fakeClient({ [method]: async () => method === 'listDriveFiles' ? { files: [] } : [] }));
    const { structuredContent: data } = await app.callTool(name, { query: 'Project' });
    assert.deepEqual(data[key], []);
    assert.equal(data.count, 0);
    assert.deepEqual(data.errors, []);
    assert.deepEqual(data.successfulAccounts, ['work']);
  });

  for (const [status, code, exitCode, retryable] of [
    [401, 'AUTH_REQUIRED', 3, false],
    [429, 'UPSTREAM_UNAVAILABLE', 5, true],
  ]) {
    test(`${name} classifies ${status} failures without leaking provider details`, async () => {
      const app = createApp(fakeClient({ [method]: async () => {
        throw Object.assign(new Error('Bearer secret-token; provider-body'), {
          response: { status, data: { refresh_token: 'secret' } },
        });
      } }));
      const { structuredContent: data } = await app.callTool(name, { account: 'work', query: 'Project' });
      assert.deepEqual(data.successfulAccounts, []);
      assert.equal(data.errors[0].account, 'work');
      assert.equal(data.errors[0].code, code);
      assert.equal(data.errors[0].exitCode, exitCode);
      assert.equal(data.errors[0].retryable, retryable);
      assert.ok(data.errors[0].message);
      assert.ok(data.errors[0].recovery);
      assert.doesNotMatch(JSON.stringify(data), /Bearer|secret|provider-body|refresh_token/);
    });
  }

  test(`${name} distinguishes an empty successful account from a failed account`, async () => {
    const app = createApp(fakeClient(), [account, otherAccount]);
    app.getClientForAccount = async (selected) => fakeClient({ [method]: async () => {
      if (selected.id === 'personal') throw Object.assign(new Error('rate limit'), { response: { status: 429 } });
      return method === 'listDriveFiles' ? { files: [] } : [];
    } });
    const { structuredContent: data } = await app.callTool(name, { query: 'Project' });
    assert.deepEqual(data[key], []);
    assert.equal(data.count, 0);
    assert.deepEqual(data.successfulAccounts, ['work']);
    assert.equal(data.errors.length, 1);
    assert.equal(data.errors[0].account, 'personal');
    assert.equal(data.errors[0].code, 'UPSTREAM_UNAVAILABLE');
  });
}

test('list_drive_files omits unsafe cross-account tokens when merged results discard fetched records', async () => {
  const app = createApp(fakeClient(), [account, otherAccount]);
  app.getClientForAccount = async (selected) => fakeClient({
    listDriveFiles: async () => ({
      files: [{ ...file, id: `${selected.id}-file`, accountId: selected.id, accountEmail: selected.email }],
      nextPageToken: `${selected.id}-next`,
    }),
  });
  const { structuredContent: data } = await app.callTool('list_drive_files', { max_results: 1 });
  assert.equal(data.files.length, 1);
  assert.equal(data.totalFound, 2);
  assert.equal('nextPageToken' in data, false);
  assert.equal('nextPageTokens' in data, false);
  assert.deepEqual(data.pagination, {
    mode: 'per_account',
    recovery: 'Repeat with an explicit account to page all results.',
  });
  assert.doesNotMatch(JSON.stringify(data), /work-next|personal-next/);
});

test('list_drive_files requires explicit account for pagination even with one configured account', async () => {
  const { structuredContent: data } = await createApp().callTool('list_drive_files', {});
  assert.equal('nextPageToken' in data, false);
  assert.equal('nextPageTokens' in data, false);
  assert.equal(data.pagination.mode, 'per_account');
});

test('list_drive_files rejects a page token without an explicit account before any API calls', async () => {
  let called = false;
  const app = createApp(fakeClient({ listDriveFiles: async () => { called = true; return { files: [] }; } }));
  await assert.rejects(app.callTool('list_drive_files', { page_token: 'page-2' }), /page_token requires an explicit account/);
  await assert.rejects(app.callTool('list_drive_files', { account: '  ', page_token: 'page-2' }), /page_token requires an explicit account/);
  assert.equal(called, false);
});

test('list_drive_files returns and consumes explicit-account page tokens without skipping files', async () => {
  const first = { ...file, id: 'first' };
  const second = { ...file, id: 'second' };
  const tokens = [];
  const app = createApp(fakeClient({ listDriveFiles: async ({ pageToken }) => {
    tokens.push(pageToken);
    return pageToken === 'page-2' ? { files: [second] } : { files: [first], nextPageToken: 'page-2' };
  } }));
  const { structuredContent: page1 } = await app.callTool('list_drive_files', { account: 'work', max_results: 1 });
  const { structuredContent: page2 } = await app.callTool('list_drive_files', {
    account: 'work', max_results: 1, page_token: page1.nextPageToken,
  });
  assert.deepEqual([...page1.files, ...page2.files].map((item) => item.id), ['first', 'second']);
  assert.deepEqual(tokens, [undefined, 'page-2']);
  assert.equal(page2.nextPageToken, undefined);
  assert.equal(page1.pagination, undefined);
});

test('docs_get keeps full structured body even when the MCP preview is shortened', async () => {
  const result = await createApp().callTool('docs_get', args);
  assert.equal(result.structuredContent.body, document.body);
  assert.match(result.content[0].text, /more characters/);
  assert.doesNotMatch(result.content[0].text, /tail/);
});

test('sheets_read preserves an empty range as an array', async () => {
  const app = createApp(fakeClient({ readSheetValues: async () => ({ range: 'Sheet1!A1:D2', values: [] }) }));
  const { structuredContent: data } = await app.callTool('sheets_read', args);
  assert.deepEqual(data.values, []);
  assert.equal(data.count, 0);
});


test('update_event preserves explicit empty fields for clearing without clearing omitted fields', async () => {
  const calls = [];
  const app = createApp(fakeClient({ updateCalendarEvent: async (...args) => { calls.push(args); return event; } }));
  await app.callTool('update_event', { account: 'work', calendar_id: 'primary', event_id: 'event-1', description: '', location: '' });
  assert.equal(calls[0][2].description, '');
  assert.equal(calls[0][2].location, '');
  assert.equal(calls[0][2].summary, undefined);
});
