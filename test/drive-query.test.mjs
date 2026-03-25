import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDriveSearchQuery } from '../dist/gmail-client.js';

test('drive search query targets names and full text while excluding trashed files', () => {
  assert.equal(
    buildDriveSearchQuery('budget 2026'),
    "trashed = false and (name contains 'budget 2026' or fullText contains 'budget 2026')"
  );
});
