import test from 'node:test';
import assert from 'node:assert/strict';
import { generateAuthUrlFromCredentials } from '../dist/gmail-client.js';

const sampleCredentials = {
  installed: {
    client_id: 'client-id',
    client_secret: 'client-secret',
    redirect_uris: ['http://localhost'],
  },
};

test('begin auth requests Drive metadata access alongside Gmail', () => {
  const { authUrl } = generateAuthUrlFromCredentials(sampleCredentials);
  const url = new URL(authUrl);
  const scopes = url.searchParams.get('scope') ?? '';

  assert.match(scopes, /https:\/\/mail\.google\.com\//);
  assert.match(scopes, /https:\/\/www\.googleapis\.com\/auth\/drive\.metadata\.readonly/);
  assert.notEqual(url.searchParams.get('include_granted_scopes'), 'true');
});
