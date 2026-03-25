import test from 'node:test';
import assert from 'node:assert/strict';
import { describeDriveApiError } from '../dist/gmail-client.js';

test('service-disabled Drive errors tell the operator to enable the Drive API', () => {
  const message = describeDriveApiError(
    {
      code: 403,
      message:
        'Google Drive API has not been used in project 31682034682 before or it is disabled.',
      response: {
        data: {
          error: {
            details: [
              {
                '@type': 'type.googleapis.com/google.rpc.ErrorInfo',
                reason: 'SERVICE_DISABLED',
                metadata: {
                  activationUrl:
                    'https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=31682034682',
                },
              },
            ],
          },
        },
      },
    },
    'Drive search failed.'
  );

  assert.match(message, /Enable the Google Drive API/);
  assert.match(message, /31682034682/);
});
