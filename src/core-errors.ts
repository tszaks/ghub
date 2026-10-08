/** Sanitized errors shared by transport adapters. Never expose provider request bodies. */
export class CliError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly exitCode = 2,
    readonly recovery = 'Run ghub --help or ghub schema TOOL and correct the input.',
    readonly retryable = false,
  ) { super(message); }
}

export function safeError(error: unknown, mutation = false): CliError {
  if (error instanceof CliError) return error;
  const candidate = error as { code?: unknown; message?: unknown; response?: { status?: number } } | null;
  const status = candidate?.response?.status ?? (typeof candidate?.code === 'number' ? candidate.code : 0);
  const message = typeof candidate?.message === 'string' ? candidate.message : '';
  const retryWarning = mutation ? ' The operation may have taken effect; inspect the target before retrying.' : '';
  if (status === 401 || /invalid_grant/.test(message)) {
    return new CliError('AUTH_REQUIRED', 'The selected account is missing, disabled, or requires authentication.', 3,
      'Run ghub call list_accounts, then use the documented separate OAuth bootstrap if needed.' + retryWarning);
  }
  if (status === 429 || status >= 500 || /ETIMEDOUT|ECONNRESET|EAI_AGAIN|ENOTFOUND|ECONNREFUSED/.test(String(candidate?.code))) {
    return new CliError('UPSTREAM_UNAVAILABLE', 'The upstream service is unavailable or rate limited.', 5,
      mutation ? 'Inspect the target before retrying; the operation may have taken effect.' : 'Retry later with exponential backoff.', !mutation);
  }
  if (status === 403) return new CliError('PERMISSION_DENIED', 'The account cannot perform this operation.', 4,
    'Check Google API enablement, granted scopes, and resource permissions.' + retryWarning);
  if (status === 404) return new CliError('NOT_FOUND', 'The requested resource was not found.', 4,
    'Verify the resource ID and its account.' + retryWarning);
  if (status >= 400) return new CliError('API_ERROR', 'The upstream service rejected the request.', 4,
    'Check the command schema and resource state.' + retryWarning);
  if (/accounts config/i.test(message)) return new CliError('CONFIG_ERROR', 'The account configuration could not be loaded.', 3,
    'Check --config-dir or GMAILMCPCONFIG_DIR and validate accounts.json.');
  if (/token file|credentials|oauth|no enabled|disabled|unknown account/i.test(message)) {
    return new CliError('AUTH_REQUIRED', 'The selected account is missing, disabled, or requires authentication.', 3,
      'Run ghub call list_accounts, then use the documented separate OAuth bootstrap if needed.' + retryWarning);
  }
  if (/required|must be|invalid|provide |unsupported|unknown .*argument|not found|ENOENT|cannot read/i.test(message)) {
    return new CliError('INVALID_ARGUMENT', 'An argument or local input file is invalid.', 2,
      'Check the schema, resource IDs, and readable input/attachment paths.' + retryWarning);
  }
  return new CliError('INTERNAL_ERROR', 'The operation failed unexpectedly.', 1,
    'Check configuration and report the command name and error code without credentials.' + retryWarning);
}

