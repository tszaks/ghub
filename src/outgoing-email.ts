export interface OutgoingEmailAttachmentArgs {
  path: string;
  filename?: string;
  content_type?: string;
}

export const OUTGOING_EMAIL_ARGS = [
  'account',
  'to',
  'subject',
  'body',
  'cc',
  'bcc',
  'html',
  'attachments',
  'thread_id',
  'in_reply_to',
  'references',
] as const;

function optionalString(value: unknown, field: string, index: number): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    throw new Error(`attachments[${index}].${field} must be a string.`);
  }
  return value.trim() || undefined;
}

// Parse the `attachments` tool argument. Anything that cannot be turned into a file
// to attach is an error: dropping it would send the message without the file.
export function valueToAttachmentArray(value: unknown): OutgoingEmailAttachmentArgs[] {
  if (value === undefined || value === null) return [];

  // Some MCP clients pass array arguments as a JSON string.
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return [];
    if (trimmed.startsWith('[')) {
      try {
        value = JSON.parse(trimmed);
      } catch {
        throw new Error('attachments is a string that is not valid JSON. Pass an array of {path} objects.');
      }
    } else {
      value = [trimmed];
    }
  }

  if (!Array.isArray(value)) {
    throw new Error('attachments must be an array of {path, filename?, content_type?} objects or path strings.');
  }

  return value.map((item, index) => {
    if (typeof item === 'string') {
      const filePath = item.trim();
      if (!filePath) throw new Error(`attachments[${index}] is an empty path.`);
      return { path: filePath };
    }

    if (!item || typeof item !== 'object' || Array.isArray(item)) {
      throw new Error(`attachments[${index}] must be an object with a "path" or a path string.`);
    }

    const candidate = item as Record<string, unknown>;
    const filePath = typeof candidate.path === 'string' ? candidate.path.trim() : '';
    if (!filePath) {
      const keys = Object.keys(candidate).join(', ') || 'none';
      throw new Error(`attachments[${index}] has no "path" (got keys: ${keys}).`);
    }

    const attachment: OutgoingEmailAttachmentArgs = { path: filePath };
    const filename = optionalString(candidate.filename, 'filename', index);
    if (filename) attachment.filename = filename;
    const contentType = optionalString(candidate.content_type, 'content_type', index);
    if (contentType) attachment.content_type = contentType;
    return attachment;
  });
}

// Arguments the tool does not define are ignored by the MCP SDK, so a caller who
// writes `file_path` instead of `attachments` would otherwise send without the file.
export function assertKnownOutgoingEmailArgs(rawArgs: Record<string, unknown>): void {
  const known = new Set<string>(OUTGOING_EMAIL_ARGS);
  const unknown = Object.keys(rawArgs).filter((key) => !known.has(key));
  if (unknown.length > 0) {
    throw new Error(
      `Unknown argument(s): ${unknown.join(', ')}. Files are attached with attachments: [{ "path": "/absolute/path" }].`
    );
  }
}
