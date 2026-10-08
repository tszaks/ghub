/** Stable agent protocol. No provider error bodies or caller values are echoed. */
export const PROTOCOL_VERSION = 1;
import { CliError } from './core-errors.js';
export { CliError, safeError } from './core-errors.js';

export const READ_TOOLS = new Set([
  'list_accounts', 'read_emails', 'search_emails', 'search_drive_files', 'get_email_thread',
  'get_labels', 'list_blocked_senders', 'list_drafts', 'search_drafts', 'get_attachment',
  'get_all_attachments', 'list_drive_files', 'get_drive_file', 'get_drive_file_content',
  'sheets_get', 'sheets_read', 'docs_get', 'list_calendars', 'list_events', 'get_event',
]);
export function operationKind(name: string): 'read' | 'write' | 'auth' {
  if (name === 'begin_account_auth' || name === 'finish_account_auth') return 'auth';
  return READ_TOOLS.has(name) ? 'read' : 'write';
}


type Schema = Record<string, any>;
export function validateInput(value: unknown, schema: Schema, at = 'input'): void {
  if (schema.enum && !schema.enum.includes(value)) throw new CliError('INVALID_ARGUMENT', `${at} is not an allowed value.`);
  if (schema.type === 'object') {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new CliError('INVALID_ARGUMENT', `${at} must be an object.`);
    const record = value as Record<string, unknown>;
    for (const key of schema.required ?? []) {
      if (!Object.hasOwn(record, key)) throw new CliError('INVALID_ARGUMENT', `${at}.${key} is required.`);
    }
    for (const [key, item] of Object.entries(record)) {
      // Prototype keys are never meaningful tool arguments.
      if (['__proto__', 'prototype', 'constructor'].includes(key)) throw new CliError('INVALID_ARGUMENT', `${at} contains a forbidden property.`);
      const property = schema.properties && Object.hasOwn(schema.properties, key) ? schema.properties[key] : undefined;
      if (!property && schema.additionalProperties === false) throw new CliError('INVALID_ARGUMENT', `${at} contains an unknown property.`);
      if (property) validateInput(item, property, `${at}.${key}`);
    }
  } else if (schema.type === 'array') {
    if (!Array.isArray(value)) throw new CliError('INVALID_ARGUMENT', `${at} must be an array.`);
    if (value.length > 1000) throw new CliError('INVALID_ARGUMENT', `${at} exceeds 1000 items.`);
    if (schema.items) value.forEach((item, index) => validateInput(item, schema.items, `${at}[${index}]`));
  } else if (schema.type === 'string') {
    if (typeof value !== 'string') throw new CliError('INVALID_ARGUMENT', `${at} must be a string.`);
  } else if (schema.type === 'boolean') {
    if (typeof value !== 'boolean') throw new CliError('INVALID_ARGUMENT', `${at} must be a boolean.`);
  } else if (schema.type === 'number' || schema.type === 'integer') {
    if (typeof value !== 'number' || !Number.isFinite(value) || (schema.type === 'integer' && !Number.isInteger(value))) {
      throw new CliError('INVALID_ARGUMENT', `${at} must be a finite ${schema.type}.`);
    }
  }
  if (typeof value === 'number') {
    if (schema.minimum !== undefined && value < schema.minimum) throw new CliError('INVALID_ARGUMENT', `${at} is below the minimum.`);
    if (schema.maximum !== undefined && value > schema.maximum) throw new CliError('INVALID_ARGUMENT', `${at} exceeds the maximum.`);
  }
}

export function cliInputSchema(schema: Schema): Schema {
  const copy = structuredClone(schema);
  for (const [name, property] of Object.entries(copy.properties ?? {}) as Array<[string, Schema]>) {
    if (property.type === 'number') {
      property.type = 'integer';
      property.minimum = ['font_size', 'width_pixels', 'height_pixels', 'count', 'rows', 'columns'].includes(name) ? 1 : 0;
    }
  }
  if (copy.properties?.max_results) copy.properties.max_results = {
    ...copy.properties.max_results, type: 'integer', minimum: 1, maximum: 100,
    description: 'Maximum records to request (1-100). Use page_token where supported.',
  };
  return copy;
}

export function validateSemantics(name: string, args: Record<string, unknown>): void {
  const kind = operationKind(name);
  if (kind === 'write' && (typeof args.account !== 'string' || !args.account.trim())) {
    throw new CliError('ACCOUNT_REQUIRED', 'Mutations require an explicit account.');
  }
  for (const key of ['account', 'account_id']) {
    if (args[key] !== undefined && (typeof args[key] !== 'string' || !/^[a-zA-Z0-9_-]+$/.test(args[key] as string))) {
      throw new CliError('INVALID_ARGUMENT', `${key} must contain only letters, numbers, underscores, or hyphens.`);
    }
  }
  for (const [key, value] of Object.entries(args)) {
    if ((key.endsWith('_id') || key === 'query' || key === 'to') && typeof value === 'string' && !value.trim()) {
      if (key === 'query' && !name.startsWith('search_')) continue;
      throw new CliError('INVALID_ARGUMENT', `${key} must not be empty.`);
    }
    if ((key.endsWith('_ids') || key === 'attendees') && Array.isArray(value) && value.some((v) => typeof v !== 'string' || !v.trim())) {
      throw new CliError('INVALID_ARGUMENT', `${key} must contain nonempty strings.`);
    }
    if (key.endsWith('_ids') && Array.isArray(value) && !value.length) throw new CliError('INVALID_ARGUMENT', `${key} must not be empty.`);
  }
  if (args.credentials_json !== undefined && typeof args.credentials_json !== 'string' && (!args.credentials_json || typeof args.credentials_json !== 'object' || Array.isArray(args.credentials_json))) {
    throw new CliError('INVALID_ARGUMENT', 'credentials_json must be a JSON object or string.');
  }
  if (name === 'create_event' && ((!args.start_date_time && !args.start_date) || (!args.end_date_time && !args.end_date))) {
    throw new CliError('INVALID_ARGUMENT', 'create_event requires a start and end date or date_time.');
  }
  if (name === 'unblock_sender' && !args.filter_id && !args.sender) throw new CliError('INVALID_ARGUMENT', 'Provide filter_id or sender.');
  if (name === 'docs_apply_style' && Number(args.end_index) <= Number(args.start_index)) throw new CliError('INVALID_ARGUMENT', 'end_index must be greater than start_index.');
  if (args.page_token !== undefined && !args.account) throw new CliError('ACCOUNT_REQUIRED', 'Page tokens require one explicit account.');
}

/** Project through arrays without flattening their shape. */
export function project(value: unknown, paths: string[][]): unknown {
  if (paths.some((path) => path.length === 0)) return value;
  if (Array.isArray(value)) return value.map((item) => project(item, paths));
  if (!value || typeof value !== 'object') return value;
  const out: Record<string, unknown> = {};
  for (const key of new Set(paths.map((path) => path[0]))) {
    if (Object.hasOwn(value, key)) out[key] = project((value as Record<string, unknown>)[key], paths.filter((path) => path[0] === key).map((path) => path.slice(1)));
  }
  return out;
}

export interface Truncation { path: string; omittedItems?: number; originalLength?: number; }
export function boundData(value: unknown, itemLimit: number, stringLimit: number, notes: Truncation[], path = 'data'): unknown {
  if (typeof value === 'string' && value.length > stringLimit) {
    notes.push({ path, originalLength: value.length });
    return value.slice(0, stringLimit);
  }
  if (Array.isArray(value)) {
    if (value.length > itemLimit) notes.push({ path, omittedItems: value.length - itemLimit });
    return value.slice(0, itemLimit).map((item, index) => boundData(item, itemLimit, stringLimit, notes, `${path}[${index}]`));
  }
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, boundData(item, itemLimit, stringLimit, notes, `${path}.${key}`)]));
  return value;
}
