# Agent CLI v1 specification

## Goal

Expose the entire existing GHub tool catalog (Gmail, Drive, Sheets, Docs,
Calendar and account onboarding) as short-lived, noninteractive CLI calls. The
CLI and MCP invoke one registry, dispatcher, Google API client and account/auth
implementation. Agents never parse the MCP Markdown to obtain data.

## Compatibility

- `ghub` with no arguments keeps the existing stdio/SSE MCP behavior, including
  `MCP_TRANSPORT`, `PORT` and the existing account configuration directory.
- `ghub mcp` explicitly starts the same server.
- `ghub tools`, `ghub schema TOOL`, and `ghub call TOOL --json JSON` are the CLI.
- Every existing MCP tool name is callable. No separate account store or token
  grant is introduced. CLI failures never automatically start OAuth.

## Agent contract

- Exactly one compact JSON object plus a newline on stdout per CLI invocation.
  No prompts, banners, spinners, browser launches or human tables.
- Success envelope: `{ "ok": true, "data": {}, "meta": { "version": 1 } }`.
  Data is native records, IDs and counts, not a rendered Markdown response.
- Errors: `{ "ok": false, "error": { "code", "message", "retryable",
  "recovery" }, "meta": { "version": 1 } }`. Never print raw Google errors,
  request headers, authorization codes, credentials or token files.
- Exit 0 success; 2 invalid command/input; 3 account/auth/config required;
  4 permission/not-found/API rejection; 5 transient upstream failure;
  6 partial account/attachment success; 1 unexpected internal failure.
  Mutations are never automatically retried. An uncertain mutation error must
  tell the caller to inspect the target before retrying.
- `tools` lists names and short descriptions without dumping every schema.
  `schema TOOL` returns its input schema, output shape, side-effect classification
  and an example. `--help` is JSON discovery, not terminal UI.
- JSON input is supplied inline, from `--input FILE`, or explicitly from
  `--input -` (stdin). No implicit stdin reads. Inputs are bounded and schema
  validated before any account or Google API operation. Unknown fields fail.
- Required account IDs on all mutations are validated before execution. Existing
  cross-account read/search semantics are preserved and failures are surfaced.
- `--fields` projects dot-separated result paths (including fields in array
  elements). Result limits and string/output byte budgets keep replies compact;
  truncation is explicit metadata, never disguised as a complete response.
- `read_emails`, `search_emails`, and `list_drive_files` surface Google
  continuation tokens for explicit-account pages. A `page_token` requires
  one `account`, with the same query/filter inputs. Cross-account aggregation
  is a bounded top-k view which can discard fetched records; it exposes no
  continuation token. Callers must start separate explicit-account reads to
  page each account fully. Resolve output-array truncation before advancing
  a provider token. No fabricated global cursor or claim of snapshot
  consistency.
- OAuth bootstrap uses existing `begin_account_auth` / `finish_account_auth`.
  Credentials and codes should enter through protected files or explicit stdin,
  not shell history. Human consent is a separate setup step, not a call retry.

## Verification and release

Use isolated temporary config and mocked Google clients only. Cover command
registry parity, stdin/file JSON, strict validation, stdout and exit behavior,
projection/truncation, structured data for each product, OAuth sharing,
account routing, partial failures, Gmail pagination and MCP handshake. Run the
entire existing test suite, typecheck and build, and ship regenerated `dist/`
(the repository intentionally tracks compiled deployment artifacts). Review
before merge; deploy only through an identified existing target/workflow.

## Deliberate limits

This release does not add a second OAuth system, TUI, interactive confirmation
prompts, automatic write retries, arbitrary command execution, or automatic
bulk write batching. API-side list pagination remains API-specific; output
truncation metadata is not a durable continuation token.
