# AGENTS.md

ghub is a TypeScript (Node 20+, ESM) MCP server and JSON CLI for Gmail, Drive, Sheets, Docs, and Calendar across several Google accounts. This file is for AI agents that work on the code. To use ghub as a tool, read [docs/agent-cli.md](docs/agent-cli.md).

## Build and test
- Install: `npm install` (the `postinstall` step compiles an OCR helper only if `swiftc` exists). In CI-like runs use `npm ci --ignore-scripts`.
- Build: `npm run build`. Typecheck: `npm run typecheck`.
- Test: `GMAILMCPCONFIG_DIR=$(mktemp -d) npm test`. Tests are mocked and never call Google. Always point `GMAILMCPCONFIG_DIR` at an empty folder.
- List tools without any account: `node dist/index.js tools`. Schema of one tool: `node dist/index.js schema TOOL`.
- `dist/` is committed on purpose. After any change in `src/`, run `npm run build` and commit the result. CI fails on drift.

## Code map
- `src/index.ts`: entrypoint. No args or `mcp` starts the MCP server, anything else runs the CLI.
- `src/mcp.ts`: stdio or SSE transport selection (`MCP_TRANSPORT`, `PORT`, `HOST`).
- `src/app.ts`: the tool registry, MCP request handlers, and tool implementations.
- `src/cli.ts`, `src/cli-contract.ts`, `src/cli-schema.ts`: the CLI, its JSON envelope, and schema output.
- `src/gmail-client.ts`: Google API clients, OAuth scopes, Gmail, Drive, Sheets, Docs, Calendar calls.
- `src/accounts.ts`, `src/account-session.ts`, `src/config.ts`: account store and config directory.
- `src/attachments.ts`, `vendor/ocr/`: attachment extraction and the macOS OCR helper.
- `src/outgoing-email.ts`: message building (encoding, threading, attachments).
- `test/`: `node --test` suites. `docs/`: CLI guide and spec, release process, API reference.

## Conventions
- Match the surrounding code: naming, comment density, idiom.
- Keep the tool count and docs honest: `README.md` and `docs/agent-cli.md` state the number of tools. Update both when you add or remove one.
- Every tool has a `kind` (`read`, `write`, `auth`). Write tools require an explicit `account`.
- CLI stdout is exactly one JSON object. Do not print diagnostics to stdout.
- Never log or return tokens, secrets, authorization codes, or raw Google errors.

## Do not
- Run `begin_account_auth`, `finish_account_auth`, or any write tool against a real account while developing.
- Commit `credentials.json`, `token.json`, or `.env` files.
- Add dependencies, telemetry, or network calls to anything but Google's APIs.
- Edit the release workflow to widen its permissions. See [docs/releases.md](docs/releases.md).
