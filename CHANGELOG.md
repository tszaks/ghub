# Changelog

All notable changes to ghub are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and ghub uses [Semantic Versioning](https://semver.org/). Entries are built from git tags and commit history, and dates are the tag dates.

## [Unreleased]

### Changed
- Documentation reorganized: new README, CONTRIBUTING, SECURITY, AGENTS, Code of Conduct, issue and pull request templates, and `docs/api-reference.md`.

## [1.7.0] - 2026-10-08

### Added
- Agent CLI: `ghub tools`, `ghub schema TOOL`, and `ghub call TOOL` expose all tools as noninteractive JSON calls.
- Verified GitHub releases built by the `Verify and release` workflow.

### Fixed
- Message encoding, threading, and attachments that were silently dropped on send.
- Attachments are fetched in a single API round trip.

## [1.6.2] - 2026-06-09

### Fixed
- Attachment fetching survives Gmail attachment ID rotation.

## [1.6.1] - 2026-05-18

### Fixed
- Non-ASCII characters in email subjects are RFC 2047 encoded.

## [1.6.0] - 2026-05-15

### Fixed
- Attachment fetching for inlined attachments and stale attachment IDs.

## [1.5.1] - 2026-05-04

### Added
- `gmail.settings.basic` scope for filter and `block_sender` support.
- Pre-built `dist/` is committed for zero-build server deploys.

## [1.5.0] - 2026-05-01

### Changed
- Clearer attachment guidance for AI agents.

## [1.4.1] - 2026-04-22

### Added
- Threading headers and Cc are exposed in parsed emails.

## [1.4.0] - 2026-04-22

### Added
- Renamed the project to ghub. Added Drive, Sheets, Docs, and Calendar tools.
- `send_draft`, `list_drafts`, and `search_drafts`.
- `thread_id`, `in_reply_to`, and `references` support in `create_draft`.
- SSE transport for hosted deployments, with a Railway port fallback.

## [1.3.1] - 2026-04-21

### Changed
- Raised the `max_results` cap to 500.

## [1.3.0] - 2026-04-17

### Added
- Inbox control tools: block sender, unsubscribe, and mute thread.

## [1.2.0] - 2026-04-16

### Added
- Attachment extraction with OCR.
- Google Drive search.

## [1.1.0] - 2026-03-19

Initial tagged release.
