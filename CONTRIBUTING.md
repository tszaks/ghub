# Contributing to ghub

Thanks for helping. ghub gives AI agents access to real mail, files, and calendars, so the bar for changes is: does it stay simple, and does it keep account data safe?

By taking part you agree to follow the [Code of Conduct](CODE_OF_CONDUCT.md).

## Before you start

- For bugs, open an issue with the template. Do not paste tokens, credentials, or mail content.
- For features, open an issue first and describe the problem you want to solve.
- Small fixes (typos, clear bugs) can go straight to a pull request.
- Security problems go through [SECURITY.md](SECURITY.md), not a public issue.

## Setup

You need Node.js 20 or newer.

```sh
git clone https://github.com/tszaks/ghub.git
cd ghub
npm install
npm run build
```

`npm run dev` rebuilds on change. `node dist/index.js tools` lists the tools without touching any account.

## Run the tests

```sh
npm run typecheck
GMAILMCPCONFIG_DIR=$(mktemp -d) npm test
```

`npm test` builds first, then runs `node --test test/*.test.mjs`. The suite is mocked: it does not call Google. Point `GMAILMCPCONFIG_DIR` at an empty folder so it never reads your real accounts.

## Code style

- **Match the surrounding code.** Same naming, same comment density, same idiom as the file you are editing. There is no formatter to run, so read the neighbors.
- Keep it simple. No new dependencies without discussion.
- Comments explain why, not what.

## Committed `dist/`

The compiled `dist/` is committed so servers can deploy without a build. CI fails if `dist/` differs from a fresh build. Run `npm run build` and commit every changed file in `dist/` with your source change.

## Safety rules

- Write tools must require an explicit `account`. Never pick a default account for a mutation.
- New tools need an entry in the tool registry with the right `kind` (`read`, `write`, or `auth`) and a test.
- Never log or return tokens, client secrets, authorization codes, or raw Google error payloads.
- Tests must not call live Google APIs or read a real config directory.

## Pull request checklist

- [ ] `npm run typecheck` and `npm test` pass locally.
- [ ] New behavior has a test.
- [ ] `dist/` is rebuilt and committed.
- [ ] The change is small and focused.
- [ ] A line added under `[Unreleased]` in [CHANGELOG.md](CHANGELOG.md) for anything a user would notice.
- [ ] Docs updated if tools, CLI behavior, or configuration changed.

## Using AI agents to contribute

Fine, as long as you read and understand what you submit and the checklist above is true. [AGENTS.md](AGENTS.md) tells coding agents how this repository works.

## License

By contributing you agree that your contribution is licensed under the [MIT License](LICENSE).
