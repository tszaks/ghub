# GitHub releases

The `Verify and release` workflow runs typecheck, the full mocked test suite
(including its build), and a committed-`dist/` consistency check (including newly generated untracked files) for pull requests
and pushes to `main`. Pull requests have read-only repository permissions and
cannot publish.

After successful verification on `main`, a separate job packages the exact
commit with `npm pack --ignore-scripts`. It creates `v<VERSION>` from
`package.json` and uploads `ghub-<VERSION>.tgz` to the GitHub release. For the CLI
release these are `v1.7.0` and `ghub-1.7.0.tgz`. The release stays draft until the
package upload succeeds. Existing complete releases are left unchanged; retries
can finish an incomplete release only when it targets the same commit. Version
tags are never moved and assets are never overwritten. Before publishing, the
workflow downloads the release asset and verifies its SHA-256 against the
just-built package, including when resuming an existing draft.

To release another version, update `package.json` and `package-lock.json`, rebuild
and commit `dist/`, then merge the reviewed change to `main`. Use stable semantic
versions (`MAJOR.MINOR.PATCH`). A failed workflow can be rerun from GitHub Actions.
A version/tag conflict stops publication for review instead of replacing code.

Only the release job receives `contents: write`, through GitHub's automatic,
short-lived workflow token. Checkout credentials are not persisted. This does not
publish to the npm registry, install new long-lived credentials, authenticate
Google accounts, contact live inboxes, or update a running MCP installation.
Download the release tarball and install it with `npm install -g ./ghub-1.7.0.tgz`
on the machine where the CLI/MCP should run, then reload that machine's MCP
client if necessary.

Official references:
- [GitHub workflow token](https://docs.github.com/en/actions/tutorials/authenticate-with-github_token)
- [Creating releases with GitHub CLI](https://cli.github.com/manual/gh_release_create)
