# Security policy

## Supported versions

Only the latest release of ghub gets security fixes.

| Version | Supported |
|---|---|
| 1.7.x (latest) | Yes |
| Older | No |

## Reporting a vulnerability

Please report security problems privately, not in a public issue.

Use GitHub private vulnerability reporting: open the **Security** tab of this repository and choose **Report a vulnerability**. If that is not available to you, email tyler@szakacsmedia.com.

Please include:

- What you found and what an attacker could do with it.
- Steps to reproduce.
- Your Node.js version and ghub version.

Do not include real tokens, credentials, or mail content. You can expect a first reply within a few days.

## Sensitive surface

ghub is an account-access tool. With a valid token it can read and send mail, read and change Drive files, Sheets, Docs, and Calendar events. It requests broad Google scopes (full Gmail, Drive, Sheets, Docs, Calendar). These areas matter most:

- Token and client-secret handling: files under the config directory (`~/.gmail-multi-mcp` by default) hold OAuth credentials in plain JSON.
- Anything that leaks tokens, authorization codes, or credentials into output, logs, or errors.
- Writes that run without an explicit `account`, or against a different account than the one requested.
- The SSE transport (`MCP_TRANSPORT=sse` or `PORT`). It has no built-in authentication and binds to `0.0.0.0` by default. Run it only on a private network or behind your own access control.
- Local file access: attachment paths, uploads, and saved attachments.

ghub sends no telemetry. Its only network traffic is to Google's APIs.
