# Changelog

All notable changes to this project will be documented in this file.

## 0.1.0 - 2025-12-12

Initial public release of `gog`: a single Go CLI that unifies Gmail, Calendar, Drive, and Contacts (People API).

### Added

- Unified CLI (`gog`) with service subcommands: `gmail`, `calendar`, `drive`, `contacts`, plus `auth`.
- OAuth setup and account management:
  - Store OAuth client credentials: `gog auth credentials <credentials.json>`.
  - Authorize accounts and store refresh tokens securely via OS keychain using `github.com/99designs/keyring`.
  - List/remove accounts: `gog auth list`, `gog auth remove <email>`.
  - Token management helpers: `gog auth tokens list|delete|export|import`.
- Consistently parseable output:
  - `--output=text` (tab-separated lists on stdout) and `--output=json` (JSON on stdout).
  - Human hints/progress/errors go to stderr.
- Colorized output in rich TTY (`--color=auto|always|never`), automatically disabled for JSON output.
- Gmail features:
  - Search threads, show thread, generate web URLs.
  - Label listing/get (including counts) and thread label modify.
  - Send mail (supports reply headers + attachments).
  - Drafts: list/get/create/send/delete.
- Calendar features:
  - List calendars, list ACL rules.
  - List/get/create/update/delete events and free/busy queries.
- Drive features:
  - List/search/get files, download (including Google Docs export), upload, mkdir, delete, move, rename.
  - Sharing helpers: share/unshare/permissions, and web URL output.
- Contacts / People API features:
  - Contacts list/search/get/create/update/delete.
  - “Other contacts” list/search.
  - Workspace directory list/search (Workspace accounts only).
- Developer experience:
  - Formatting via `gofumpt` + `goimports` (and `gofmt` implicitly) using `make fmt` / `make fmt-check`.
  - Linting via pinned `golangci-lint` with repo config.
  - Tests using stdlib `testing` + `httptest`, with steadily increased unit coverage.
  - GitHub Actions CI running format checks, tests, and lint.
  - `pnpm gog` helper to build+run (`pnpm gog auth add you@gmail.com`).

### Notes / Known Limitations

- Importing tokens into macOS Keychain may require a local (GUI) session; headless/SSH sessions can fail due to Keychain user-interaction restrictions.
- Workspace directory commands require a Google Workspace account; `@gmail.com` accounts will not work for directory endpoints.

