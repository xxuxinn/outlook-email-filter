# Outlook Email Agent — Web UI

A Python Flask web server that provides a browser-based interface for configuring and controlling the Outlook Email Agent.

**Technical:** The server binds to `127.0.0.1` only, and every `/api/*` route requires an `X-Auth-Token` header matching a random 32-hex-char token stored at `%APPDATA%\OutlookEmailFilter\webui_token.txt`. The token is injected into `index.html` when it is served, so the browser gets it automatically.
**In plain terms:** the UI only works on the same computer, and every request carries a secret password that the page picks up for you when it loads — other programs or people on your network can't call the API without it.

## Features

- **Settings tab** — view and edit all `settings.ini` sections in-browser (secrets are masked)
- **Macros tab** — click-to-run buttons for all allowlisted macros (via file bridge to Outlook)
- **Learned Rules tab** — browse sender rules, subject rules, and reply examples
- **Digest tab** — view the latest daily digest markdown; generate one on demand
- **Proposals tab** — review LLM-proposed rules; approve (appends to learned rules + reloads Outlook) or reject
- **Decisions tab** — table of the last 100 classification decisions from `decision_log.txt`
- **Chat tab** — conversational command interface (keyword-based, no LLM required)
- **Logs tab** — view recent `error.log` and `llm_debug.log` entries

## Setup

### Requirements

- Python 3.10+
- Any OS for settings/logs/rules browsing; Windows with Outlook desktop running (VBA macros enabled) for macro execution, digest generation, and rule proposals via the bridge

### Install

```bash
cd webui
pip install -r requirements.txt
```

### Run

```bash
python server.py
```

Open `http://localhost:5000` in your browser.

## Architecture

```
Browser (localhost:5000, X-Auth-Token on every /api call)
    ↕ HTTP/JSON
Python Flask (server.py)
    ├── auth.py             token load/create + constant-time check
    ├── config.py           all data-dir paths (env-overridable for tests)
    ├── macros.py           server-side macro allowlist + arg validation
    ├── settings_manager.py settings.ini read/write (utf-8-sig, cp950 fallback)
    ├── datafiles.py        learned rules, digest, decisions, proposals I/O
    ├── chat.py             keyword → action parser (validated against macros.py)
    └── bridge.py           write command JSON → VBA polls → result JSON
                    ↓
        %APPDATA%\OutlookEmailFilter\commands\<id>.json
                    ↓  (VBA polls every 2 seconds)
        VBA executes macro → writes <id>.result
                    ↓
        Server reads result (and deletes both files) → returns to browser
```

**Technical:** There is no COM integration — all Outlook interaction goes through the file bridge; everything else is plain file I/O against `%APPDATA%\OutlookEmailFilter\`.
**In plain terms:** the web server never talks to Outlook directly. It leaves note files in a shared folder; a small VBA timer inside Outlook checks that folder every 2 seconds, does the work, and leaves a reply note.

### Why a file bridge?

**Technical:** Outlook has no external `Application.Run()` API — external programs cannot invoke VBA macros directly. The bridge writes a JSON command file, VBA picks it up via Win32 `SetTimer` polling, and writes a result file.
**In plain terms:** Outlook can't be remote-controlled, so we pass notes through a mailbox folder instead — like sliding requests under a door and waiting for the answer to slide back.

### VBA Command Poller

The poller is in `src/Utilities.bas`:
- `StartCommandPollerStd` — called automatically from `Application_Startup`
- `PollForCommandsTimer` — runs every 2 seconds via Win32 `SetTimer`
- `StopCommandPollerStd` — called from `DisableRealTimeFilter`

**For macro results to appear in the Web UI, Outlook must be running** with the VBA project loaded. The UI checks `GET /api/bridge/health` before sending long-running macros and shows an "Outlook poller not responding" banner if command files are sitting unconsumed (older than 10 s).

## Security Model

**Technical:**
- Token auth on all `/api/*` routes (`X-Auth-Token`, compared with `hmac.compare_digest`); `/` and `/static` are open. 401 JSON on failure.
- Server-side macro allowlist in `macros.py` — unknown macros are rejected with 400; declared args are validated (`days` int 1–365, `pattern` string 3–100 chars, undeclared args rejected).
- `GET /api/settings` masks `APIKeyHardcoded` as `__MASKED__`; writing `__MASKED__` back is a silent no-op so the key can never be leaked or clobbered through the UI.
- INI writes are restricted to known sections (General, Folders, Patterns, LLM, Agent, Sync, Digest), keys matching `^[A-Za-z0-9_]{1,64}$`, and values without CR/LF (rejected with 400 — prevents INI section injection).
- Command/proposal ids must match `^[0-9a-f]{8}$` (blocks path traversal).
- Binds `127.0.0.1`, `debug=False`.

**In plain terms:** the server keeps its own short list of allowed buttons and double-checks every request against it — even a script that steals the page's JavaScript can't invent new commands, sneak extra parameters in, read your API key back out, or trick the server into opening files outside its folder.

## API Reference

**Technical:** all `/api/*` routes require the `X-Auth-Token` header; `/` is open and serves the token.
**In plain terms:** this table is the complete "menu" of things the browser can ask the server to do — you'd normally never call these by hand, but they're handy for debugging with `curl`.

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Serve the SPA (auth token substituted into index.html) |
| `/api/settings` | GET | Read settings.ini as JSON (secrets masked) |
| `/api/settings` | POST | Write `{section, key, value}` (validated) |
| `/api/settings/section` | POST | Write `{section, values:{}}` (validated) |
| `/api/settings/reload` | POST | Send ReinitializeFilter to Outlook |
| `/api/macros` | GET | Macro manifest (name, label, description, args, destructive, category) |
| `/api/learned/senders` | GET | Learned sender rules |
| `/api/learned/subjects` | GET | Learned subject rules |
| `/api/learned/replies` | GET | Learned reply examples |
| `/api/errors` | GET | error.log (last N lines, `?n=100`, clamped 1–1000) |
| `/api/errors/clear` | POST | Truncate error.log (reports failure honestly) |
| `/api/llm-debug-log` | GET | Read llm_debug.log (last N lines, `?n=200`, clamped 1–1000) |
| `/api/llm-debug-log/clear` | POST | Truncate llm_debug.log (reports failure honestly) |
| `/api/command` | POST | Send allowlisted macro command to Outlook (400 on unknown macro/bad args) |
| `/api/command/<id>/result` | GET | Poll for command result (id must be 8 hex chars; files deleted after read) |
| `/api/command/debug` | GET | List files in commands directory (debug) |
| `/api/bridge/health` | GET | `{poller_responsive, stale_commands}` — detects a dead Outlook poller |
| `/api/digest` | GET | Latest `digests/digest_YYYY-MM-DD.md` as `{date, content}` (404 if none) |
| `/api/digest/generate` | POST | Send GenerateDailyDigest to Outlook, returns command id |
| `/api/proposals` | GET | All rule_proposals.txt entries |
| `/api/proposals/generate` | POST | Send ProposeRules to Outlook, returns command id |
| `/api/proposals/<id>/approve` | POST | Append rule to learned file, mark APPROVED, send reload to Outlook |
| `/api/proposals/<id>/reject` | POST | Mark proposal REJECTED |
| `/api/decisions` | GET | Last N decision_log.txt rows (`?n=100`, clamped 1–1000) |
| `/api/chat` | POST | Parse chat message → action |
| `/api/status` | GET | Version, provider, rule counts, bridge health, latest digest date |

## Chat Commands

**Technical:** keyword substring matching, first match wins; every macro mapping is validated against the `macros.py` manifest.
**In plain terms:** type any phrase containing one of these keywords and the right action fires — there's no AI parsing here, just a lookup table, so wording must roughly match.

| Say | Action |
|-----|--------|
| dry run / preview | Run FilterExistingDryRun |
| filter inbox | Run FilterExistingEmails |
| generate digest / daily digest | Run GenerateDailyDigest |
| propose rules / suggest rules / mine rules | Run ProposeRules |
| show version / version / status | Run ShowVersionInfo |
| show senders | View learned sender rules |
| show subjects | View learned subject rules |
| show replies | View learned reply examples |
| show decisions / decision log | View decision log |
| show errors | View error log |
| reload settings | Send ReinitializeFilter |
| provider local/azure/claude/openai | Switch LLM provider |
| enable/disable llm | Toggle LLM on/off |
| scan sent | Run ScanSentForReplyPatterns |
| draft replies / draft reply | Run DraftReplyForSelected |
| sync rules | Run SyncLearnedRules |
| restore review | Run RestoreFromReview |
| help | Show available commands |

## Data Encoding

**Technical:** settings.ini is written as UTF-8 with BOM (matching the VBA side); reads try `utf-8-sig` → `cp950` → `latin-1` so legacy ANSI files still load. Learned-rule and proposal files are read as UTF-8 with `errors="replace"`.
**In plain terms:** old settings files saved in the Windows Chinese encoding still open correctly — the server tries the modern format first and falls back to the legacy one, so Chinese patterns like 優惠 don't turn into garbage.

## Tests

```bash
pip install pytest
pytest tests/test_webui.py -v
```

**Technical:** the suite injects a temp data dir via the `OUTLOOK_FILTER_DATA_DIR` env var (config.py resolves paths at call time), covering auth, the macro allowlist, secret masking, encoding fallback, cmd-id validation, the proposals approve/reject flow, digest endpoints, result-file lifecycle, and n-param clamping.
**In plain terms:** tests run against a throwaway folder, never your real Outlook data, and they exercise the same security checks a browser request would hit.

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `5000` | Port to listen on |
| `OUTLOOK_FILTER_DATA_DIR` | `%APPDATA%\OutlookEmailFilter` | Override the data directory (used by tests / non-Windows dev) |
