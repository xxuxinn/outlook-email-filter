# Outlook Email Agent — Web UI

A Python Flask web server that provides a browser-based interface for configuring and controlling the Outlook Email Agent.

## Features

- **Settings tab** — view and edit all 5 `settings.ini` sections in-browser
- **Macros tab** — click-to-run buttons for all major macros (via file bridge to Outlook)
- **Learned Rules tab** — browse sender rules, subject rules, and reply examples
- **Chat tab** — conversational command interface (keyword-based, no LLM required)
- **Logs tab** — view recent `error.log` and `llm_debug.log` entries

## Setup

### Requirements

- Python 3.10+
- Windows (for Outlook COM features; settings/logs/rules work on any OS)
- Outlook desktop running with VBA macros enabled (for macro execution via bridge)

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
Browser (localhost:5000)
    ↕ HTTP/JSON
Python Flask (server.py)
    ├── Direct: settings.ini read/write      (settings_manager.py)
    ├── Direct: learned_*.txt / error.log    (file I/O)
    └── Bridge: write command JSON           (bridge.py)
                    ↓
        %APPDATA%\OutlookEmailFilter\commands\<id>.json
                    ↓  (VBA polls every 2 seconds)
        VBA executes macro → writes <id>.result
                    ↓
        Server reads result and returns to browser
```

### Why a file bridge?

Outlook has no external `Application.Run()` API — external programs cannot invoke VBA macros directly. The file bridge is the standard pattern: write a JSON command file, let VBA pick it up via Win32 `SetTimer` polling, and read the result file.

### VBA Command Poller

The poller is in `src/Utilities.bas`:
- `StartCommandPollerStd` — called automatically from `Application_Startup`
- `PollForCommandsTimer` — runs every 2 seconds via Win32 `SetTimer`
- `StopCommandPollerStd` — called from `DisableRealTimeFilter`

**For macro results to appear in the Web UI, Outlook must be running** with the VBA project loaded. The poller starts automatically; no user action required.

## API Reference

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Serve the SPA |
| `/api/settings` | GET | Read settings.ini as JSON |
| `/api/settings` | POST | Write `{section, key, value}` |
| `/api/settings/section` | POST | Write `{section, values:{}}` |
| `/api/settings/reload` | POST | Send ReinitializeFilter to Outlook |
| `/api/learned/senders` | GET | Learned sender rules |
| `/api/learned/subjects` | GET | Learned subject rules |
| `/api/learned/replies` | GET | Learned reply examples |
| `/api/errors` | GET | error.log (last N lines, `?n=100`) |
| `/api/command` | POST | Send macro command to Outlook |
| `/api/command/<id>/result` | GET | Poll for command result |
| `/api/chat` | POST | Parse chat message → action |
| `/api/errors/clear` | POST | Truncate error.log |
| `/api/llm-debug-log` | GET | Read llm_debug.log (last N lines, `?n=100`) |
| `/api/llm-debug-log/clear` | POST | Truncate llm_debug.log |
| `/api/command/debug` | GET | List files in commands directory (debug) |
| `/api/status` | GET | Version, provider, rule counts |

## Chat Commands

| Say | Action |
|-----|--------|
| dry run / preview | Run FilterExistingDryRun |
| filter inbox | Run FilterExistingEmails |
| show version / status | Run ShowVersionInfo |
| show senders | View learned sender rules |
| show subjects | View learned subject rules |
| show replies | View learned reply examples |
| show errors | View error log |
| reload settings | Send ReinitializeFilter |
| provider local/azure/claude/openai | Switch LLM provider |
| enable/disable llm | Toggle LLM on/off |
| scan sent | ScanSentForReplyPatterns |
| draft replies / batch draft | Run DraftRepliesForInbox |
| restore review | Run RestoreFromReview |
| help | Show available commands |

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `5000` | Port to listen on |
