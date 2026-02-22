# Outlook Email Agent v3.0

A VBA-based email agent for Microsoft Outlook that classifies emails using a priority rule chain, learns from your manual sorting, and drafts replies using LLM few-shot examples.

## What's New in v3.0

- **Multi-provider LLM**: Local (Ollama/LM Studio/Inferencer), Azure OpenAI, Anthropic Claude, or any OpenAI-compatible API (OpenRouter, Groq, etc.) — switch via `settings.ini`
- **Auto-reply drafting**: Agent learns your reply style from curated examples and drafts replies for new emails
- **Addressing pattern generation**: LLM generates all name/greeting variations from your name input
- **Structured error logging**: Centralized `error.log` with call stack, replacing ad-hoc error dialogs
- **New `EmailAgent.bas` module**: Agent features cleanly separated from the filter core
- **LearnReply folder**: Drag sent replies here to teach the agent your reply style

## Features

- **Real-time filtering**: Automatically process new emails as they arrive
- **Batch processing**: Filter existing emails in bulk with dry-run preview
- **Rule-based classification**: Fast, offline pattern matching (10-rule priority chain)
- **Self-improving**: Learns from drag-and-drop sorting (LearnKeep / LearnDelete / LearnSubjectDelete)
- **Reply learning**: Learns your reply style from sent emails (LearnReply folder or Sent Items scan)
- **Multi-provider LLM**: Local, Azure, Claude, or OpenAI-compatible (OpenRouter, Groq, etc.) for classification, summarization, and reply drafting
- **Dry-run mode**: Preview all decisions before taking action
- **Server rule export**: Push learned DELETE rules to Exchange for 24/7 server-side filtering
- **External configuration**: All settings in `settings.ini` — no VBA code editing needed
- **Web UI**: Browser-based interface for settings, macros, learned rules, email browsing, and chat

## Quick Start

### Installation

1. **Enable macros**: File → Options → Trust Center → Macro Settings → Enable all macros
2. **Import VBA modules** (Alt+F11 → File → Import): `Config.bas`, `Utilities.bas`, `EmailFilter.bas`, `EmailAgent.bas`, `BatchFilter.bas`
3. **Paste ThisOutlookSession**: Copy-paste `src/ThisOutlookSession.bas` into the built-in module (see [docs/INSTALL.md](docs/INSTALL.md))
4. **Compile**: Debug → Compile Project → Ctrl+S
5. **Restart Outlook** — `settings.ini` is auto-created with defaults
6. **Create learning folders** under Inbox: `LearnKeep`, `LearnDelete`, `LearnSubjectDelete`, `LearnReply`
7. **Test**: In Immediate Window (Ctrl+G), type `FilterExistingDryRun`

See [docs/INSTALL.md](docs/INSTALL.md) for full installation details.

## Web UI (Optional)

A browser-based interface for managing the agent without touching the VBA Editor.

```bash
cd webui
pip install -r requirements.txt
python server.py
# → open http://localhost:5000
```

| Tab | What you can do |
|-----|----------------|
| **Settings** | View and edit all `settings.ini` sections in-browser, save with one click |
| **Macros** | Click-to-run buttons for all major macros (requires Outlook running) |
| **Learned Rules** | Browse sender rules, subject rules, and reply examples |
| **Emails** | Read-only email browser via Outlook COM (Windows only) |
| **Chat** | Conversational commands: "dry run", "show version", "provider claude", etc. |
| **Logs** | Live view of `error.log` |

The Web UI reads data files directly. Macro execution uses a file-based command bridge: the server writes a JSON file, the VBA poller (started automatically at Outlook startup) picks it up, executes the macro, and writes the result. Outlook must be running for macros to execute.

Requires: Python 3.10+, Windows for COM features (settings/rules/logs work on any OS). See [webui/README.md](webui/README.md) for full details.

## Classification Rules

| Priority | Rule | Action |
|----------|------|--------|
| 0 | **Learned sender rule** (self-improving) | Keep or Delete |
| 0.5 | **Learned subject rule** (self-improving) | Delete |
| 1 | Protected domain (e.g., substack.com) | Move to Protected |
| 2 | Personally addressed (name/greeting) | Keep |
| 3 | Organizational tags ([MM], [HRO]) | Keep |
| 4 | VIP keywords (thesis, deadline) | Keep |
| 5 | Reply chain (RE:, AW:) | Keep |
| 6 | Forward chain (FW:, FWD:, WG:) | Keep |
| 7 | Known spam sender names | Delete |
| 8 | Spam sender email patterns (noreply) | Delete |
| 9 | Spam subject keywords (newsletter) | Delete |
| 10 | No match | Review folder (or LLM) |

First match wins. Rules 0 and 0.5 are learned from your drag-and-drop actions.

## Files

```
outlook-email-filter/
├── src/
│   ├── Config.bas                 # DEFAULT_* constants + Runtime* variables
│   ├── Utilities.bas              # Helpers, CallLLM, INI I/O, error handling, bridge helpers
│   ├── EmailFilter.bas            # Classification engine + LLM wrappers
│   ├── EmailAgent.bas             # Agent: addressing, auto-reply, reply learning
│   ├── BatchFilter.bas            # Bulk processing + macro launchers
│   └── ThisOutlookSession.bas     # Event handlers + Web UI command poller (copy-paste, not import)
├── webui/
│   ├── server.py                  # Flask app — all API routes
│   ├── bridge.py                  # File-based command bridge to VBA
│   ├── settings_manager.py        # settings.ini read/write
│   ├── chat.py                    # Keyword command parser
│   ├── requirements.txt           # flask, pywin32
│   ├── static/                    # index.html, style.css, app.js (SPA)
│   └── README.md                  # Web UI setup and API reference
├── docs/
│   ├── INSTALL.md                 # Installation guide (fresh install, upgrade, migration)
│   ├── USER_MANUAL.md             # Complete user reference
│   ├── PATTERNS.md                # Pattern configuration guide
│   ├── macros.md                  # Full macro reference table
│   └── data-files.md              # Data file formats and locations
├── CLAUDE.md                      # Developer reference
└── README.md
```

## Configuration

All settings live in `%APPDATA%\OutlookEmailFilter\settings.ini` (auto-created on first run). Edit with any text editor — pattern changes take effect immediately, folder name changes require an Outlook restart.

### Key Macros

| Macro | What It Does |
|-------|-------------|
| `FilterExistingDryRun` | Preview what the filter would do (no changes) |
| `FilterExistingEmails` | Filter all Inbox emails |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation |
| `FilterCurrentFolder` | Filter current folder with confirmation |
| `GenerateAddressingPatterns` | LLM-generates your name/greeting patterns |
| `DraftReplyToSelected` | Draft a reply using LLM + learned examples |
| `ScanSentForReplyPatterns` | Scan Sent Items to learn your reply style |
| `ShowVersionInfo` | Show version, data file paths, LLM provider, status |

See [docs/macros.md](docs/macros.md) for the full macro reference.

## Self-Improving Filter

The agent learns from your manual sorting — no configuration needed:

| Action | What the Agent Learns |
|--------|-----------------------|
| Drag email to **LearnKeep** | Always **keep** future emails from that sender |
| Drag email to **LearnDelete** | Always **delete** future emails from that sender |
| Drag email to **LearnSubjectDelete** | Always **delete** future emails with matching subject |
| Drag sent reply to **LearnReply** | Learns your reply style for that type of email |

Learned rules are stored in `%APPDATA%\OutlookEmailFilter\` and persist across restarts.

## LLM Integration (Optional)

Set `UseLLMAPI=True` in `settings.ini` and configure your provider:

| Provider | Setting | Notes |
|----------|---------|-------|
| **Local** (Ollama, LM Studio, Inferencer) | `Provider=local` | Free, private, requires local server |
| **Azure OpenAI** | `Provider=azure` | Requires Azure subscription + deployment |
| **Anthropic Claude** | `Provider=claude` | Requires Anthropic API key |
| **OpenAI-compatible** (OpenRouter, Groq, etc.) | `Provider=openai` | Requires API key, any OpenAI-compatible endpoint |

LLM features: classify ambiguous emails, summarize, draft replies with your learned style.

See [docs/INSTALL.md](docs/INSTALL.md) for provider setup instructions.

## Requirements

- Microsoft Outlook desktop (Windows only — not "New Outlook", not the web app)
- Outlook 2016, 2019, 2021, or Microsoft 365
- **Web UI only**: Python 3.10+, `pip install -r webui/requirements.txt`

## License

MIT License — use freely, modify as needed.
