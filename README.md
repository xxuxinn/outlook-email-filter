# Outlook Email Agent v3.1

A VBA-based email agent for Microsoft Outlook that classifies emails using a priority rule chain with a structured, confidence-gated LLM fallback, learns from your manual sorting AND from its own mistakes, drafts replies in your style, delivers a ranked daily triage digest, and proposes new filtering rules for your approval.

## What's New in v3.1

- **Daily triage digest**: a ranked morning briefing (*Needs action / Worth a look / FYI*) built from the last 24 h, grouped by conversation, saved as markdown and emailed to yourself — scheduled automatically, no cron needed
  - *In plain terms:* instead of opening a full inbox, you get one email that says "these 3 need answers, this deadline is Friday, the rest was junk and here's why."
- **Structured LLM classification**: JSON verdicts with category, urgency (1–5), and confidence; uncertain DELETEs are demoted to the Review folder instead of the bin; urgent KEEPs get an "Urgent" category and high-importance flag inside Outlook
- **Context-enriched decisions**: the LLM is told the sender's track record ("47 emails, 45 deleted, you never replied") from the new `decision_log.txt` before it judges
- **Correction loop**: reverse an LLM decision by dragging to a learn folder, and that mistake is replayed to the model as a "don't do this again" example in every future classification
- **Rule mining**: weekly, the LLM proposes new sender/subject rules from your Review folder and history — you approve or reject each in the Web UI before anything goes live
- **MCP server** (`mcp/`): connect Claude Desktop or Claude Code directly to the agent — run dry-runs, read learned rules, generate digests, add rules, all through the local file bridge
- **Real-time = batch**: emails arriving live now get the same LLM classification as batch runs (previously they skipped the LLM entirely)
- **Hardened foundations**: HTTP timeouts on all LLM calls (a stalled server can no longer freeze Outlook), re-entrancy-guarded command poller, honest bridge results with real counts, token-authenticated Web UI, UTF-8 data files end-to-end (Chinese patterns survive the VBA↔Python boundary), and ~40 audited bug fixes including one that could delete a real Exchange rule

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
2. **Import VBA modules** (Alt+F11 → File → Import): `Config.bas`, `Utilities.bas`, `AgentMemory.bas`, `EmailFilter.bas`, `EmailAgent.bas`, `EmailDigest.bas`, `BatchFilter.bas`, `Bridge.bas`
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
| **Settings** | View and edit all `settings.ini` sections in-browser (API keys masked), save with one click |
| **Macros** | Click-to-run buttons for allowlisted macros with real result output (requires Outlook running) |
| **Learned Rules** | Browse sender rules, subject rules, and reply examples |
| **Digest** | Read the latest daily digest; generate one on demand |
| **Proposals** | Approve/reject LLM-mined rule proposals before they go live |
| **Decisions** | Browse the classification decision log (who/what/why/confidence) |
| **Chat** | Conversational commands: "dry run", "daily digest", "propose rules", "provider claude", etc. |
| **Logs** | Live view of `error.log` and the LLM debug log |

The Web UI reads data files directly. Macro execution uses a file-based command bridge: the server writes a JSON file, the VBA poller (started automatically at Outlook startup) picks it up, executes the macro, and writes the result. Outlook must be running for macros to execute. All API routes require a local auth token (auto-generated, injected into the page — this blocks cross-site request forgery from other websites in your browser).

Requires: Python 3.10+ (settings/rules/logs/digest tabs work on any OS; macros need Outlook on Windows). See [webui/README.md](webui/README.md) for full details.

## Claude Desktop / Claude Code integration (Optional)

The `mcp/` directory contains an MCP (Model Context Protocol) server that exposes the agent to Claude Desktop and Claude Code: run dry-runs, generate the digest, read learned rules and recent decisions, add sender rules, and draft replies in your learned style — all through the same local file bridge, nothing leaves your machine except your configured LLM calls. Destructive bulk operations are deliberately not exposed. Setup snippets: [mcp/README.md](mcp/README.md).

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
│   ├── Utilities.bas              # Helpers, CallLLM (with timeouts), UTF-8 I/O, INI, error handling
│   ├── AgentMemory.bas            # Decision log, sender history, LLM correction capture
│   ├── EmailFilter.bas            # Classification engine + structured LLM classify
│   ├── EmailAgent.bas             # Agent: addressing, auto-reply, reply learning
│   ├── EmailDigest.bas            # Daily triage digest + rule mining
│   ├── BatchFilter.bas            # Interactive macros over headless *Core functions
│   ├── Bridge.bas                 # Web UI/MCP command bridge + scheduler
│   └── ThisOutlookSession.bas     # Event handlers (copy-paste, not import)
├── webui/
│   ├── server.py                  # Flask app — all API routes (token auth)
│   ├── auth.py                    # Local auth token management
│   ├── macros.py                  # Macro manifest: allowlist + arg validation
│   ├── bridge.py                  # File-based command bridge to VBA
│   ├── settings_manager.py        # settings.ini read/write (UTF-8, secrets masked)
│   ├── datafiles.py               # Digest/decisions/proposals readers + approval
│   ├── chat.py                    # Keyword command parser
│   ├── tests/                     # pytest suite (runs on any OS)
│   ├── requirements.txt           # flask
│   ├── static/                    # index.html, style.css, app.js (SPA)
│   └── README.md                  # Web UI setup and API reference
├── mcp/
│   ├── outlook_agent_mcp.py       # MCP stdio server for Claude Desktop/Code (15 tools)
│   ├── tests/                     # pytest suite
│   └── README.md                  # Registration snippets + tool table
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
