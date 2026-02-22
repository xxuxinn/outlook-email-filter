# Quick Installation Guide

Outlook Email Agent v3.0 — how to install.

## Installation

1. **Enable macros** (File → Options → Trust Center → Trust Center Settings → Macro Settings → "Enable all macros")

2. **Import VBA modules** (File → Import each):
   - `src/Config.bas`
   - `src/Utilities.bas`
   - `src/EmailFilter.bas`
   - `src/EmailAgent.bas`
   - `src/BatchFilter.bas`

3. **Paste ThisOutlookSession**: Open `src/ThisOutlookSession.bas` in a text editor, copy all code, paste into the built-in `ThisOutlookSession` module in VBA Editor.

4. `Debug` → `Compile Project` → `Ctrl+S` → restart Outlook

5. **Create folders** under Inbox: `LearnKeep`, `LearnDelete`, `LearnSubjectDelete`, `LearnReply`

See [docs/INSTALL.md](docs/INSTALL.md) for full details.

---

## What Gets Installed?

### VBA Modules
- `Config.bas` — Configuration constants and runtime variables
- `Utilities.bas` — Helper functions, INI reader/writer, CallLLM (multi-provider)
- `EmailFilter.bas` — Core classification engine (10-rule chain)
- `EmailAgent.bas` — Agent: addressing patterns, auto-reply, reply learning
- `BatchFilter.bas` — Bulk operations and macro launchers
- `ThisOutlookSession` — Event handlers (startup, learn-folder watchers)

### Folders (created under Inbox)
- **Review** — Ambiguous emails for manual review
- **Protected** — Protected domain emails (newsletters, digests)
- **LearnKeep** — Drag here → always KEEP from that sender
- **LearnDelete** — Drag here → always DELETE from that sender
- **LearnSubjectDelete** — Drag here → DELETE future emails matching that subject
- **LearnReply** — Drag sent replies here → agent learns your reply style

### Data Files (at `%APPDATA%\OutlookEmailFilter\`)
- `settings.ini` — All configuration (5 sections: General, Folders, Patterns, LLM, Agent)
- `learned_senders.txt` — Auto-created when you use LearnKeep/LearnDelete
- `learned_subjects.txt` — Auto-created when you use LearnSubjectDelete
- `learned_replies.txt` — Auto-created when you use LearnReply or ScanSentForReplyPatterns
- `error.log` — Structured error log (written only on errors)

---

## LLM Configuration (Optional)

Set `UseLLMAPI=True` in `settings.ini` and choose your provider:

| Provider | Setting | Notes |
|----------|---------|-------|
| **Local** (Ollama, LM Studio, Inferencer) | `Provider=local` | Free, private, runs offline |
| **Azure OpenAI** | `Provider=azure` | Requires Azure subscription |
| **Anthropic Claude** | `Provider=claude` | Requires Anthropic API key |
| **OpenAI-compatible** (OpenRouter, Groq, etc.) | `Provider=openai` | Requires API key, any OpenAI-compatible endpoint |

---

## Web UI (Optional)

A browser-based interface for settings, macros, learned rules, and logs. Requires Python 3.10+.

```bash
cd webui
pip install -r requirements.txt
python server.py
# → open http://localhost:5000
```

Outlook must be running for macro execution. Settings/rules/logs work without Outlook.

---

## Troubleshooting

**Compile errors**: VBA Editor → Tools → References → uncheck any "MISSING" entries.

---

## Uninstall

1. Open VBA Editor (Alt+F11)
2. Right-click each module (`Config`, `Utilities`, `EmailFilter`, `EmailAgent`, `BatchFilter`) → **Remove** → **No** (don't export)
3. Open **ThisOutlookSession** → **Ctrl+A** → **Delete** to clear the code
4. **Ctrl+S** → close VBA Editor → restart Outlook

Data files (`settings.ini`, learned rules) are NOT deleted.

---

For detailed documentation, see [docs/INSTALL.md](docs/INSTALL.md) and [docs/USER_MANUAL.md](docs/USER_MANUAL.md).
