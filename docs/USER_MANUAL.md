# Outlook Email Agent v3.0 — User Manual

Complete reference for daily use, configuration, and maintenance.

> **v3.1 update**: this manual covers the v3.0 feature set, which is unchanged. New in v3.1 —
> daily triage digest (`GenerateDailyDigest`), LLM rule mining with Web UI approval
> (`ProposeRules`), confidence-gated structured classification with sender-history context,
> a correction loop (drag-to-learn now also teaches the LLM), an MCP server for Claude
> Desktop/Code (`mcp/README.md`), Web UI token auth + Digest/Proposals/Decisions tabs, and
> three new VBA modules to import (`AgentMemory.bas`, `EmailDigest.bas`, `Bridge.bas`).
> See [README.md](../README.md), [macros.md](macros.md), and [data-files.md](data-files.md)
> for the v3.1 additions; `ShowVersionInfo` should report **v3.1.0** after upgrading.

---

## Table of Contents

1. [Overview](#overview)
2. [Daily Usage](#daily-usage)
3. [Web UI (Optional)](#web-ui-optional)
4. [Self-Improving Filter](#self-improving-filter)
5. [Agent Tools (v3.0)](#agent-tools-v30)
6. [LLM Features](#llm-features)
7. [Settings & Configuration](#settings--configuration)
8. [Folder Setup](#folder-setup)
9. [Classification Rules](#classification-rules)
10. [Server Rule Import & Export](#server-rule-import--export)
11. [Quick Access Toolbar (QAT)](#quick-access-toolbar-qat)
12. [All Available Macros](#all-available-macros)
13. [Dry-Run Icon Reference](#dry-run-icon-reference)
14. [Pattern Configuration](#pattern-configuration)
15. [Migration from v1.x / v2.0](#migration-from-v1x--v20)
16. [Updating the Code](#updating-the-code)
17. [Troubleshooting](#troubleshooting)

---

## Overview

The Outlook Email Agent is a VBA system that classifies incoming emails via a 10-rule priority chain, learns from your manual sorting, and (optionally) uses an LLM to classify ambiguous emails, summarize, and draft replies in your personal style.

Key capabilities:
- **External settings** — configure via `settings.ini` or the browser-based Web UI, no VBA code editing needed
- **Rule-based filtering** — fast, offline, no external dependencies
- **Self-improving** — learns from drag-and-drop sorting into learning folders
- **Reply learning** — learns your reply style from sent emails
- **Multi-provider LLM** — local (Ollama, LM Studio, Inferencer), Azure OpenAI, Anthropic Claude, or OpenAI-compatible (OpenRouter, Groq, etc.)
- **Server rule export** — push learned rules to Exchange for 24/7 filtering
- **Web UI** — optional Python Flask server for browser-based management (settings, macros, learned rules, logs)
- **Structured error logging** — all errors written to `error.log` with call stack

---

## Daily Usage

All operations run from the **Immediate Window** (Alt+F11 then Ctrl+G) or via **QAT buttons**.

### Common Tasks

| Task | Macro |
|------|-------|
| Preview what agent would do | `FilterExistingDryRun` |
| Filter all Inbox emails | `FilterExistingEmails` |
| Filter selected email(s) | `FilterSelectedEmails` |
| Filter current folder | `FilterCurrentFolder` |
| Filter last 7 days | `FilterLastNDays 7` |
| Check version and status | `ShowVersionInfo` |

### Teaching the Agent

Drag emails to the learning folders under Inbox:

| Folder | Agent learns... |
|--------|----------------|
| **LearnKeep** | Always keep from this sender |
| **LearnDelete** | Always delete from this sender |
| **LearnSubjectDelete** | Always delete emails with this subject text |
| **LearnReply** | Your reply style (drag your *sent* replies here) |

The agent learns instantly. No restart needed.

---

## Web UI (Optional)

A Python Flask web server that provides a browser-based interface — an alternative to running macros from the VBA Immediate Window.

### Start the Web UI

```bash
cd webui
pip install -r requirements.txt   # first time only
python server.py
```

Open `http://localhost:5000` in your browser. Outlook must be running for macro commands to execute (the Web UI reads settings/logs/rules directly without Outlook).

### Tabs

| Tab | What you can do |
|-----|----------------|
| **Settings** | View and edit all `settings.ini` sections in-browser, save with one click |
| **Macros** | Click-to-run buttons for all major macros — results display in the page |
| **Learned Rules** | Browse sender rules, subject rules, and reply examples with search |
| **Emails** | Read-only email browser via Outlook COM (Windows only) |
| **Chat** | Type conversational commands — no LLM required, keyword-based |
| **Logs** | Live view of `error.log` with ERROR/WARN highlighting |

### Chat Commands

Type natural language in the Chat tab — the parser recognizes these commands:

| Say | Action |
|-----|--------|
| `dry run` / `preview` | Run `FilterExistingDryRun` |
| `filter inbox` | Run `FilterExistingEmails` |
| `show version` / `status` | Run `ShowVersionInfo` |
| `show senders` | View learned sender rules |
| `show subjects` | View learned subject rules |
| `show errors` | View error log |
| `reload settings` | Send `ReinitializeFilter` to Outlook |
| `provider local` / `provider azure` / `provider claude` / `provider openai` | Switch LLM provider |
| `enable llm` / `disable llm` | Toggle `UseLLMAPI` on/off |
| `scan sent` | Run `ScanSentForReplyPatterns` |
| `help` | List all available commands |

### How the Macro Bridge Works

The Web UI cannot call VBA macros directly (Outlook has no external API for this). Instead:

1. Clicking a macro button sends a JSON command to `%APPDATA%\OutlookEmailFilter\commands\`
2. The VBA command poller (started automatically at Outlook startup) checks that folder every 2 seconds
3. VBA executes the macro and writes a result file
4. The browser polls for the result and displays the output

If Outlook is not running, macro commands time out after 30 seconds. Settings, rules, and logs always work.

### What Works Without Outlook Running

- **Settings tab** — reads and writes `settings.ini` directly
- **Learned Rules tab** — reads all `learned_*.txt` files directly
- **Logs tab** — reads `error.log` directly
- **Chat tab** — setting changes work; macro commands require Outlook

See [webui/README.md](../webui/README.md) for the full API reference.

---

## Self-Improving Filter

Learned rules have the **highest priority** (Rule 0 and 0.5) and override all static patterns.

### Learning by Sender (Rule 0)

| You do this... | The agent learns... |
|----------------|---------------------|
| Drag email to **LearnKeep** | Always **KEEP** from that sender |
| Drag email to **LearnDelete** | Always **DELETE** from that sender |
| Drag to opposite folder | New rule **overwrites** the old one |

**Retroactive rule changes:**
- DELETE → KEEP: automatically rescues that sender's emails from Deleted Items
- KEEP → DELETE: removes that sender's emails from Inbox

### Learning by Subject (Rule 0.5)

| You do this... | The agent learns... |
|----------------|---------------------|
| Drag email to **LearnSubjectDelete** | Always **DELETE** emails containing that subject text |

Subject matching is **case-insensitive substring** — "Marketing" matches "Weekly Marketing Update".

### Bulk Import (First-Time Setup)

If you already have emails in your learning folders and want to backfill all rules:

```
ImportExistingLearnedFolders       ' backfill sender rules from LearnKeep/LearnDelete
ImportExistingLearnedSubjectFolder ' backfill subject rules from LearnSubjectDelete
```

### Managing Learned Rules

| Macro | What It Does |
|-------|-------------|
| `ShowLearnedSendersList` | Dump all sender rules to Immediate Window |
| `ShowLearnedSubjectsList` | Dump all subject rules to Immediate Window |
| `CleanLearnedSendersFile` | Remove duplicate entries from learned_senders.txt |
| `CleanLearnedSubjectsFile` | Remove duplicate entries from learned_subjects.txt |
| `ReloadLearnedSenders` | Force reload from file (after manual edits) |
| `ShowLearnedRepliesSummary` | Show learned reply pair count and file path |

---

## Agent Tools (v3.0)

### Generate Addressing Patterns

If you're setting up the filter for a new name, the LLM can generate all name/greeting variants:

```
GenerateAddressingPatterns
```

Prompts for your full name and title → generates `NamePatterns` and `GreetingPatterns` → saves directly to `settings.ini`. Requires `UseLLMAPI=True`.

### Reply Style Learning

The agent learns your reply style from examples so it can draft contextually appropriate replies.

**Method A — Drag sent replies:**
Drag any of your past sent replies into the **LearnReply** folder under Inbox. The agent extracts the reply pair (original email + your reply) and appends it to `learned_replies.txt`.

**Method B — Scan Sent Items:**
```
ScanSentForReplyPatterns
```
Scans your Sent Items for the last N days (configurable via `ScanSentDays` in settings.ini) and extracts reply pairs automatically. Only processes emails that are replies (start with RE:).

### Draft Replies

Once reply examples are learned:

```
DraftReplyToSelected     ' draft for a single selected email
DraftRepliesForInbox     ' batch draft for all unread KEEP emails in Inbox
```

Drafts are saved to your Outlook **Drafts folder** — never sent automatically. Review before sending.

---

## LLM Features

Enable LLM by setting `UseLLMAPI=True` in `settings.ini` under `[LLM]`.

### Providers

| Provider | `Provider=` value | Auth |
|----------|------------------|------|
| Ollama / LM Studio / Inferencer | `local` | None (or dummy key) |
| Azure OpenAI | `azure` | `api-key` header |
| Anthropic Claude | `claude` | `x-api-key` + `anthropic-version` headers |
| OpenAI-compatible (OpenRouter, Groq, etc.) | `openai` | `Authorization: Bearer` header |

Switch providers by changing `Provider=` in settings.ini. No restart required.

### Summarize Email

1. Select an email in Outlook
2. In Immediate Window: `SummarizeSelectedEmail`
3. Summary appears in a MsgBox (2–3 bullet points)

### Draft Reply

1. Select an email in Outlook
2. In Immediate Window: `DraftReplyToSelected`
3. LLM drafts a reply using your learned style (few-shot examples if available)
4. Draft is displayed in a MsgBox (copy to Drafts as needed)

### LLM Classification

When LLM is enabled, ambiguous emails (Rule 10 — no pattern match) are classified by the LLM instead of going to Review folder.

---

## Settings & Configuration

All settings are stored in `%APPDATA%\OutlookEmailFilter\settings.ini`.

### Editing Settings

1. **Direct edit** — open `settings.ini` in any text editor, save
2. **Web UI** — open the Settings tab at `http://localhost:5000`, edit in-browser, click Save
3. **Reset to defaults** — delete settings.ini and restart Outlook

Pattern changes take effect immediately. Folder name changes require an Outlook restart.

### INI File Structure

```ini
[General]
Version=3.0.0
EnableLogging=True
LogLevel=INFO          ; DEBUG | INFO | WARN | ERROR
EnableSelfImproving=True
DebugMode=False        ; True = MsgBox on every error (debugging)
ProgressInterval=100
DryRunLimit=50
LLMBatchSize=10

[Folders]
Protected=Protected
Review=Review
LearnKeep=LearnKeep
LearnDelete=LearnDelete
LearnSubject=LearnSubjectDelete
LearnReply=LearnReply

[Patterns]
ProtectedDomains=substack.com,reddit.com,redditmail.com
NamePatterns=Xu Xin,XuXin,...
GreetingPatterns=Dear Professor Xu,...
PolyUTags=[MM],[HRO],[CUS],ToXX
VIPSubjectKeywords=thesis,dissertation,...
DeleteSenderPatterns=notice,noreply,...
DeleteKnownSenders=LinkedIn Job Alerts,...
DeleteSubjectPatterns=優惠,offer,digest,...

[LLM]
UseLLMAPI=False
Provider=azure          ; local | azure | claude | openai
AzureEndpoint=https://YOUR-RESOURCE.openai.azure.com/...
LocalEndpoint=http://localhost:11434/v1/chat/completions
LocalModel=qwen3:8b
ClaudeEndpoint=https://api.anthropic.com/v1/messages
ClaudeModel=claude-opus-4-20250115
OpenAIEndpoint=https://openrouter.ai/api/v1/chat/completions
OpenAIModel=qwen/qwen3-8b
APIKeyMethod=ENV        ; ENV | HARDCODED
APIKeyEnvVar=LLM_API_KEY
APIKeyHardcoded=
ClassifyMaxTokens=100
SummarizeMaxTokens=300
ReplyMaxTokens=800
Temperature=0.3
ReplyTemperature=0.7
SystemPrompt=You are filtering emails for...

[Agent]
EnableAutoReply=False
AutoReplyOnArrival=False
LearnReplyFolder=LearnReply
MaxReplyExamples=5
ReplyPersona=
ScanSentItems=False
ScanSentDays=30
AutoReplyForSenders=
```

---

## Folder Setup

Create these folders under Inbox (right-click Inbox → New Folder):

| Folder | Default Name | Purpose | Auto-created? |
|--------|-------------|---------|---------------|
| Protected | `Protected` | Protected domain emails | Yes |
| Review | `Review` | Ambiguous emails for manual triage | Yes |
| Learn Keep | `LearnKeep` | Drag here → always KEEP from that sender | No |
| Learn Delete | `LearnDelete` | Drag here → always DELETE from that sender | No |
| Learn Subject | `LearnSubjectDelete` | Drag here → always DELETE by subject | No |
| Learn Reply | `LearnReply` | Drag sent replies here → learn reply style | No |

Folder names are configurable via the `[Folders]` section of settings.ini.

---

## Classification Rules

The agent checks rules in priority order. **First match wins.**

| Priority | Rule | Action | Source |
|----------|------|--------|--------|
| 0 | **Learned sender** (from LearnKeep/LearnDelete) | KEEP or DELETE | `learned_senders.txt` |
| 0.5 | **Learned subject** (from LearnSubjectDelete) | DELETE | `learned_subjects.txt` |
| 1 | **Protected domain** | Move to Protected | settings.ini |
| 2 | **Personally addressed** (name in subject/body) | KEEP | settings.ini |
| 3 | **Organizational tags** ([MM], [HRO], etc.) | KEEP | settings.ini |
| 4 | **VIP keywords** (thesis, deadline, etc.) | KEEP | settings.ini |
| 5 | **Reply chain** (RE:/AW:) | KEEP | Built-in |
| 6 | **Forward chain** (FW:/FWD:/WG:) | KEEP | Built-in |
| 7 | **Known spam senders** | DELETE | settings.ini |
| 8 | **Spam email patterns** (noreply, marketing) | DELETE | settings.ini |
| 9 | **Spam subject keywords** (newsletter, unsubscribe) | DELETE | settings.ini |
| 10 | **No match** | Move to Review (or LLM) | settings.ini |

---

## Server Rule Import & Export

### Importing Server Rules

Migrate existing server-side Outlook Rules into the VBA agent:

1. In Immediate Window: `ImportServerRules`
2. Confirm when prompted
3. Manually delete server rules afterward: Home → Rules → Manage Rules & Alerts

### Exporting to Server Rules

Push learned DELETE rules to Exchange for 24/7 filtering:

1. In Immediate Window: `ExportLearnedRulesToServer`
2. Confirm when prompted
3. Only DELETE rules are exported (no server equivalent for KEEP)

---

## Quick Access Toolbar (QAT)

Add macros to QAT for one-click access without the VBA Editor:

1. **File → Options → Quick Access Toolbar**
2. Select **Macros** from dropdown
3. Recommended buttons:
   - **FilterSelectedEmails** — filter selected email(s)
   - **FilterCurrentFolder** — filter current folder
   - **FilterExistingDryRun** — dry run preview

---

## All Available Macros

See [macros.md](macros.md) for the complete reference. Key categories:

### Filtering
`FilterExistingDryRun`, `FilterExistingEmails`, `FilterAllFolders`, `FilterSelectedEmail`, `FilterSelectedEmails`, `FilterCurrentFolder`, `FilterLastNDays`, `GenerateClassificationReport`, `BulkDeleteBySender`, `MoveProtectedSources`

### Agent Tools (v3.0)
`GenerateAddressingPatterns`, `ScanSentForReplyPatterns`, `DraftRepliesForInbox`, `ShowLearnedRepliesSummary`

### LLM Tools
`SummarizeSelectedEmail`, `DraftReplyToSelected`

### Learned Rules
`ShowLearnedSenders`, `ShowLearnedSendersList`, `ReloadLearnedSenders`, `CleanLearnedSendersFile`, `ImportExistingLearnedFolders`, `ShowLearnedSubjectsList`, `CleanLearnedSubjectsFile`, `ImportExistingLearnedSubjectFolder`

### Server Rules
`ImportServerRules`, `ExportLearnedRulesToServer`

### Undo / Recovery
`RestoreFromReview`, `RestoreDeletedKeepEmails`

### System
`ThisOutlookSession.ReinitializeFilter`, `ThisOutlookSession.EnableRealTimeFilter`, `ThisOutlookSession.DisableRealTimeFilter`, `DetectAndMigrateOldFolders`, `ShowVersionInfo`

> **Note**: The first three macros live in `ThisOutlookSession` (a document module) and must be called with the full qualified name from the Immediate Window.

---

## Dry-Run Icon Reference

| Icon | Meaning |
|------|---------|
| `[DEL]` | Will be deleted (static rule 7/8/9) |
| `[xLR]` | Will be deleted (learned sender rule) |
| `[xLS]` | Will be deleted (learned subject rule) |
| `[II]` | Will be moved to Protected folder |
| `[OK]` | Will stay in Inbox |
| `[+LR]` | Will stay in Inbox (learned keep rule) |
| `[???]` | Will go to Review folder (or LLM) |

---

## Pattern Configuration

Edit patterns in `settings.ini` under `[Patterns]` section. All patterns are comma-separated strings. Matching is **case-insensitive substring** for all patterns including PolyU tags.

| Pattern Key | What It Controls | Matching |
|-------------|-----------------|----------|
| `ProtectedDomains` | Domains to never delete | Sender's domain contains pattern |
| `NamePatterns` | Your name variations (keep if found) | Subject or body contains |
| `GreetingPatterns` | Personal greetings (keep) | Body starts with pattern |
| `PolyUTags` | Organizational tags (keep) | Subject contains |
| `VIPSubjectKeywords` | Important subject keywords (keep) | Subject contains |
| `DeleteSenderPatterns` | Spam email patterns (delete) | Sender email contains |
| `DeleteKnownSenders` | Spam sender names (delete) | Sender display name contains |
| `DeleteSubjectPatterns` | Spam subject keywords (delete) | Subject contains |

See [PATTERNS.md](PATTERNS.md) for examples and configuration tips.

---

## Migration from v1.x / v2.0

### From v1.x (Roman numeral folder names)

Run the migration macro to rename folders automatically:

```
DetectAndMigrateOldFolders
```

| Old Name | New Default |
|----------|------------|
| I | Review |
| II | Protected |
| III | LearnKeep |
| IIII | LearnDelete |
| V | LearnSubjectDelete |

After migration, create the new `LearnReply` folder manually, then restart Outlook.

### From v2.0

Manually remove each old module (right-click → Remove → No), then re-import the updated `.bas` files via File → Import. Paste the updated `ThisOutlookSession.bas` into the built-in module (Ctrl+A → Delete → paste). Your `settings.ini`, `learned_senders.txt`, and `learned_subjects.txt` are preserved. After upgrade, run `ShowVersionInfo` to confirm v3.0.0 is active. See [INSTALL.md](INSTALL.md) Part 2 for full details.

---

## Updating the Code

### ⚠️ CRITICAL: Stop timers before removing modules

**You MUST stop the command poller and event handlers before removing/reimporting modules.**
The Win32 `SetTimer` command poller fires every 2 seconds. If you remove a module while the timer is active, the callback points to deallocated memory and **Outlook will crash**.

### Module-only changes

1. Open VBA Editor (Alt+F11) → Immediate Window (Ctrl+G)
2. Run: `StopCommandPollerStd` — stops the Win32 timer
3. Run: `ThisOutlookSession.DisableRealTimeFilter` — stops event handlers
   (Must use the full qualified name — `DisableRealTimeFilter` alone won't resolve because it lives in a document module, not a standard module)
4. Right-click the changed module(s) → **Remove** → **No** (don't export)
5. **File → Import File...** → import the updated `.bas` file(s)
6. **Debug → Compile Project** → **Ctrl+S**
7. Run: `ThisOutlookSession.ReinitializeFilter` — restarts event handlers + command poller + reloads settings

### ThisOutlookSession changes

1. Run `StopCommandPollerStd` and `ThisOutlookSession.DisableRealTimeFilter` first (steps 2–3 above)
2. Double-click ThisOutlookSession to open → Ctrl+A → Delete → paste new code
3. Compile, save, **restart Outlook**

---

## Troubleshooting

### Settings not loading

- Check that `%APPDATA%\OutlookEmailFilter\settings.ini` exists
- Run `LoadAllSettings` manually in the Immediate Window
- If corrupt, delete settings.ini and restart Outlook

### "Self-improving filter not active" in log

Learning folders are missing or misnamed. Check:
- Folders exist **directly under Inbox** (not nested)
- Names match what's in `[Folders]` section of settings.ini

### LLM features return empty responses

- Check `UseLLMAPI=True` in settings.ini
- For `local`: confirm Ollama/LM Studio/Inferencer is running on `LocalEndpoint`
- For `claude`: confirm `ANTHROPIC_API_KEY` environment variable is set
- For `azure`: confirm the full deployment URL including `?api-version=...`
- For `openai`: confirm `OpenAIEndpoint` and `OpenAIModel` are set, and your API key is valid
- Set `DebugMode=True` in settings.ini for MsgBox on every error
- Check `%APPDATA%\OutlookEmailFilter\error.log` for detailed error info + call stack

### Need to undo deletions

- Deleted emails go to **Deleted Items** (recoverable for 30 days)
- Run `RestoreDeletedKeepEmails` to rescue wrongly deleted emails
- Run `RestoreFromReview` to restore Review folder emails to Inbox

### Draft replies not appearing

- Drafts are saved to the Outlook **Drafts folder** — check there
- `UseLLMAPI` must be `True` and a provider must be configured
- Run `ShowLearnedRepliesSummary` to verify reply examples are loaded

### Web UI — macro commands time out

- Outlook must be running and the VBA project must be loaded
- The command poller starts automatically at Outlook startup — if it stopped, run `ThisOutlookSession.ReinitializeFilter` in the Immediate Window
- Commands directory: `%APPDATA%\OutlookEmailFilter\commands\` — check that this folder exists and is writable
- Settings, learned rules, and logs always work even without Outlook running

### Web UI — "COM not available"

- The Emails tab requires Windows + pywin32 installed (`pip install pywin32`)
- On macOS or Linux, the Emails tab is disabled but all other tabs work normally

---

## Data Files

See [data-files.md](data-files.md) for file formats, field descriptions, and I/O functions.

| File | Location | Contents |
|------|----------|----------|
| `settings.ini` | `%APPDATA%\OutlookEmailFilter\` | All configurable settings |
| `learned_senders.txt` | `%APPDATA%\OutlookEmailFilter\` | Sender rules (email\|action\|timestamp) |
| `learned_subjects.txt` | `%APPDATA%\OutlookEmailFilter\` | Subject rules (subject\|DELETE\|timestamp) |
| `learned_replies.txt` | `%APPDATA%\OutlookEmailFilter\` | Reply pairs (subject\|from\|body\|reply\|timestamp) |
| `error.log` | `%APPDATA%\OutlookEmailFilter\` | Structured error log with call stacks |

All data files are pipe-delimited, append-only. Last entry per key wins (for senders/subjects).
