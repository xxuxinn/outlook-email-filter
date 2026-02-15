# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

VBA-based email filtering system for Microsoft Outlook desktop (Windows only). Classifies incoming emails using a priority-ordered rule chain with optional Azure OpenAI LLM fallback for ambiguous cases. External settings file, self-improving learned rules, and optional graphical Dashboard. Designed for Professor Xu Xin at PolyU Hong Kong.

## Architecture

Five VBA modules + two optional UserForms:

```
Config.bas          → DEFAULT_* constants (compile-time fallbacks) + Runtime* public variables
                      (loaded from settings.ini at startup). Version constants.
    ↓
Utilities.bas       → Helpers: string matching, JSON encoding/parsing, folder management, logging,
                      email address extraction, INI reader/writer (settings.ini),
                      learned senders/subjects cache (in-memory Dictionary + file I/O),
                      learned rule deletion functions, cache accessor functions
    ↓
EmailFilter.bas     → Core classification engine (ClassifyEmail with Rule 0/0.5 learned rules)
                      + action executor (ExecuteAction) + LLM integration (classify/summarize/reply)
    ↓
BatchFilter.bas     → Bulk operations: dry-run preview, filter inbox/all/selected/current folder,
                      bulk delete by sender, reporting, undo helpers, diagnostics,
                      Dashboard launcher, migration helper, export/version macros
    ↓
ThisOutlookSession.bas → Outlook event handlers (Application_Startup calls LoadAllSettings first,
                         inboxItems_ItemAdd for real-time filtering)
                         + learning folder watchers (learnKeepItems, learnDeleteItems, learnSubjectDeleteItems)
                         NOT a regular module — code is pasted into the built-in ThisOutlookSession object

Optional UserForms (not required for core functionality):
  frmFilterDashboard → 4-tab Dashboard (Filter Actions, Patterns, Settings, Learned Rules)
  frmDraftReply      → LLM draft reply viewer with copy/create reply actions
```

UserForm references use late binding (`VBA.UserForms.Add("formName")`) so the code compiles and runs without the forms installed. `OpenDashboard` and `DraftReplyToSelected` fall back to MsgBox when forms are absent.

## Two-Layer Configuration

All settings use a two-layer system:

1. **`DEFAULT_*` constants** in `Config.bas` — compile-time fallbacks, never change at runtime
2. **`Runtime*` public variables** in `Config.bas` — loaded from `settings.ini` at startup via `LoadAllSettings`

Settings file: `%APPDATA%\OutlookEmailFilter\settings.ini` (INI format with `[General]`, `[Folders]`, `[Patterns]`, `[LLM]` sections). Auto-created with defaults on first run.

**Critical**: `LoadAllSettings` must be the VERY FIRST call in `Application_Startup`, before any folder resolution or event handler setup. The `RuntimeSettingsLoaded` flag gates `LogMessage` behavior before settings are loaded.

### Key Runtime Variables

| Variable | Source INI Key | Purpose |
|----------|---------------|---------|
| `RuntimeFolderProtected` | `[Folders] Protected` | Protected domain folder name |
| `RuntimeFolderReview` | `[Folders] Review` | Ambiguous email folder name |
| `RuntimeFolderLearnKeep` | `[Folders] LearnKeep` | Learn-keep folder name |
| `RuntimeFolderLearnDelete` | `[Folders] LearnDelete` | Learn-delete folder name |
| `RuntimeFolderLearnSubject` | `[Folders] LearnSubject` | Learn-subject-delete folder name |
| `RuntimeProtectedDomains` | `[Patterns] ProtectedDomains` | Comma-separated domain list |
| `RuntimeUseLLM` | `[LLM] UseLLMAPI` | Enable/disable LLM integration |
| `RuntimeEnableSelfImproving` | `[General] EnableSelfImproving` | Enable/disable learning |
| `RuntimeEnableLogging` | `[General] EnableLogging` | Enable/disable logging |

## Classification Priority (first match wins)

0. **Learned sender rule** (self-improving) → KEEP or DELETE (highest priority)
0.5. **Learned subject rule** (self-improving) → DELETE (substring match)
1. Protected domain → Move to Protected folder
2. Personally addressed (name in subject/body, or greeting match) → Keep
3. Organizational tags (e.g. `[MM]`, `[HRO]`) → Keep
4. VIP subject keywords → Keep
5. Reply chain (RE:/AW:) → Keep
6. Forward chain (FW:/FWD:/WG:) → Keep
7. Known spam sender names → Delete
8. Spam sender email patterns (noreply, marketing, etc.) → Delete
9. Spam subject keywords → Delete
10. No match → LLM_REVIEW (LLM call if enabled, otherwise Review folder)

All pattern matching is comma-separated strings checked via `ContainsAny()` (case-insensitive substring match). The `POLYU_TAGS` check is the exception — uses case-sensitive matching.

Rule 0 uses an in-memory `Scripting.Dictionary` cache loaded from `learned_senders.txt`. Users teach the filter by dragging emails into the LearnKeep or LearnDelete folders.

Rule 0.5 uses a separate `Scripting.Dictionary` cache loaded from `learned_subjects.txt`. Users drag emails into the LearnSubjectDelete folder. Lookup uses case-insensitive substring matching (iterates all cached keys).

## Key VBA Patterns

- **Reverse iteration for deletions**: Batch operations iterate `For i = Count To 1 Step -1` because `.Delete` and `.Move` invalidate indices.
- **Pre-capture before action**: `ExecuteAction` captures `SenderName` and `Subject` before calling `.Delete`/`.Move` since the mail object becomes invalid after these operations.
- **Exchange address resolution**: `GetSenderEmail` handles Exchange internal addresses (`/O=...`) by resolving to SMTP via `GetExchangeUser.PrimarySmtpAddress`.
- **Error-safe default**: `ClassifyEmail` returns "KEEP" on any error (safe fallback).
- **`On Error Resume Next` blocks**: Used for optional folder access and Exchange address resolution — always followed by `On Error GoTo 0`.
- **Append-only learned data file**: `RecordLearnedSender` appends to file and updates cache simultaneously. Last entry per sender wins when file is reloaded.
- **Learning folder watchers**: `WithEvents` on LearnKeep/LearnDelete/LearnSubjectDelete folder Items collections. Since the filter never moves emails *to* these folders, every `ItemAdd` event is a manual user action.
- **Retroactive rule reversal**: DELETE→KEEP triggers `RestoreSenderFromDeleted`; KEEP→DELETE triggers `DeleteSenderFromInbox`.
- **Subject sanitization**: `SanitizeSubject()` strips `vbCr`, `vbLf`, `|`, `Chr(0)` from subjects before Dictionary key or file write — Exchange subjects can contain embedded newlines/null chars.
- **INI read-modify-write**: `WriteINISetting` reads all lines into a Collection, finds/replaces the target key, handles missing sections/keys by appending, then rewrites the entire file.

## Available Macros

### Dashboard & Version

| Macro | Purpose |
|-------|---------|
| `OpenDashboard` | Open the Dashboard UserForm (falls back to MsgBox if not installed) |
| `ShowVersionInfo` | Display version, settings paths, and status |

### Filtering

| Macro | Purpose |
|-------|---------|
| `FilterExistingDryRun` | Preview decisions, no changes |
| `FilterExistingEmails` | Filter all Inbox emails |
| `FilterAllFolders` | Filter Inbox + Other + PST archives |
| `FilterSelectedEmail` | Test classification on one selected email |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation |
| `FilterCurrentFolder` | Filter current folder with confirmation |
| `FilterLastNDays 7` | Filter last N days |
| `GenerateClassificationReport` | Count classifications without acting |
| `BulkDeleteBySender "pattern"` | Delete all from matching senders |
| `MoveProtectedSources` | Move protected domain emails to Protected folder |

### LLM Tools

| Macro | Purpose |
|-------|---------|
| `SummarizeSelectedEmail` | Summarize selected email using LLM |
| `DraftReplyToSelected` | Draft a reply using LLM (frmDraftReply if installed, else MsgBox) |

### Learned Rules

| Macro | Purpose |
|-------|---------|
| `ShowLearnedSenders` | Display learned rules count and file path |
| `ShowLearnedSendersList` | Dump all sender rules to Immediate Window |
| `ReloadLearnedSenders` | Force reload learned rules from file |
| `CleanLearnedSendersFile` | Remove duplicate entries from learned senders file |
| `ImportExistingLearnedFolders` | Bulk import senders from LearnKeep/LearnDelete folders |
| `ShowLearnedSubjectsList` | Dump all subject rules to Immediate Window |
| `CleanLearnedSubjectsFile` | Remove duplicate entries from learned subjects file |
| `ImportExistingLearnedSubjectFolder` | Bulk import subjects from LearnSubjectDelete folder |

### Server Rules

| Macro | Purpose |
|-------|---------|
| `ImportServerRules` | Import server-side Outlook Rules as learned DELETE rules |
| `ExportLearnedRulesToServer` | Export learned DELETE rules as server-side Outlook Rules |

### Undo / Recovery

| Macro | Purpose |
|-------|---------|
| `RestoreFromReview` | Move Review folder emails back to Inbox |
| `RestoreDeletedKeepEmails` | Rescue wrongly deleted emails from Deleted Items |

### Migration & System

| Macro | Purpose |
|-------|---------|
| `DetectAndMigrateOldFolders` | Rename v1.x folders (I/II/III/IIII/V) to v2.0 names |
| `ExportAllModules` | Export all VBA modules to Desktop |
| `ReinitializeFilter` | Restart event handlers |
| `EnableRealTimeFilter` | Turn on automatic filtering |
| `DisableRealTimeFilter` | Turn off automatic filtering |

## Quick Access Toolbar (QAT) Integration

Recommended QAT buttons:
- **`FilterSelectedEmails`** — filter selected email(s)
- **`FilterCurrentFolder`** — filter current folder
- **`FilterExistingDryRun`** — dry run preview

Setup: File → Options → Quick Access Toolbar → choose "Macros" from dropdown → add macro.

Context menu events (`ItemContextMenuDisplay` etc.) are deprecated and non-functional in Outlook 2013+; QAT buttons are the recommended pure-VBA approach.

## Development Notes

- No build system — VBA modules are imported directly into Outlook's VBA Editor (Alt+F11)
- `ThisOutlookSession.bas` must be copy-pasted into the built-in module, not imported as a regular module
- UserForms are **optional** — `.frm` files contain code-behind as text; binary `.frx` must be recreated manually in VBA Editor. All functionality works without forms via macros + settings.ini
- `OpenDashboard` and `DraftReplyToSelected` use late binding (`VBA.UserForms.Add`) — compiles without forms, falls back to MsgBox
- Real-time filtering (`inboxItems_ItemAdd`) is commented out by default; uncomment to enable
- LLM integration defaults to off (`RuntimeUseLLM = False`); ambiguous emails go to Review folder
- API keys: ENV method (environment variable) or HARDCODED method — configured via settings.ini `[LLM]` section
- All configuration lives in `settings.ini` — edit with any text editor, no VBA code editing needed
- Uses `Scripting.Dictionary` for statistics tracking, learned senders cache, and learned subjects cache (Windows COM dependency)
- Self-improving data stored at `%APPDATA%\OutlookEmailFilter\learned_senders.txt` and `learned_subjects.txt` (auto-created)
- Learning folders (LearnKeep, LearnDelete, LearnSubjectDelete) must be manually created under Inbox
- `CallAzureOpenAICustom` is the general-purpose LLM call function; `CallAzureOpenAI` is a thin wrapper for classification

## Data Files

| File | Location | Contents |
|------|----------|----------|
| `settings.ini` | `%APPDATA%\OutlookEmailFilter\` | All configurable settings (INI format) |
| `learned_senders.txt` | `%APPDATA%\OutlookEmailFilter\` | Sender rules (email\|action\|timestamp) |
| `learned_subjects.txt` | `%APPDATA%\OutlookEmailFilter\` | Subject rules (subject\|DELETE\|timestamp) |

Data files are pipe-delimited, append-only. Last entry per key wins.
