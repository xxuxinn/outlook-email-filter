# Outlook Email Filter v2.0 -- User Manual

Complete reference for the VBA-based email filtering system.

---

## Table of Contents

1. [Overview](#overview)
2. [Daily Usage](#daily-usage)
3. [Settings & Configuration](#settings--configuration)
4. [Folder Setup](#folder-setup)
5. [Classification Rules](#classification-rules)
6. [Self-Improving Filter](#self-improving-filter)
7. [LLM Features](#llm-features)
8. [Server Rule Import & Export](#server-rule-import--export)
9. [Quick Access Toolbar (QAT)](#quick-access-toolbar-qat)
10. [All Available Macros](#all-available-macros)
11. [Dry-Run Icon Reference](#dry-run-icon-reference)
12. [Pattern Configuration](#pattern-configuration)
13. [Dashboard (Optional)](#dashboard-optional)
14. [Migration from v1.x](#migration-from-v1x)
15. [Updating the Code](#updating-the-code)
16. [Troubleshooting](#troubleshooting)

---

## Overview

This VBA filter automatically classifies incoming Outlook emails using a priority-ordered rule chain. Emails are either kept in Inbox, deleted, moved to a Protected folder, or sent for review.

Key capabilities:
- **External settings** -- configure via `settings.ini`, no code editing needed
- **Rule-based filtering** -- fast, offline, no external dependencies
- **Self-improving** -- learns from your manual sorting (drag to learning folders)
- **LLM integration** -- summarize emails, draft replies, classify ambiguous emails
- **Server rule export** -- push learned rules to Exchange for 24/7 filtering
- **Batch processing** -- filter existing emails in bulk, with dry-run preview
- **Dashboard UI** (optional) -- 4-tab graphical interface for all operations

---

## Daily Usage

All operations are run from the **Immediate Window** (Alt+F11 then Ctrl+G) or via **QAT buttons**.

### Common Tasks

| Task | Macro |
|------|-------|
| Preview what filter would do | `FilterExistingDryRun` |
| Filter all Inbox emails | `FilterExistingEmails` |
| Filter selected email(s) | `FilterSelectedEmails` |
| Filter current folder | `FilterCurrentFolder` |
| Filter last 7 days | `FilterLastNDays 7` |
| Check version and status | `ShowVersionInfo` |

### Teaching the Filter

Drag emails to the learning folders under Inbox:
- **LearnKeep** -- always keep from this sender
- **LearnDelete** -- always delete from this sender
- **LearnSubjectDelete** -- always delete emails with this subject

The filter learns instantly. No restart needed.

---

## Settings & Configuration

All settings are stored in `%APPDATA%\OutlookEmailFilter\settings.ini`.

### INI File Structure

```ini
[General]
Version=2.0.0
EnableLogging=True
LogLevel=INFO
EnableSelfImproving=True
ProgressInterval=100
DryRunLimit=50
LLMBatchSize=10

[Folders]
Protected=Protected
Review=Review
LearnKeep=LearnKeep
LearnDelete=LearnDelete
LearnSubject=LearnSubjectDelete

[Patterns]
ProtectedDomains=substack.com,reddit.com,...
NamePatterns=Xu Xin,XuXin,...
...

[LLM]
UseLLMAPI=False
Endpoint=https://YOUR-RESOURCE.openai.azure.com/...
APIKeyMethod=ENV
APIKeyEnvVar=AZURE_OPENAI_KEY
MaxTokens=100
Temperature=0.3
SystemPrompt=You are filtering emails for...
```

### Editing Settings

1. **Direct edit** (recommended) -- Open settings.ini in any text editor, save
2. **Dashboard** -- If the optional UserForm is installed, run `OpenDashboard`
3. **Reset to defaults** -- Delete settings.ini and restart Outlook

Pattern changes take effect immediately. Folder name changes require an Outlook restart.

---

## Folder Setup

Create these folders manually under Inbox (right-click Inbox -> New Folder):

| Folder | Default Name | Purpose | Auto-created? |
|--------|-------------|---------|---------------|
| Protected | `Protected` | Protected domain emails | Yes |
| Review | `Review` | Ambiguous emails for manual triage | Yes |
| Learn Keep | `LearnKeep` | Drag here to always KEEP from that sender | No |
| Learn Delete | `LearnDelete` | Drag here to always DELETE from that sender | No |
| Learn Subject | `LearnSubjectDelete` | Drag here to always DELETE by subject | No |

Folder names are configurable via the `[Folders]` section of settings.ini.

---

## Classification Rules

The filter checks rules in priority order. **First match wins**.

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

## Self-Improving Filter

The filter learns from your manual sorting. Learned rules have the **highest priority** and override all static patterns.

### Learning by Sender (Rule 0)

| You do this... | The filter learns... |
|----------------|---------------------|
| Drag email to **LearnKeep** | Always **KEEP** from that sender |
| Drag email to **LearnDelete** | Always **DELETE** from that sender |
| Drag to opposite folder | New rule **overwrites** the old one |

**Retroactive rule changes:**
- DELETE -> KEEP: automatically rescues that sender's emails from Deleted Items
- KEEP -> DELETE: removes that sender's emails from Inbox

### Learning by Subject (Rule 0.5)

| You do this... | The filter learns... |
|----------------|---------------------|
| Drag email to **LearnSubjectDelete** | Always **DELETE** emails containing that subject text |

Subject matching is **case-insensitive substring**.

### Managing Learned Rules

| Macro | What It Does |
|-------|-------------|
| `ShowLearnedSendersList` | Dump all sender rules to Immediate Window |
| `ShowLearnedSubjectsList` | Dump all subject rules to Immediate Window |
| `CleanLearnedSendersFile` | Remove duplicate entries |
| `CleanLearnedSubjectsFile` | Remove duplicate entries |
| `ReloadLearnedSenders` | Force reload from file |

---

## LLM Features

Enable LLM by setting `UseLLMAPI=True` in settings.ini under `[LLM]`.

### Summarize Email

1. Select an email in Outlook
2. In Immediate Window: `SummarizeSelectedEmail`
3. LLM returns a 2-3 bullet point summary in a MsgBox

### Draft Reply

1. Select an email in Outlook
2. In Immediate Window: `DraftReplyToSelected`
3. LLM drafts a professional reply shown in a MsgBox
4. If the optional frmDraftReply UserForm is installed, you also get "Copy to Clipboard" and "Create Reply Email" buttons

### Classification

When LLM is enabled, ambiguous emails (Rule 10) are classified by the LLM instead of going to the Review folder.

---

## Server Rule Import & Export

### Importing Server Rules

Migrate existing server-side Outlook Rules into the VBA filter:

1. In Immediate Window: `ImportServerRules`
2. Confirm when prompted
3. Manually delete the server rules afterward via: Home -> Rules -> Manage Rules & Alerts

### Exporting to Server Rules

Push learned DELETE rules to Exchange for 24/7 filtering:

1. In Immediate Window: `ExportLearnedRulesToServer`
2. Confirm when prompted
3. Only DELETE rules are exported (no server equivalent for KEEP)

---

## Quick Access Toolbar (QAT)

Add macros to QAT for one-click access without the VBA Editor:

1. **File -> Options -> Quick Access Toolbar**
2. Select **Macros** from dropdown
3. Recommended buttons:
   - **FilterSelectedEmails** -- filter selected email(s)
   - **FilterCurrentFolder** -- filter current folder
   - **FilterExistingDryRun** -- dry run preview

---

## All Available Macros

### Filtering

| Macro | What It Does |
|-------|-------------|
| `FilterExistingDryRun` | Preview decisions (no changes) |
| `FilterExistingEmails` | Filter all Inbox emails |
| `FilterAllFolders` | Filter Inbox + Other + PST |
| `FilterSelectedEmail` | Test one selected email |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation |
| `FilterCurrentFolder` | Filter current folder with confirmation |
| `FilterLastNDays 7` | Filter last N days |

### LLM Tools

| Macro | What It Does |
|-------|-------------|
| `SummarizeSelectedEmail` | Summarize selected email using LLM |
| `DraftReplyToSelected` | Draft a reply using LLM |

### Learned Rules

| Macro | What It Does |
|-------|-------------|
| `ShowLearnedSenders` | Show rule count and file path |
| `ShowLearnedSendersList` | Dump all sender rules to Immediate Window |
| `ReloadLearnedSenders` | Force reload from file |
| `CleanLearnedSendersFile` | Remove duplicates |
| `ImportExistingLearnedFolders` | Bulk import from learning folders |
| `ShowLearnedSubjectsList` | Dump all subject rules |
| `CleanLearnedSubjectsFile` | Remove duplicates |
| `ImportExistingLearnedSubjectFolder` | Bulk import subjects |

### Server Rules

| Macro | What It Does |
|-------|-------------|
| `ImportServerRules` | Import server-side rules as learned DELETE rules |
| `ExportLearnedRulesToServer` | Export learned DELETE rules to server |

### Dashboard & Version

| Macro | What It Does |
|-------|-------------|
| `OpenDashboard` | Open the Dashboard UserForm (if installed) |
| `ShowVersionInfo` | Display version, settings paths, and status |

### Migration & Portability

| Macro | What It Does |
|-------|-------------|
| `DetectAndMigrateOldFolders` | Rename v1.x folders to v2.0 names |
| `ExportAllModules` | Export all VBA modules to Desktop |

### Undo / Recovery

| Macro | What It Does |
|-------|-------------|
| `RestoreFromReview` | Move Review folder emails back to Inbox |
| `RestoreDeletedKeepEmails` | Rescue wrongly deleted emails |

### System

| Macro | What It Does |
|-------|-------------|
| `ReinitializeFilter` | Restart event handlers |
| `EnableRealTimeFilter` | Turn on automatic filtering |
| `DisableRealTimeFilter` | Turn off automatic filtering |

---

## Dry-Run Icon Reference

| Icon | Meaning |
|------|---------|
| `[DEL]` | Will be deleted (static rule) |
| `[xLR]` | Will be deleted (learned sender rule) |
| `[xLS]` | Will be deleted (learned subject rule) |
| `[II]` | Will be moved to Protected folder |
| `[OK]` | Will stay in Inbox |
| `[+LR]` | Will stay in Inbox (learned keep rule) |
| `[???]` | Will go to Review folder (or LLM) |

---

## Pattern Configuration

Edit patterns in settings.ini under `[Patterns]` section.

All patterns are comma-separated strings. Matching is **case-insensitive substring** except PolyU Tags (case-sensitive).

| Pattern | What It Controls | Matching |
|---------|-----------------|----------|
| ProtectedDomains | Domains to never delete | Sender's domain contains pattern |
| NamePatterns | Your name variations (keep if found) | Subject or body contains |
| GreetingPatterns | Personal greetings (keep) | Body starts with pattern |
| PolyUTags | Organizational tags (keep) | Subject contains (**case-sensitive**) |
| VIPSubjectKeywords | Important subject keywords (keep) | Subject contains |
| DeleteSenderPatterns | Spam email patterns (delete) | Sender email contains |
| DeleteKnownSenders | Spam sender names (delete) | Sender name contains |
| DeleteSubjectPatterns | Spam subject keywords (delete) | Subject contains |

---

## Dashboard (Optional)

A 4-tab graphical Dashboard UserForm is available for users who prefer a GUI over the Immediate Window. Since VBA UserForms require manual creation in the VBA Editor, this is **optional** -- all functionality is accessible via macros and settings.ini.

If installed, run `OpenDashboard` to access:

| Tab | Features |
|-----|----------|
| **Filter Actions** | Dry run, filter inbox/all/selected/current folder, import/export server rules, LLM tools |
| **Patterns** | Browse, add, edit, remove filter patterns with category dropdown |
| **Settings** | Toggle logging/self-improving/LLM, configure folder names, LLM endpoint |
| **Learned Rules** | View/search/delete learned sender and subject rules |

See [INSTALL.md](INSTALL.md) for UserForm setup instructions.

---

## Migration from v1.x

### Automatic Migration

Run `DetectAndMigrateOldFolders` to automatically rename old folders:

| Old Name | New Default |
|----------|------------|
| I | Review |
| II | Protected |
| III | LearnKeep |
| IIII | LearnDelete |
| V | LearnSubjectDelete |

After migration, restart Outlook to refresh event handlers.

### Manual Migration

1. Update VBA modules (remove old, import new)
2. Rename folders manually in Outlook
3. Delete `settings.ini` to regenerate with defaults (or create one manually)
4. Your `learned_senders.txt` and `learned_subjects.txt` files are preserved automatically

---

## Updating the Code

### Module-only changes (no restart needed)

1. Open VBA Editor (Alt+F11)
2. Right-click the changed module(s) -> **Remove** -> **No** (don't export)
3. **File -> Import File...** -> import the updated `.bas` file(s)
4. **Debug -> Compile Project** -> **Ctrl+S**

### ThisOutlookSession changes

1. Double-click to open -> Ctrl+A -> Delete -> paste new code
2. Compile, save, **restart Outlook**

---

## Troubleshooting

### Settings not loading

- Check that `%APPDATA%\OutlookEmailFilter\settings.ini` exists
- Run `LoadAllSettings` manually in the Immediate Window
- If corrupt, delete settings.ini and restart Outlook

### "Self-improving filter not active" in log

Learning folders are missing. Check:
- Folders exist **directly under Inbox** (not nested)
- Names match what's in settings.ini `[Folders]` section

### LLM features return empty responses

- Check that `UseLLMAPI=True` in settings.ini
- Verify the API endpoint URL is correct
- Ensure the API key environment variable is set
- Check the Immediate Window for error messages

### Need to undo deletions?

- Deleted emails go to **Deleted Items** (recoverable for 30 days)
- Run `RestoreDeletedKeepEmails` to rescue wrongly deleted emails
- Run `RestoreFromReview` to restore Review folder emails

---

## Data Files

| File | Location | Contents |
|------|----------|----------|
| `settings.ini` | `%APPDATA%\OutlookEmailFilter\` | All configurable settings |
| `learned_senders.txt` | `%APPDATA%\OutlookEmailFilter\` | Sender rules (email\|action\|timestamp) |
| `learned_subjects.txt` | `%APPDATA%\OutlookEmailFilter\` | Subject rules (subject\|DELETE\|timestamp) |

Data files are pipe-delimited, append-only. Last entry per key wins.
