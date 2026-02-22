# Macro Reference — Outlook Email Agent v3.0

Full list of callable macros. Assign frequently used ones to the Quick Access Toolbar (QAT) via File → Options → Quick Access Toolbar → Macros.

**Recommended QAT buttons**: `FilterSelectedEmails`, `FilterCurrentFolder`, `FilterExistingDryRun`

## Version

| Macro | Purpose |
|-------|---------|
| `ShowVersionInfo` | Display version, settings paths, and status |

## Filtering

| Macro | Purpose |
|-------|---------|
| `FilterExistingDryRun` | Preview decisions without making changes |
| `FilterExistingEmails` | Filter all Inbox emails |
| `FilterAllFolders` | Filter Inbox + Other + PST archives |
| `FilterSelectedEmail` | Test classification on one selected email (prompts before acting) |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation |
| `FilterCurrentFolder` | Filter current folder with confirmation |
| `FilterLastNDays 7` | Filter last N days (change 7 to any number) |
| `GenerateClassificationReport` | Count classifications without acting |
| `BulkDeleteBySender "pattern"` | Delete all from matching senders |
| `MoveProtectedSources` | Move protected domain emails to Protected folder |

## Agent Tools (v3.0)

| Macro | Purpose |
|-------|---------|
| `GenerateAddressingPatterns` | LLM-generates name/greeting patterns from inputted name |
| `ScanSentForReplyPatterns` | Scans Sent Items for reply pairs → `learned_replies.txt` |
| `DraftRepliesForInbox` | Batch draft replies for unread KEEP emails in Inbox |
| `ShowLearnedRepliesSummary` | Show learned reply pair count and file path |

## LLM Tools

| Macro | Purpose |
|-------|---------|
| `SummarizeSelectedEmail` | Summarize selected email using LLM |
| `DraftReplyToSelected` | Draft a reply using LLM (few-shot style if replies learned) |

> **Note**: `SummarizeSelectedEmail` and `DraftReplyToSelected` have bridge-friendly `*Std()` variants in Utilities.bas (`SummarizeSelectedEmailStd`, `DraftReplyToSelectedStd`) that are called automatically from the Web UI command bridge.

## Learned Rules

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
| `ShowLearnedRepliesSummary` | Show learned reply pair count |

## Server Rules

| Macro | Purpose |
|-------|---------|
| `ImportServerRules` | Import server-side Outlook Rules as learned DELETE rules |
| `ExportLearnedRulesToServer` | Export learned DELETE rules as server-side Outlook Rules |

## Undo / Recovery

| Macro | Purpose |
|-------|---------|
| `RestoreFromReview` | Move Review folder emails back to Inbox |
| `RestoreDeletedKeepEmails` | Rescue wrongly deleted emails from Deleted Items |

## Migration & System

| Macro | Purpose |
|-------|---------|
| `DetectAndMigrateOldFolders` | Rename v1.x folders (I/II/III/IIII/V) to v2.0 names |
| `ThisOutlookSession.ReinitializeFilter` | Restart event handlers + command poller + reload settings |
| `ThisOutlookSession.EnableRealTimeFilter` | Re-enable filtering + command poller |
| `ThisOutlookSession.DisableRealTimeFilter` | Disable filtering + command poller |

> **Note**: These three macros live in `ThisOutlookSession` (a document module), so they **must** be called with the full qualified name from the Immediate Window. `ReinitializeFilter` alone will give "Sub or Function not defined."
