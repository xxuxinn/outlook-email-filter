# Macro Reference — Outlook Email Agent v3.1

Full list of callable macros. Assign frequently used ones to the Quick Access Toolbar (QAT) via File → Options → Quick Access Toolbar → Macros.

**Recommended QAT buttons**: `FilterSelectedEmails`, `FilterCurrentFolder`, `FilterExistingDryRun`, `DraftReplyForSelected`, `GenerateDailyDigest`

> **v3.1 note — interactive vs. headless**: every bulk macro now has two halves. The macro listed here is the interactive "button" (confirmation dialog + result popup). A matching `<Name>Core` function runs silently and returns a real result string — that's what the Web UI / MCP bridge calls, so bridge results now report actual counts and honest `ERROR:` messages instead of a hardcoded "completed."

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
| `FilterCurrentFolder` | Filter current folder with confirmation (Review folder: DELETE-only mode, non-DELETE emails stay in Review) |
| `FilterLastNDays 7` | Filter last N days (change 7 to any number) |
| `GenerateClassificationReport` | Count classifications without acting |
| `BulkDeleteBySender "pattern"` | Delete all from matching senders |
| `MoveProtectedSources` | Move protected domain emails to Protected folder |

## Digest & Rule Mining (v3.1)

| Macro | Purpose |
|-------|---------|
| `GenerateDailyDigest` | Build the ranked 24 h triage digest → `digests\digest_YYYY-MM-DD.md` (+ self-email, + deadline Tasks if enabled) |
| `ProposeRules` | LLM mines Review folder + decision log → rule proposals (approve in Web UI) |

> Both also run automatically via the poller scheduler when `EnableDailyDigest` / `EnableRuleMining` are on (`[Digest]` in settings.ini) — daily after `DigestHour`, mining weekly.

## Agent Tools

| Macro | Purpose |
|-------|---------|
| `GenerateAddressingPatterns` | LLM-generates name/greeting patterns from inputted name (bridge variant `GenerateAddressingPatternsStd` takes name/title/role as arguments) |
| `ScanSentForReplyPatterns` | Scans Sent Items for reply pairs → `learned_replies.txt` |
| `DraftReplyForSelected` | Draft few-shot replies for selected email(s) → Drafts folder |
| `ShowLearnedRepliesSummary` | Show learned reply pair count and file path |

## LLM Tools

| Macro | Purpose |
|-------|---------|
| `SummarizeSelectedEmail` | Summarize selected email using LLM |
| `DraftReplyToSelected` | Draft a reply using LLM few-shot engine (delegates to `DraftAutoReply`) |

> **Note**: `SummarizeSelectedEmail` and `DraftReplyToSelected` have bridge-friendly `*Std()` variants in Bridge.bas (`SummarizeSelectedEmailStd`, `DraftReplyToSelectedStd`) that are called automatically from the Web UI command bridge.

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

## Cloud Sync

| Macro | Purpose |
|-------|---------|
| `SyncLearnedRules` | Bidirectional sync of learned senders, subjects, and replies with cloud folder (OneDrive) |

> Configure in `settings.ini` under `[Sync]`: set `EnableCloudSync=True` and `CloudSyncPath` to your OneDrive path. Merges rules using timestamp-based conflict resolution — the later decision wins.

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
