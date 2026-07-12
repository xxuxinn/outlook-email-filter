# Data Files — Outlook Email Agent v3.1

All data files live under `%APPDATA%\OutlookEmailFilter\`. The folder is auto-created on first run.

**Encoding (v3.1)**: `settings.ini` and all pipe-delimited data files are **UTF-8 with BOM**. Legacy ANSI files are detected by their missing BOM and migrated to UTF-8 on first write. VBA reads/writes them via `ReadTextFileSmart` / `WriteTextFileUTF8` / `AppendLineUTF8` (Utilities.bas); Python reads `utf-8-sig` with a `cp950`/`latin-1` fallback for legacy files. This fixes Chinese patterns (e.g. `優惠`) corrupting across the VBA/Python boundary.

## File Summary

| File | Format | Purpose |
|------|--------|---------|
| `settings.ini` | INI sections | All configurable settings |
| `learned_senders.txt` | Pipe-delimited, append-only | Sender → KEEP/DELETE rules |
| `learned_subjects.txt` | Pipe-delimited, append-only | Subject fragment → DELETE rules |
| `learned_replies.txt` | Pipe-delimited, append-only | Reply style examples for few-shot drafting |
| `decision_log.txt` | Pipe-delimited, append-only | Every executed classification (v3.1) |
| `llm_corrections.txt` | Pipe-delimited, append-only | User reversals of LLM decisions → few-shot corrections (v3.1) |
| `rule_proposals.txt` | Pipe-delimited, rewritten on status change | LLM-mined rule proposals awaiting approval (v3.1) |
| `digests/digest_YYYY-MM-DD.md` | Markdown | Daily triage digest, one file per day (v3.1) |
| `webui_token.txt` | 32 hex chars | Web UI auth token (created by server.py) |
| `error.log` | Pipe-delimited, append-only | Structured error log (rotated to `.old` at 2 MB) |
| `llm_debug.log` | Multi-line blocks | LLM request/response debug log (only when LogLevel=DEBUG; rotated at 5 MB) |
| `commands/*.json` / `*.result` | JSON | Web UI/MCP command bridge (transient, auto-cleaned) |

---

## decision_log.txt (v3.1)

**Format**: `timestamp|senderEmail|subject(≤80 chars)|source|action|confidence`
```
2026-07-12 09:15:00|noreply@example.com|Weekly Update|RULE8_SENDER_PATTERN|DELETE|1.00
2026-07-12 09:16:30|student@connect.polyu.hk|Question about thesis|LLM|KEEP|0.92
```

- `source`: `LEARNED_SENDER`, `LEARNED_SUBJECT`, `RULE1_PROTECTED` … `RULE9_SUBJECT_PATTERN`, `LLM`, `DEFAULT`
- `action`: `KEEP` | `DELETE` | `MOVE_II` | `REVIEW`; `confidence` is 1.00 for deterministic rules
- Written ONLY by `ExecuteAction` (dry-run/report paths never record)
- Powers: sender-history context in LLM prompts (`GetSenderContext`), the digest's "filter activity" section, rule mining evidence, and the Web UI Decisions tab
- I/O: `RecordDecision` / `LoadSenderStats` / `GetLastDecisionForSender` in AgentMemory.bas

## llm_corrections.txt (v3.1)

**Format**: `timestamp|senderEmail|subject|wrongAction|correctAction`

- Appended when a learn-folder drag reverses the LLM's most recent decision for that sender (e.g. LLM said DELETE, user dragged to LearnKeep)
- The 5 most recent corrections are injected into every LLM classification prompt as few-shot examples
- I/O: `RecordCorrection` / `GetRecentCorrectionsBlock` in AgentMemory.bas

## rule_proposals.txt (v3.1)

**Format**: `id|type|value|action|reason|status|timestamp` where `id` = 8 lowercase hex chars, `type` ∈ SENDER/SUBJECT, `action` ∈ KEEP/DELETE, `status` ∈ PENDING/APPROVED/REJECTED
```
0c91a2f3|SENDER|newsletter@vendor.com|DELETE|Deleted 6 times, never kept|PENDING|2026-07-12 08:05:00
```

- Written by `ProposeRulesCore` (EmailDigest.bas) — weekly via the scheduler or on demand
- Approved/rejected in the Web UI Proposals tab; approval appends to the learned files and reloads VBA caches
- Subject proposals are validated hard: DELETE-only, ≥12 chars, multi-word (substring matching makes short patterns dangerous)

---

## settings.ini

Created automatically with defaults if missing. Edit with any text editor.

```ini
[General]
Version=3.1.0
EnableLogging=True
LogLevel=INFO          ; DEBUG | INFO | WARN | ERROR
EnableSelfImproving=True
DebugMode=False        ; True = MsgBox on every LogError call
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
ClassifyBodyChars=800
ClassifyMaxTokens=200      ; raised in v3.1 for structured JSON output
SummarizeMaxTokens=300
ReplyMaxTokens=800
Temperature=0.3
ReplyTemperature=0.7
RequestTimeoutSeconds=60   ; HTTP receive timeout for LLM calls (v3.1)
ConfidenceThreshold=0.60   ; LLM DELETEs below this confidence go to Review (v3.1)
SystemPrompt=You are filtering emails for...

[Agent]
EnableAutoReply=False
AutoReplyOnArrival=False   ; draft reply on every new KEEP email (requires real-time filter enabled)
LearnReplyFolder=LearnReply
MaxReplyExamples=5
ReplyPersona=              ; blank = auto-generated from name patterns
ScanSentItems=False
ScanSentDays=30
AutoReplyForSenders=       ; blank = draft for all KEEP emails; comma-separated senders to restrict
EnableTaskExtraction=False ; create draft Outlook Tasks from digest deadlines (v3.1)
EnableContextEnrichment=True ; inject sender history into LLM prompts (v3.1)

[Digest]
EnableDailyDigest=False    ; daily ranked triage digest (v3.1)
DigestHour=8               ; digest runs once per day after this hour (0-23)
DigestMaxEmails=50
DigestSendEmail=True       ; also send the digest as a self-addressed email
EnableRuleMining=False     ; weekly LLM rule proposals (approve in Web UI)
LastDigestDate=            ; state, managed by the scheduler
LastRuleMiningDate=        ; state, managed by the scheduler

[Sync]
EnableCloudSync=False
; Path to shared cloud folder (OneDrive, Google Drive, etc.)
CloudSyncPath=C:\Users\xxuxinn\OneDrive - The Hong Kong Polytechnic University
```

---

## learned_senders.txt

**Format**: one rule per line, pipe-delimited
```
email@domain.com|KEEP|2026-02-19 14:30:00
spam@example.com|DELETE|2026-02-19 15:00:00
```

- Append-only; last entry per sender wins when file is reloaded
- Keys are case-insensitive (Dictionary CompareMode=1)
- `RecordLearnedSender(email, action)` writes to this file
- `LoadLearnedSenders([forceReload])` populates the in-memory cache
- `DeduplicateLearnedSenders` rewrites file keeping only last entry per sender
- `DeleteLearnedSenderRule(email)` filters out all entries for that sender

---

## learned_subjects.txt

**Format**: one rule per line, pipe-delimited
```
Funding Application Submitted For Your Information Only|DELETE|2026-03-14 10:00:00
Weekly Digest|DELETE|2026-02-19 14:30:00
```

- Only DELETE action is valid for subject rules
- **Smart pattern extraction**: `RecordLearnedSubject` calls `ExtractSubjectPattern()` to strip variable parts (unique codes, reference IDs, dates) before storing. For example, dragging an email with subject "Funding Application Submitted For Your Information Only (A0061323)" stores just "Funding Application Submitted For Your Information Only" — matching all future emails with any code.
- Patterns stripped: `(A0061323)`, `[TICKET-4521]`, `INV-2026-0042`, `#WX-98234`, standalone 5+ digit numbers, dates. Org tags like `[MM]`, `[HRO]` (2-4 uppercase letters) are preserved.
- Duplicate patterns are skipped (no redundant rules from dragging multiple emails of the same type)
- Falls back to verbatim subject if pattern extraction strips too much (< 8 chars)
- Lookup is **substring** iteration — any cached key that is a substring of the incoming subject triggers DELETE
- Keys sanitized via `SanitizeSubject()` (strips CR/LF/pipe/null)
- `RecordLearnedSubject(subject, "DELETE")` extracts pattern and writes to this file
- `ExtractSubjectPattern(subject)` returns generalized pattern (uses VBScript.RegExp)
- `LookupLearnedSubject(subject)` returns "DELETE" or "" (iterates all keys)

---

## learned_replies.txt

**Format**: one pair per line, pipe-delimited
```
Re: Meeting Request|John Smith|Dear Prof, could we meet...|Thank you for reaching out...|2026-02-19 14:30:00
```

Fields: `original_subject|original_from|original_body_snippet|reply_body_snippet|timestamp`

- All fields sanitized via `SanitizeSubject()` / `SanitizeSnippet()` before writing
- `original_body_snippet` = first 500 chars of original email
- `reply_body_snippet` = first 1000 chars of user's reply
- Append-only; `LoadRecentReplyPairs(n)` returns the last N lines
- Populated by: dragging sent replies into LearnReply folder, or running `ScanSentForReplyPatterns`

---

## error.log

**Format**: pipe-delimited structured log
```
2026-02-19 14:30:00|EmailFilter.ClassifyEmail|13|Type mismatch||Stack: Application_Startup -> FilterInbox -> ClassifyEmail
```

Fields: `timestamp|module.procedure|errNum|errDesc|[context]|Stack: ...`

- Written by `LogError(moduleName, procName, errNum, errDesc, [context])`
- All errors also go to `Debug.Print` (VBA Immediate Window)
- Enable `DebugMode=True` in settings.ini to also get a MsgBox on each error

---

## llm_debug.log

**Format**: multi-line blocks separated by `========================================`
```
======================================== 2026-02-19 14:30:00
Provider: azure | Model: gpt-4o | Endpoint: https://...
--- REQUEST ---
{"messages":[...]}
--- RESPONSE ---
{"choices":[...]}
--- PARSED CONTENT ---
KEEP
========================================
```

- Only written when `LogLevel=DEBUG` in settings.ini
- Each block contains: timestamp, provider/model/endpoint, raw request JSON, raw response JSON, parsed content
- Written by `WriteLLMDebugLog()` in Utilities.bas
- Path from `GetLLMDebugLogPath()` → `%APPDATA%\OutlookEmailFilter\llm_debug.log`
- Viewable in Web UI via the Logs tab
