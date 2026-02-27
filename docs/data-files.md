# Data Files — Outlook Email Agent v3.0

All data files live under `%APPDATA%\OutlookEmailFilter\`. The folder is auto-created on first run.

## File Summary

| File | Format | Purpose |
|------|--------|---------|
| `settings.ini` | INI sections | All configurable settings |
| `learned_senders.txt` | Pipe-delimited, append-only | Sender → KEEP/DELETE rules |
| `learned_subjects.txt` | Pipe-delimited, append-only | Subject fragment → DELETE rules |
| `learned_replies.txt` | Pipe-delimited, append-only | Reply style examples for few-shot drafting |
| `error.log` | Pipe-delimited, append-only | Structured error log |
| `llm_debug.log` | Multi-line blocks | LLM request/response debug log (only when LogLevel=DEBUG) |
| `commands/*.json` / `*.result` | JSON | Web UI command bridge (transient, auto-cleaned) |

---

## settings.ini

Created automatically with defaults if missing. Edit with any text editor.

```ini
[General]
Version=3.0.0
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
NamePatterns=Alex Smith,AlexSmith,...
GreetingPatterns=Dear Professor Smith,...
OrgTags=[DEPT1],[DEPT2],[DEPT3],ToAS
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
ClassifyMaxTokens=100
SummarizeMaxTokens=300
ReplyMaxTokens=800
Temperature=0.3
ReplyTemperature=0.7
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
Weekly Digest|DELETE|2026-02-19 14:30:00
```

- Only DELETE action is valid for subject rules
- Lookup is **substring** iteration — any cached key that is a substring of the incoming subject triggers DELETE
- Keys sanitized via `SanitizeSubject()` (strips CR/LF/pipe/null)
- `RecordLearnedSubject(subject, "DELETE")` writes to this file
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
