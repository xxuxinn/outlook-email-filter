# CLAUDE.md

VBA-based email agent for Outlook desktop (Windows only). Classifies emails via a priority rule chain with structured, confidence-gated LLM fallback, learns from user feedback (rules + corrections), drafts replies, generates a daily triage digest, and mines rule proposals. Designed for Professor Xu Xin at PolyU Hong Kong.

## Architecture (v3.1)

```
Config.bas       → DEFAULT_* constants + Runtime* public variables only (no logic)
    ↓
Utilities.bas    → Helpers: string match, robust JSON (ExtractJSONStringValue/NumberValue,
                   EscapeJSON/UnescapeJSON), UTF-8 file I/O (ReadTextFileSmart /
                   WriteTextFileUTF8 / AppendLineUTF8 / SplitLines), logging (size-rotated),
                   INI I/O, learned rules I/O, CallLLM (multi-provider, ServerXMLHTTP with
                   timeouts), error handling (PushCallStack/LogError), cloud sync
AgentMemory.bas  → decision_log.txt writer (RecordDecision), sender-history stats cache
                   (GetSenderContext), replied-to set, LLM correction capture
    ↓
EmailFilter.bas  → ClassifyEmail (10-rule chain, sets lastClassifySource) + ExecuteAction
                   (records every decision, resolves LLM_REVIEW, urgency → Outlook categories)
                   + ClassifyViaLLMEx (structured JSON classify with confidence gate)
EmailAgent.bas   → GenerateAddressingPatterns(+Std), DraftAutoReply (few-shot),
                   ScanSentForReplyPatterns(+Core), shared reply-body extraction helpers
EmailDigest.bas  → GenerateDailyDigestCore (24h ranked digest → digests/*.md + self-email +
                   deadline Tasks) and ProposeRulesCore (rule mining → rule_proposals.txt)
    ↓
BatchFilter.bas  → Interactive macros (confirm + MsgBox) as thin wrappers over headless
                   *Core functions that return honest result strings (used by the bridge)
Bridge.bas       → Command bridge: Win32 SetTimer poller (re-entrancy-guarded, orphan-timer
                   self-kill), DispatchMacroStd allowlist → *Core functions, WriteResultFile,
                   scheduler (daily digest + weekly rule mining — no external cron)
    ↓
ThisOutlookSession.bas → Event handlers ONLY (Application_Startup, ItemAdd watchers for 4 learn
                         folders, inboxItems_ItemAdd real-time filter → ExecuteAction, so
                         real-time uses the SAME LLM path as batch). NOT a regular module —
                         paste into built-in ThisOutlookSession object
```

**Web UI** (`webui/` — Python Flask, runs separately):
```
server.py           → Flask routes; token auth (X-Auth-Token header) on all /api/*
auth.py             → webui_token.txt load/create; token injected into served index.html
macros.py           → single macro manifest (server-side allowlist + arg validation), /api/macros
bridge.py           → writes commands\<id>.json → reads commands\<id>.result (collision-safe
                      8-hex ids, result cleanup, truncated-JSON handling, poller health)
settings_manager.py → settings.ini read/write (utf-8-sig w/ cp950 fallback, section/key
                      validation, APIKeyHardcoded masked as __MASKED__)
datafiles.py        → digest/decisions/proposals readers + proposal approve/reject
chat.py             → keyword→action parser (validated against macros.py manifest)
static/             → SPA (Digest, Proposals, Decisions tabs added in v3.1)
```

**MCP server** (`mcp/` — lets Claude Desktop / Claude Code drive this agent): `outlook_agent_mcp.py`,
a self-contained FastMCP stdio server with 15 tools over the same file bridge + data files
(dry runs, digest, rule reads/writes; no destructive bulk tools). Setup: `mcp/README.md`.

VBA side of bridge (`GetCommandsDir`, `WriteResultFile`, `StartCommandPollerStd` /
`PollForCommandsTimer` / `StopCommandPollerStd`) lives in **Bridge.bas** (moved from Utilities.bas in v3.1).

Visual diagrams (architecture, classification flow, learning flows, command bridge): see [diagrams/README.md](diagrams/README.md) (drawn for v3.0 — bridge internals now live in Bridge.bas).

## Two-Layer Configuration

- `DEFAULT_*` constants in `Config.bas` — compile-time fallbacks, never change at runtime
- `Runtime*` variables in `Config.bas` — loaded from `settings.ini` at startup via `LoadAllSettings`
- Settings file: `%APPDATA%\OutlookEmailFilter\settings.ini` (sections: General, Folders, Patterns, LLM, Agent, Digest, Sync) — **UTF-8 with BOM** as of v3.1 (legacy ANSI files are read via BOM detection and migrated on first write)
- `LoadAllSettings` MUST be the first call in `Application_Startup`
- New settings require: add DEFAULT const + Runtime var + LoadAllSettings entry + CreateDefaultSettingsFile entry (+ webui `macros.py`/settings section allowlist if web-editable)
- Numeric settings are `Long` (not Integer) — values > 32,767 are valid

## Classification Priority (first match wins)

0. Learned sender rule (KEEP or DELETE) — from `learned_senders.txt`
0.5. Learned subject DELETE rule — substring match against `learned_subjects.txt` (patterns auto-generalized; guard: stored pattern must be ≥12 chars AND multi-word, else verbatim subject is stored — prevents one-word substring rules that would over-delete)
1–6. Rule-based KEEP: protected domain, personally addressed, org tags, VIP keywords, RE:, FW:
7–9. Rule-based DELETE: known senders, sender patterns, subject patterns
10. LLM_REVIEW → `ClassifyViaLLMEx` if enabled, else moves to Review folder

**Structured LLM classify (v3.1)**: prompt includes sender history from `decision_log.txt` ("47 emails, 45 deleted"), replied-before signal, and recent user corrections as few-shot examples. Response is JSON `{action, category, urgency 1-5, confidence 0-1, reason}` (robust parse, legacy substring fallback at 0.5 confidence). DELETE below `ConfidenceThreshold` (default 0.6) is demoted to REVIEW. KEEP with urgency ≥4 gets "Urgent" category + high importance; urgency 3 gets "Action" category.

Every executed decision is recorded to `decision_log.txt` by `ExecuteAction` (dry-run/report paths classify only and must NOT record).

`ContainsAny()` is case-insensitive substring matching for all patterns including PolyU tags.

**Review folder special handling**: `FilterCurrentFolder` detects the Review folder (case-insensitive) and switches to DELETE-only mode — non-DELETE emails stay in Review instead of moving to Inbox.

## Daily Digest & Rule Mining (v3.1)

- `GenerateDailyDigestCore` (EmailDigest.bas): last 24 h from Inbox + Review → batched LLM triage (category/urgency/summary/deadline per email, one JSON object per line) → ranked markdown (`Needs action / Worth a look / FYI / Review / filter activity`), conversation-grouped → saved to `digests\digest_YYYY-MM-DD.md`, optionally emailed to self (`DigestSendEmail`) and deadline Tasks created (`EnableTaskExtraction`, drafts only).
- `ProposeRulesCore`: mines Review folder + decision log repeat offenders → LLM proposes sender/subject rules → hard validation (subject rules: DELETE-only, ≥12 chars, multi-word) → `rule_proposals.txt` as PENDING. Approval happens in the Web UI Proposals tab (human-in-the-loop); approval appends to the learned files and reloads VBA caches.
- Scheduling: `Bridge.CheckScheduledJobs` (piggybacks on the 2 s poller, checked ~every 60 s) runs the digest once daily after `DigestHour` and mining weekly. State keys `LastDigestDate`/`LastRuleMiningDate` in `[Digest]`.

## Learning & Corrections

- Drag-to-learn folders (LearnKeep/LearnDelete/LearnSubjectDelete/LearnReply) as before.
- **Correction capture (v3.1)**: if the LLM's last recorded decision for a sender is reversed by a learn-folder drag, the reversal is appended to `llm_corrections.txt` and fed back into future classify prompts (up to 5 most recent).
- `RecordLearnedSender` now SKIPS @-less addresses (unresolved Exchange DNs) instead of recording them.

## Cloud Sync

`SyncLearnedRulesCore` (Utilities.bas) syncs `learned_senders.txt`, `learned_subjects.txt`, `learned_replies.txt` bidirectionally with a cloud folder; `SyncLearnedRules` (MsgBox wrapper) and `SyncLearnedRulesAuto` (silent, on startup/quit) both delegate to it. `[Sync]`: `EnableCloudSync=True`, `CloudSyncPath=<cloud root>`; sync subfolder `<CloudSyncPath>\OutlookEmailFilter\`. Merge: last-timestamp-wins per key; both copies identical after sync. Skips gracefully when the cloud folder is unavailable.

## Multi-Provider LLM

`CallLLM(userPrompt, systemPrompt, maxTokens As Long, [temperature])` in Utilities.bas routes to:
- `"local"` → Ollama/LM Studio OpenAI-compatible endpoint at `RuntimeLocalEndpoint` (no auth)
- `"azure"` → Azure OpenAI at `RuntimeLLMEndpoint` (`api-key` header)
- `"claude"` → Anthropic `/v1/messages` (`x-api-key` + `anthropic-version` headers, different JSON schema)
- `"openai"` → OpenAI-compatible API (OpenRouter, Groq, etc.) at `RuntimeOpenAIEndpoint` (`Bearer` auth)

All providers use `MSXML2.ServerXMLHTTP.6.0` with timeouts (`RequestTimeoutSeconds`, default 60) — a stalled endpoint can no longer freeze Outlook. Responses are parsed with the escape-aware `ExtractJSONStringValue` (tolerates `"content": "..."` spacing, `null`, `\\"` sequences).

`ClassifyBodyChars` (default 800) controls the body preview length; `ClassifyMaxTokens` default is 200 (raised for JSON output). API key: `GetAPIKey()` via `APIKeyMethod=ENV` (`APIKeyEnvVar=LLM_API_KEY`, set system-wide) or `HARDCODED`. `CallAzureOpenAICustom` remains a backwards-compatible wrapper.

## Error Handling Pattern

```vba
Public Sub MyProcedure()
    On Error GoTo PROC_ERR
    PushCallStack "ModuleName.MyProcedure"
    ' ... logic ...
PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "ModuleName", "MyProcedure", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub
```

`LogError` writes to `%APPDATA%\OutlookEmailFilter\error.log` (rotated at 2 MB to `.old`) and optionally shows MsgBox when `RuntimeDebugMode=True`. Headless `*Core`/`*Std` functions return `"ERROR: ..."` strings from their handlers instead of showing dialogs.

## Key VBA Gotchas

- **No block scoping**: `Dim` inside a loop scopes to the procedure. Declare at top.
- **Dictionary CompareMode**: Set `dict.CompareMode = 1` BEFORE first `.Add` or `dict(key) = value`.
- **Reverse iteration**: Required for all loops that delete/move items (`For i = Count To 1 Step -1`).
- **Pre-capture before action**: Capture `.SenderName` and `.Subject` before `.Delete`/`.Move` — object becomes invalid after.
- **Exchange addresses**: Use `GetSenderEmail(mail)` not `.SenderEmailAddress` directly (handles `/O=...` internal format).
- **SanitizeSubject()**: Must strip `vbCr`, `vbLf`, `|`, `Chr(0)` from subjects before Dictionary key or file write.
- **Locale-safe decimal**: `Format(temp, "0.00")` + `Replace(..., ",", ".")` before embedding in JSON strings.
- **Locale-safe Restrict dates**: `Format(d, "ddddd h:nn AMPM")`, never `"mm/dd/yyyy"` (HK is dd/mm).
- **Restrict has no `Like`**: use manual loops or DASL `@SQL=`; Jet `Like` raises "Condition is not valid".
- **File encoding**: settings.ini + all learned/agent data files are UTF-8 with BOM. Always read via `ReadTextFileSmart` and write via `WriteTextFileUTF8`/`AppendLineUTF8` — raw `fso.OpenTextFile` reads mojibake Chinese from UTF-8 files.
- **Document module scope**: Subs in `ThisOutlookSession` must be called with full qualified name from the Immediate Window: `ThisOutlookSession.ReinitializeFilter`.

## Data Files

See [@docs/data-files.md](docs/data-files.md) for file locations, formats, and I/O functions.

## Available Macros

See [@docs/macros.md](docs/macros.md) for the full macro reference table.

## Development Notes

- No build system — import `.bas` files directly in VBA Editor (Alt+F11). Import ALL of: `Config.bas`, `Utilities.bas`, `AgentMemory.bas`, `EmailFilter.bas`, `EmailAgent.bas`, `EmailDigest.bas`, `BatchFilter.bas`, `Bridge.bas`
- `ThisOutlookSession.bas` must be copy-pasted into the built-in module, not imported
- Real-time filtering (`inboxItems_ItemAdd`) is enabled by default; disable via `ThisOutlookSession.DisableRealTimeFilter`

### ⚠️ CRITICAL: Module Update Procedure

**You MUST stop all timers and event handlers before removing/reimporting modules.**
The Win32 `SetTimer` command poller fires every 2 s and will crash Outlook if its callback
code is unloaded mid-flight. Follow this exact order:

1. Open Immediate Window (Ctrl+G)
2. Run: `StopCommandPollerStd` — stops the Win32 timer
3. Run: `ThisOutlookSession.DisableRealTimeFilter` — stops event handlers
   (Must use full qualified name — `DisableRealTimeFilter` alone won't resolve from the Immediate Window because it lives in a document module, not a standard module)
4. Now safe to Remove → Re-import `.bas` modules
5. **Ctrl+S** to save the VBA project
6. Run: `ThisOutlookSession.ReinitializeFilter` — restarts event handlers + command poller + reloads settings

### Verifying After Module Update

After `ReinitializeFilter`, verify in the Immediate Window (Ctrl+G):
```
? RuntimeClassifyBodyChars   → should show current setting (default: 800)
? RuntimeUseLLM              → True if LLM is enabled
? RuntimeLLMProvider         → "local" | "azure" | "claude" | "openai"
? Len(GetAPIKey())           → >0 if API key is resolved
? RuntimeEnableDailyDigest   → True if the daily digest is scheduled
```

Then run `FilterExistingDryRun` to preview classifications, or `FilterSelectedEmail` to test a single email. Run `GenerateDailyDigest` to test the digest end-to-end.

- LLM integration defaults to off (`RuntimeUseLLM = False`); digest/mining/task-extraction also default off
- Learning folders (LearnKeep, LearnDelete, LearnSubjectDelete, LearnReply) must be manually created under Inbox
- Web UI: `cd webui && python server.py` → `http://localhost:5000`
  - All `/api/*` calls need the `X-Auth-Token` header (token at `%APPDATA%\OutlookEmailFilter\webui_token.txt`, injected into the SPA automatically)
  - Command bridge polls `%APPDATA%\OutlookEmailFilter\commands\` every 2 s (Win32 `SetTimer` in Bridge.bas)
  - Bridge timeout is 120 s in Web UI; result files written as UTF-8 via `ADODB.Stream`
- Python tests: `webui/tests/` and `mcp/tests/` (pytest; run on any OS with a temp data dir via `OUTLOOK_FILTER_DATA_DIR`)

## DO NOT

- Do NOT remove/reimport modules without first running `StopCommandPollerStd` and `ThisOutlookSession.DisableRealTimeFilter` — the Win32 timer WILL crash Outlook
- Do NOT add business logic to `Config.bas` — constants and Runtime variables only
- Do NOT call `CallAzureOpenAICustom` in new code — use `CallLLM` instead
- Do NOT assume `.SenderEmailAddress` returns SMTP — always use `GetSenderEmail(mail)`
- Do NOT write to pipe-delimited data files without `SanitizeSubject()` on all text fields
- Do NOT read/write data files with raw FSO text streams — use the UTF-8 helpers (`ReadTextFileSmart`/`WriteTextFileUTF8`/`AppendLineUTF8`)
- Do NOT show MsgBox/InputBox in `*Core`/`*Std` functions — they run headless from the bridge/scheduler
- Do NOT record decisions from dry-run/report paths — only `ExecuteAction` calls `RecordDecision`
- Do NOT change `DEFAULT_*` constants to modify runtime behavior — use `WriteINISetting` to update settings.ini
