# CLAUDE.md

VBA-based email agent for Outlook desktop (Windows only). Classifies emails via a priority rule chain with LLM fallback, learns from user feedback, and drafts replies. Designed for Professor Xu Xin at PolyU Hong Kong.

## Architecture (v3.0)

```
Config.bas       → DEFAULT_* constants + Runtime* public variables only (no logic)
    ↓
Utilities.bas    → All helpers: string match, JSON, logging, INI I/O, learned rules I/O,
                   CallLLM (multi-provider), error handling (PushCallStack/LogError), reply pair I/O,
                   Web UI command bridge (WriteResultFile, command poller via Win32 SetTimer)
    ↓
EmailFilter.bas  → ClassifyEmail (10-rule chain) + ExecuteAction + LLM wrappers
EmailAgent.bas   → Agent features: GenerateAddressingPatterns, DraftAutoReply, ScanSentForReplyPatterns
    ↓
BatchFilter.bas  → Batch macros + thin wrappers calling EmailFilter/EmailAgent
    ↓
ThisOutlookSession.bas → Event handlers ONLY (Application_Startup, ItemAdd watchers for 4 learn folders,
                         + inboxItems_ItemAdd real-time filter)
                         NOT a regular module — paste into built-in ThisOutlookSession object
```

**Web UI** (`webui/` — Python Flask, runs separately):
```
server.py          → Flask routes + orchestration
bridge.py          → writes commands\<id>.json → reads commands\<id>.result
settings_manager.py → settings.ini read/write (configparser)
chat.py            → keyword→action parser
static/            → SPA (index.html + style.css + app.js)
```
VBA side of bridge: `GetCommandsDir()` + `WriteResultFile()` in Utilities.bas,
`StartCommandPollerStd` / `PollForCommandsTimer` / `StopCommandPollerStd` in Utilities.bas.

## Two-Layer Configuration

- `DEFAULT_*` constants in `Config.bas` — compile-time fallbacks, never change at runtime
- `Runtime*` variables in `Config.bas` — loaded from `settings.ini` at startup via `LoadAllSettings`
- Settings file: `%APPDATA%\OutlookEmailFilter\settings.ini` (sections: General, Folders, Patterns, LLM, Agent, Sync)
- `LoadAllSettings` MUST be the first call in `Application_Startup`
- New settings require: add DEFAULT const + Runtime var + LoadAllSettings entry + CreateDefaultSettingsFile entry

## Classification Priority (first match wins)

0. Learned sender rule (KEEP or DELETE) — from `learned_senders.txt`
0.5. Learned subject DELETE rule — substring match against `learned_subjects.txt` (patterns are auto-generalized: unique codes, IDs, and dates are stripped so one rule matches all variants)
1–6. Rule-based KEEP: protected domain, personally addressed, org tags, VIP keywords, RE:, FW:
7–9. Rule-based DELETE: known senders, sender patterns, subject patterns
10. LLM_REVIEW → calls `CallLLM` if enabled, else moves to Review folder

`ContainsAny()` is case-insensitive substring matching for all patterns including PolyU tags.

**Review folder special handling**: `FilterCurrentFolder` detects when run on the Review folder and switches to DELETE-only mode — non-DELETE emails stay in Review instead of moving to Inbox. This prevents ambiguous emails (which fell through to Rule 10) from cycling back to Inbox on re-classification.

## Cloud Sync

`SyncLearnedRules` in Utilities.bas syncs `learned_senders.txt`, `learned_subjects.txt`, and `learned_replies.txt` bidirectionally with a cloud folder (OneDrive, Google Drive, etc.). Configured via `[Sync]` section in settings.ini:
- `EnableCloudSync=True` to enable
- `CloudSyncPath=C:\Users\...\OneDrive - ...` — path to the cloud-synced root folder
- Sync subfolder: `<CloudSyncPath>\OutlookEmailFilter\` (auto-created)

Merge logic: both files are read into dictionaries, conflicts resolved by timestamp (later wins), merged result written to both local and cloud. After sync, both copies are identical. `SyncLearnedRulesAuto` runs silently on `Application_Startup` and `Application_Quit` — if OneDrive is unavailable, it logs a warning and falls back to local rules. Manual `SyncLearnedRules` macro shows a MsgBox summary.

## Multi-Provider LLM (v3.0)

`CallLLM(userPrompt, systemPrompt, maxTokens, [temperature])` in Utilities.bas routes to:
- `"local"` → Ollama/LM Studio/Inferencer OpenAI-compatible endpoint at `RuntimeLocalEndpoint` (no auth)
- `"azure"` → Azure OpenAI at `RuntimeLLMEndpoint` (`api-key` header)
- `"claude"` → Anthropic `/v1/messages` (`x-api-key` + `anthropic-version` headers, different JSON schema)
- `"openai"` → External OpenAI-compatible API (OpenRouter, Groq, custom routers, etc.) at `RuntimeOpenAIEndpoint` (`Bearer` token auth)

`ClassifyBodyChars` (default 800) controls how many characters of email body are sent to the LLM for classification. Configurable in `settings.ini` under `[LLM]`.

API key: `GetAPIKey()` reads from environment variable (`APIKeyMethod=ENV`, `APIKeyEnvVar=LLM_API_KEY`) or hardcoded value (`APIKeyMethod=HARDCODED`). Set environment variables system-wide (not terminal-scoped) so Outlook can read them.

`CallAzureOpenAICustom` is kept as a backwards-compatible wrapper; new code always calls `CallLLM`.

## Error Handling Pattern (v3.0)

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

`LogError` writes to `%APPDATA%\OutlookEmailFilter\error.log` and optionally shows MsgBox when `RuntimeDebugMode=True`.

## Key VBA Gotchas

- **No block scoping**: `Dim` inside a loop scopes to the procedure. Declare at top.
- **Dictionary CompareMode**: Set `dict.CompareMode = 1` BEFORE first `.Add` or `dict(key) = value`.
- **Reverse iteration**: Required for all loops that delete/move items (`For i = Count To 1 Step -1`).
- **Pre-capture before action**: Capture `.SenderName` and `.Subject` before `.Delete`/`.Move` — object becomes invalid after.
- **Exchange addresses**: Use `GetSenderEmail(mail)` not `.SenderEmailAddress` directly (handles `/O=...` internal format).
- **SanitizeSubject()**: Must strip `vbCr`, `vbLf`, `|`, `Chr(0)` from subjects before Dictionary key or file write.
- **Locale-safe decimal**: `Format(temp, "0.00")` + `Replace(..., ",", ".")` before embedding in JSON strings.
- **Document module scope**: Subs in `ThisOutlookSession` (a document module) must be called with full qualified name from the Immediate Window: `ThisOutlookSession.ReinitializeFilter`, not just `ReinitializeFilter`.

## Data Files

See [@docs/data-files.md](docs/data-files.md) for file locations, formats, and I/O functions.

## Available Macros

See [@docs/macros.md](docs/macros.md) for the full macro reference table.

## Development Notes

- No build system — import `.bas` files directly in VBA Editor (Alt+F11)
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
```

Then run `FilterExistingDryRun` to preview classifications, or `FilterSelectedEmail` to test a single email.

- LLM integration defaults to off (`RuntimeUseLLM = False`)
- Learning folders (LearnKeep, LearnDelete, LearnSubjectDelete, LearnReply) must be manually created under Inbox
- Web UI: `cd webui && python server.py` → `http://localhost:5000`
  - Command bridge polls `%APPDATA%\OutlookEmailFilter\commands\` every 2 s
  - Win32 `SetTimer` in Utilities.bas fires `PollerCallback` every 2 s
  - Bridge timeout is 120 s in Web UI; result files written as UTF-8 via `ADODB.Stream`

## DO NOT

- Do NOT remove/reimport modules without first running `StopCommandPollerStd` and `ThisOutlookSession.DisableRealTimeFilter` — the Win32 timer WILL crash Outlook
- Do NOT add business logic to `Config.bas` — constants and Runtime variables only
- Do NOT call `CallAzureOpenAICustom` in new code — use `CallLLM` instead
- Do NOT assume `.SenderEmailAddress` returns SMTP — always use `GetSenderEmail(mail)`
- Do NOT write to pipe-delimited data files without `SanitizeSubject()` on all text fields
- Do NOT change `DEFAULT_*` constants to modify runtime behavior — use `WriteINISetting` to update settings.ini
