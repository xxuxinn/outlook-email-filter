# Installation Guide — Outlook Email Agent v3.0

This guide covers fresh install, upgrade from v2.0, and migration to a new PC.

---

## Prerequisites

- Microsoft Outlook **desktop** (Windows only — not the web app, not "New Outlook")
- Microsoft 365 or Outlook 2016/2019/2021
- The `.bas` source files from `outlook-email-filter/src/`

---

## Part 1: Enable Macros

1. Open **Outlook**
2. Go to **File → Options → Trust Center → Trust Center Settings...**
3. Select **Macro Settings** from the left panel
4. Choose **"Enable all macros"** (or "Notifications for digitally signed macros" if you sign your code)
5. Click **OK** twice → **Restart Outlook**

---

## Part 2: Import VBA Code

#### Step 1: Import the regular modules

In the VBA Editor, go to **File → Import File...** and import each of these (in order):

| File | Notes |
|------|-------|
| `src/Config.bas` | Constants and runtime variables |
| `src/Utilities.bas` | Helpers, LLM routing, file I/O |
| `src/EmailFilter.bas` | Classification engine |
| `src/EmailAgent.bas` | Agent tools (addressing, auto-reply) |
| `src/BatchFilter.bas` | Bulk operations and macros |

> **Do NOT import `src/ThisOutlookSession.bas`** — it requires special handling (next step).

#### Step 2: Paste ThisOutlookSession

This module **cannot** be imported normally — it must be copy-pasted into the built-in object:

1. In the Project Explorer (left panel), expand **"Microsoft Outlook Objects"**
2. Double-click **ThisOutlookSession** to open its code window
3. Press **Ctrl+A** → **Delete** to clear any existing code
4. Open `src/ThisOutlookSession.bas` in a text editor (Notepad, VS Code, etc.)
5. Copy everything from `Option Explicit` to the end of the file
6. Paste into the VBA Editor's ThisOutlookSession window

#### Step 3: Compile and save

1. **Debug → Compile Project** — fix any errors before proceeding
2. **Ctrl+S** to save
3. Close the VBA Editor

#### Step 4: Create Inbox sub-folders

Right-click your **Inbox** → **New Folder** and create these six folders:

| Folder | Purpose |
|--------|---------|
| `Review` | Ambiguous emails awaiting manual triage |
| `Protected` | Emails from protected domains (kept untouched) |
| `LearnKeep` | Drag here → sender always KEPT in future |
| `LearnDelete` | Drag here → sender always DELETED in future |
| `LearnSubjectDelete` | Drag here → subject keyword always DELETED in future |
| `LearnReply` | Drag your sent replies here → agent learns your reply style |

> Folder names are configurable in `settings.ini` under `[Folders]`.

---

## Part 3: First Run & Verification

1. **Restart Outlook** completely (check the system tray — quit fully)
2. On startup, the agent auto-creates `settings.ini` at:
   ```
   %APPDATA%\OutlookEmailFilter\settings.ini
   ```
3. Check the VBA **Immediate Window** (Alt+F11 → Ctrl+G) for:
   ```
   [INFO] Settings loaded from: C:\Users\...\AppData\Roaming\OutlookEmailFilter\settings.ini
   [INFO] Email Agent v3.0.0 initialized
   ```
4. Run a dry-run to confirm the classification chain is working:
   ```
   FilterExistingDryRun
   ```
   Results appear in the Immediate Window — no emails are moved or deleted.

5. Run the version/status check:
   ```
   ShowVersionInfo
   ```
   This shows the LLM provider, learned rule counts, data file paths, and auto-reply status.

---

## Part 4: Configure Settings

Open `%APPDATA%\OutlookEmailFilter\settings.ini` in any text editor. Changes take effect immediately (no restart needed for pattern changes; folder name changes require a restart).

The file has five sections:

### `[General]`
```ini
EnableLogging=True
LogLevel=INFO          ; DEBUG | INFO | WARN | ERROR
EnableSelfImproving=True
DebugMode=False        ; True = MsgBox on every error (for debugging)
DryRunLimit=50
LLMBatchSize=10
```

### `[Folders]`
```ini
Protected=Protected
Review=Review
LearnKeep=LearnKeep
LearnDelete=LearnDelete
LearnSubject=LearnSubjectDelete
LearnReply=LearnReply
```

### `[Patterns]`
Comma-separated strings — case-insensitive substring match:
```ini
ProtectedDomains=substack.com,reddit.com
NamePatterns=Xu Xin,XuXin,Prof Xu
GreetingPatterns=Dear Professor Xu,Dear Prof. Xu
PolyUTags=[MM],[HRO],[CUS],ToXX
VIPSubjectKeywords=thesis,dissertation
DeleteSenderPatterns=notice,noreply,no-reply
DeleteSubjectPatterns=优惠,offer,digest,unsubscribe
```

### `[LLM]`
```ini
UseLLMAPI=False
Provider=azure         ; local | azure | claude | openai
AzureEndpoint=https://YOUR-RESOURCE.openai.azure.com/openai/deployments/YOUR-DEPLOYMENT/chat/completions?api-version=2024-02-01
LocalEndpoint=http://localhost:11434/v1/chat/completions
LocalModel=qwen3:8b
ClaudeEndpoint=https://api.anthropic.com/v1/messages
ClaudeModel=claude-opus-4-20250115
OpenAIEndpoint=https://openrouter.ai/api/v1/chat/completions
OpenAIModel=qwen/qwen3-8b
APIKeyMethod=ENV       ; ENV | HARDCODED
APIKeyEnvVar=LLM_API_KEY
APIKeyHardcoded=
ClassifyMaxTokens=100
SummarizeMaxTokens=300
ReplyMaxTokens=800
Temperature=0.3
ReplyTemperature=0.7
SystemPrompt=You are filtering emails for...
```

**To enable LLM classification**: set `UseLLMAPI=True` and configure your provider.
- `local`: Ollama, LM Studio, or Inferencer must be running on `LocalEndpoint`; no API key needed
- `azure`: set `AzureEndpoint` + API key via `APIKeyEnvVar` or `APIKeyHardcoded`
- `claude`: set `APIKeyEnvVar=ANTHROPIC_API_KEY` or hardcode your Anthropic key
- `openai`: set `OpenAIEndpoint` + `OpenAIModel` + API key via `APIKeyEnvVar` or `APIKeyHardcoded` (works with OpenRouter, Groq, Together AI, OpenAI, etc.)

### `[Agent]`
```ini
EnableAutoReply=False
AutoReplyOnArrival=False   ; draft reply automatically for every new KEEP email
LearnReplyFolder=LearnReply
MaxReplyExamples=5
ReplyPersona=              ; blank = auto-generated from your name patterns
ScanSentItems=False
ScanSentDays=30
AutoReplyForSenders=       ; blank = all KEEP emails; or comma-separated sender list
```

---

## Part 5: Quick Access Toolbar (QAT)

For one-click access to the most useful macros:

1. In Outlook: **File → Options → Quick Access Toolbar**
2. In **"Choose commands from"**, select **Macros**
3. Add these recommended macros:
   - `Project1.FilterSelectedEmails` — filter selected email(s)
   - `Project1.FilterCurrentFolder` — filter current folder
   - `Project1.FilterExistingDryRun` — dry-run preview
4. Click **Add >>** → optionally **Modify...** to pick an icon → **OK**

---

## Part 6: Enable Real-Time Filtering (Optional)

By default, the agent only filters when you run a macro manually. To filter new emails automatically as they arrive:

1. Open **ThisOutlookSession** in VBA Editor
2. Find the commented block: `'Private Sub inboxItems_ItemAdd`
3. Uncomment the entire Sub (remove the `'` from every line)
4. Save and restart Outlook

> **Auto-reply on arrival**: If you also want a draft reply created for every new KEEP email,
> set `AutoReplyOnArrival=True` and `EnableAutoReply=True` in `settings.ini`
> (requires `UseLLMAPI=True` and a configured provider).

---

## Part 7: Agent Tools Setup (Optional)

### Generate Addressing Patterns (LLM-powered)

If you're setting up the filter for a new name, let the LLM generate all variants:
```
GenerateAddressingPatterns
```
Prompts for your name and title → generates `NamePatterns` and `GreetingPatterns` → saves to `settings.ini`.
Requires `UseLLMAPI=True`.

### Teach the Agent Your Reply Style

**Option A — Drag sent replies into LearnReply folder:**
Drag any of your past sent emails (replies) into the `LearnReply` Inbox sub-folder.
The agent extracts the reply pair and appends it to `learned_replies.txt`.

**Option B — Scan Sent Items automatically:**
```
ScanSentForReplyPatterns
```
Scans your Sent Items for the last N days (set `ScanSentDays` in settings.ini) and extracts reply pairs.

Once reply pairs are learned, `DraftReplyToSelected` uses them as few-shot examples for the LLM.

---

## Part 8: Web UI Setup (Optional)

The Web UI is a Python Flask server that provides a browser-based interface for settings, macros, learned rules, email browsing, and a chat interface. It is completely optional — all functionality is also available via VBA macros and `settings.ini`.

### Requirements

- Python 3.10 or later
- Windows recommended (Outlook COM features require Windows; settings/rules/logs work on any OS)
- Outlook must be running for macro execution via the command bridge

### Install and run

```bash
cd webui
pip install -r requirements.txt
python server.py
```

Open `http://localhost:5000` in your browser.

### How the macro bridge works

The Web UI cannot call VBA directly (Outlook has no external `Application.Run()` API). Instead it uses a file-based bridge:

1. Clicking a macro button writes a JSON command file to `%APPDATA%\OutlookEmailFilter\commands\`
2. The VBA command poller (started automatically from `Application_Startup`) checks that directory every 2 seconds
3. VBA executes the macro and writes a result JSON file
4. The browser polls for the result and displays the output

The command poller starts automatically when Outlook starts and stops when you run `DisableRealTimeFilter`. If Outlook is not running, macro commands will time out after 30 seconds.

### What works without Outlook running

- Settings tab — read and edit `settings.ini` directly
- Learned Rules tab — view `learned_senders.txt`, `learned_subjects.txt`, `learned_replies.txt`
- Logs tab — view `error.log`
- Chat tab — setting-change commands work without Outlook; macro commands time out

---

## Upgrading from v2.0

1. **Export your existing modules** (optional backup): In the VBA Editor, right-click each module in the Project Explorer → **Export File...** → save to a folder of your choice. Repeat for `Config`, `Utilities`, `EmailFilter`, `EmailAgent`, `BatchFilter`, and `ThisOutlookSession`.

2. **Re-import the modules** (Part 2 above) — right-click each old module in the Project Explorer → **Remove** → **No**, then re-import the updated `.bas` files.

3. **Run the migration macro** if you have old v1.x folder names (I, II, III, IIII, V):
   ```
   DetectAndMigrateOldFolders
   ```

4. Your existing `settings.ini`, `learned_senders.txt`, and `learned_subjects.txt` are **preserved** — re-importing modules does not affect data files.

5. After restart, check `ShowVersionInfo` shows `v3.0.0`.

---

## Migrating to a New PC

### Step 1: Copy data files from the old PC

From `%APPDATA%\OutlookEmailFilter\` on the old PC, copy:

| File | Contains |
|------|----------|
| `settings.ini` | All configured settings and patterns |
| `learned_senders.txt` | Learned sender rules (KEEP/DELETE) |
| `learned_subjects.txt` | Learned subject rules (DELETE) |
| `learned_replies.txt` | Learned reply style examples |

### Step 2: Install on the new PC

Follow Parts 1–3 above. On first restart, Outlook auto-creates default data files.

### Step 3: Replace defaults with your data

1. Close Outlook on the new PC
2. Navigate to `%APPDATA%\OutlookEmailFilter\`
3. Replace all four files with your copies from the old PC
4. Restart Outlook

### Step 4: Verify

```
ShowVersionInfo
ShowLearnedSenders
FilterExistingDryRun
```

---

## Troubleshooting

### "Macros are disabled"
Enable macros in Trust Center (Part 1) and restart Outlook.

### "Compile error: Can't find project or library"
In VBA Editor → **Tools → References** → uncheck anything marked **MISSING** → ensure **"Microsoft Outlook X.X Object Library"** is checked.

### Emails not filtered automatically
- Confirm `inboxItems_ItemAdd` is uncommented in ThisOutlookSession
- Run `ReinitializeFilter` in the Immediate Window
- Restart Outlook

### LLM calls failing
- Check `error.log` at `%APPDATA%\OutlookEmailFilter\error.log` for the error message and call stack
- For `local`: confirm Ollama/LM Studio/Inferencer is running on the configured endpoint
- For `claude`: confirm `ANTHROPIC_API_KEY` environment variable is set (or use `APIKeyMethod=HARDCODED`)
- For `azure`: confirm the full deployment URL including `?api-version=...` is in `AzureEndpoint`
- For `openai`: confirm `OpenAIEndpoint` and `OpenAIModel` are set, and your API key is valid for that service
- Set `DebugMode=True` in `settings.ini` to get a MsgBox on every error

### Undo accidental deletions
- `RestoreDeletedKeepEmails` — rescues emails from Deleted Items
- `RestoreFromReview` — moves Review folder emails back to Inbox

### Reset everything
Delete `settings.ini` and restart Outlook — a fresh default file is auto-created.
Learned rules files (`learned_senders.txt`, etc.) are **not** affected.

---

## Using a Local LLM from Parallels

If you run Outlook in a Windows VM on **Parallels Desktop** (macOS) and serve a local LLM from the Mac host, `localhost` inside the VM refers to Windows — not macOS. You need to use the host's IP instead.

### Setup

1. **Find the macOS host IP.** In Parallels Shared Networking mode (the default), the host is typically `10.211.55.2`. Verify from Windows CMD inside the VM:
   ```
   ping 10.211.55.2
   ```

2. **Configure the local endpoint in `settings.ini`:**
   ```ini
   [LLM]
   Provider=local
   LocalEndpoint=http://10.211.55.2:11434/v1/chat/completions
   LocalModel=qwen3:8b
   ```
   > **Port varies by server**: Ollama uses `11434`, LM Studio uses `1234`, Inferencer uses `54321`. Adjust the port in the URL accordingly.

3. **Make your LLM server listen on all interfaces** (on macOS):
   - **Ollama**: `OLLAMA_HOST=0.0.0.0 ollama serve`
   - **LM Studio**: open Server settings → enable **"Serve on Local Network"**
   - **Inferencer**: enable **"OpenAI Compatible API"** in settings (serves on port 54321 by default)

4. **Allow incoming connections.** macOS may show a firewall prompt when Ollama/LM Studio/Inferencer starts listening — click **Allow**.

5. **Verify connectivity** from Windows CMD inside the VM:
   ```
   curl http://10.211.55.2:11434/v1/models
   ```
   You should see a JSON response listing available models.

### Notes

- If you use a different Parallels networking mode (e.g., Bridged), the host IP will differ — check macOS **System Settings → Network** for the active IP.
- The same approach works for the `openai` provider if you run a local OpenAI-compatible server on macOS — just set `OpenAIEndpoint` to `http://10.211.55.2:<port>/v1/chat/completions`.
