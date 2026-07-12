# Outlook Email Agent — MCP Server

**Technical:** A self-contained MCP (Model Context Protocol) stdio server (`outlook_agent_mcp.py`, Python 3.10+, dependency: `mcp>=1.2`) that exposes the VBA Outlook Email Agent to Claude Desktop / Claude Code. Bridge-backed tools write `%APPDATA%\OutlookEmailFilter\commands\<id>.json` (`{"id","macro","args"}`); the VBA poller picks commands up every 2 s and writes `<id>.result` (`{"id","status":"ok"|"error","output"}`), which the server polls every 0.5 s (default timeout 120 s). File-backed tools read the agent's data files directly. It deliberately does not import anything from `webui/` — the ~40 lines of bridge logic are copied in.

**In plain terms:** this small program lets Claude talk to your Outlook email filter. Claude drops a note ("run this macro") into a folder that Outlook watches; Outlook does the work and drops an answer back. Claude can also read the agent's notebooks (rules, logs, digests) straight from disk — that part works even when Outlook is closed.

## Tools

**Technical:** three tiers — bridge-backed (need Outlook running), file-backed (read-only, no Outlook needed), and one careful write tool.
**In plain terms:** some tools ask Outlook to do something live; others just read files; one appends a single rule line and never touches existing mail.

### Bridge-backed (Outlook must be running; all take optional `timeout`, default 120 s)

| Tool | VBA macro | What it does |
|------|-----------|--------------|
| `run_dry_run()` | `FilterExistingDryRun` | Preview classifications — no changes |
| `filter_last_n_days(days)` | `FilterLastNDays` | Classify **and act** on last N days (1–365) — modifies mailbox |
| `generate_classification_report()` | `GenerateClassificationReport` | Count classifications without acting |
| `generate_daily_digest()` | `GenerateDailyDigest` | Write `digests/digest_YYYY-MM-DD.md` |
| `propose_rules()` | `ProposeRules` | Suggest new rules — nothing applied automatically |
| `summarize_selected_email()` | `SummarizeSelectedEmail` | LLM summary of the email selected in Outlook |
| `draft_reply_for_selected()` | `DraftReplyForSelected` | Draft reply → Drafts folder, **never sent** |
| `sync_learned_rules()` | `SyncLearnedRules` | Bidirectional cloud sync of learned rules |
| `reload_learned_rules()` | `ReloadLearnedSenders` | Reload rule files from disk into the VBA cache |

### File-backed (read-only; work with Outlook closed)

| Tool | Reads | What it does |
|------|-------|--------------|
| `get_status()` | `settings.ini`, rule files, `digests/` | Version, LLM provider/enabled, deduped rule counts, latest digest date |
| `get_learned_rules(rule_type)` | `learned_senders.txt` / `learned_subjects.txt` | `'senders'` or `'subjects'`; deduped last-entry-wins listing |
| `get_latest_digest()` | `digests/digest_*.md` | Full content of the newest digest |
| `get_recent_decisions(n=50)` | `decision_log.txt` | Last N decisions (`timestamp\|sender\|subject\|source\|action\|confidence`) as a table |
| `get_error_log(n=50)` | `error.log` | Last N error-log lines |

### Write-backed

| Tool | Writes | What it does |
|------|--------|--------------|
| `add_learned_sender_rule(email, action)` | `learned_senders.txt` | Validates (`KEEP`/`DELETE`, `@`, no pipe/CR/LF), appends `email\|ACTION\|timestamp`, then best-effort live reload via the bridge — if Outlook is closed the rule still loads on next start |

## Install & run

```bash
cd mcp
pip install -r requirements.txt   # just: mcp>=1.2.0
python outlook_agent_mcp.py       # or: python -m outlook_agent_mcp  (run from mcp/)
```

**Technical:** the server speaks MCP over stdio; you don't normally run it by hand — the Claude client launches it from the config below.
**In plain terms:** you register it once in Claude's settings; Claude starts and stops it automatically in the background.

## Claude Desktop registration

Add to `claude_desktop_config.json` (Windows: `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook-agent": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook-email-filter\\mcp\\outlook_agent_mcp.py"]
    }
  }
}
```

## Claude Code registration

```bash
claude mcp add outlook-agent -- python C:\\path\\to\\mcp\\outlook_agent_mcp.py
```

## Data directory

**Technical:** resolution order — `OUTLOOK_FILTER_DATA_DIR` env var → `%APPDATA%\OutlookEmailFilter` (Windows) → `~/.outlook-email-filter` (fallback). Set the env var to point the server at a different data dir (used by the tests):

```json
"outlook-agent": {
  "command": "python",
  "args": ["C:\\path\\to\\mcp\\outlook_agent_mcp.py"],
  "env": { "OUTLOOK_FILTER_DATA_DIR": "D:\\some\\other\\dir" }
}
```

**In plain terms:** the server looks for the agent's files in the standard Windows location automatically; the env var is an override knob for testing or unusual setups.

## Security notes

- **No destructive bulk tools.** `FilterExistingEmails`, `FilterAllFolders`, `BulkDeleteBySender`, and similar mailbox-wide destructive macros are deliberately **not** exposed. The only acting filter tool is `filter_last_n_days`, which is scope-limited, and anything it deletes stays recoverable in Deleted Items.
- **Drafts are never sent.** `draft_reply_for_selected` only places drafts in the Outlook Drafts folder for manual review.
- **Writes are minimal and append-only.** The single write tool appends one validated line to `learned_senders.txt`; it never edits or deletes existing rules or mail.
- Bridge commands are plain JSON files under `%APPDATA%` — same trust boundary as the existing Web UI bridge; no network ports are opened.

## Tests

```bash
pip install pytest
python -m pytest tests/ -v
```

Tests run entirely against a temp data dir (`OUTLOOK_FILTER_DATA_DIR`) with fixture files and a simulated VBA responder thread — no Outlook, no Windows required.
