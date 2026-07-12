r"""outlook_agent_mcp.py — MCP stdio server for the Outlook Email Agent (VBA).

Lets Claude Desktop / Claude Code drive the local VBA email agent through the
existing file-based command bridge, and read the agent's data files.

Command bridge protocol (VBA side polls every 2 seconds):
  1. This server writes  <data_dir>\commands\<id>.json
       {"id": "...", "macro": "...", "args": {...}}
  2. Outlook VBA picks the file up and runs the macro.
  3. VBA writes  <data_dir>\commands\<id>.result
       {"id": "...", "status": "ok" | "error", "output": "<text>"}
  4. This server polls for the result file (default timeout 120 s).

Data dir resolution (first match wins):
  1. OUTLOOK_FILTER_DATA_DIR environment variable
  2. %APPDATA%\OutlookEmailFilter          (Windows — the real deployment)
  3. ~/.outlook-email-filter               (non-Windows fallback for testing)

Run:  python outlook_agent_mcp.py     or:  python -m outlook_agent_mcp

Deliberately NOT exposed: destructive bulk macros (FilterExistingEmails,
BulkDeleteBySender, FilterAllFolders, ...). Reply drafting only creates
drafts in the Drafts folder — nothing is ever auto-sent.
"""

from __future__ import annotations

import configparser
import json
import os
import time
import uuid
from pathlib import Path

from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_TIMEOUT = 120.0     # seconds to wait for a VBA macro result
POLL_INTERVAL = 0.5         # seconds between result-file polls
RELOAD_TIMEOUT = 6.0        # best-effort reload window (3 VBA poll cycles)
MAX_ID_ATTEMPTS = 16        # collision-check attempts for command ids
MAX_DAYS = 365
MAX_LOG_LINES = 1000
ENCODINGS = ("utf-8-sig", "cp950", "latin-1")

TIMEOUT_HINT = (
    "Make sure Outlook is running on the Windows machine with the Email Agent "
    "VBA project loaded (the command poller runs every 2 seconds). If Outlook "
    "was just started, wait a few seconds and try again."
)

mcp = FastMCP(
    "outlook-email-agent",
    instructions=(
        "Tools for Professor Xu Xin's Outlook Email Agent (VBA). Bridge-backed "
        "tools require Outlook to be running on the local Windows machine; "
        "file-backed tools (get_status, get_learned_rules, get_latest_digest, "
        "get_recent_decisions, get_error_log) work anytime. No tool sends "
        "email or bulk-deletes anything."
    ),
)

# ---------------------------------------------------------------------------
# Data-dir resolution and file helpers
# ---------------------------------------------------------------------------


def get_data_dir() -> Path:
    """Resolve the agent data directory (env override > %APPDATA% > home)."""
    override = os.environ.get("OUTLOOK_FILTER_DATA_DIR")
    if override:
        return Path(override)
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "OutlookEmailFilter"
    return Path.home() / ".outlook-email-filter"


def get_commands_dir() -> Path:
    """Directory watched by the VBA command poller."""
    return get_data_dir() / "commands"


def read_text_fallback(path: Path) -> str:
    """Read a text file trying utf-8-sig, then cp950, then latin-1."""
    data = path.read_bytes()
    for encoding in ENCODINGS[:-1]:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode(ENCODINGS[-1], errors="replace")


def read_data_lines(path: Path) -> list[str]:
    """Return non-blank, non-comment (#) lines of a pipe-delimited data file."""
    if not path.is_file():
        return []
    lines = read_text_fallback(path).splitlines()
    return [ln.strip() for ln in lines if ln.strip() and not ln.strip().startswith("#")]


def parse_pipe_rules(path: Path) -> dict[str, tuple[str, str, str]]:
    """Parse ``key|ACTION|timestamp`` lines with last-entry-wins dedup.

    Returns {lowercased_key: (original_key, action, timestamp)}. Keys are
    case-insensitive to mirror the VBA Dictionary (CompareMode=1). Malformed
    lines (fewer than 2 fields) are skipped defensively.
    """
    rules: dict[str, tuple[str, str, str]] = {}
    for line in read_data_lines(path):
        fields = line.split("|")
        if len(fields) < 2 or not fields[0].strip():
            continue
        key = fields[0].strip()
        action = fields[1].strip().upper()
        timestamp = fields[2].strip() if len(fields) >= 3 else ""
        rules[key.lower()] = (key, action, timestamp)
    return rules


def load_settings() -> configparser.ConfigParser | None:
    """Parse settings.ini, or None if it does not exist."""
    path = get_data_dir() / "settings.ini"
    if not path.is_file():
        return None
    parser = configparser.ConfigParser(
        interpolation=None, strict=False, inline_comment_prefixes=(";",)
    )
    parser.optionxform = str  # type: ignore[method-assign]  # keep key case
    parser.read_string(read_text_fallback(path))
    return parser


def list_digest_files() -> list[Path]:
    """digest_*.md files in <data_dir>/digests, oldest → newest by filename."""
    digests_dir = get_data_dir() / "digests"
    if not digests_dir.is_dir():
        return []
    return sorted(digests_dir.glob("digest_*.md"), key=lambda p: p.name)


# ---------------------------------------------------------------------------
# Command bridge (self-contained copy of the webui bridge logic)
# ---------------------------------------------------------------------------


def _new_command_id(commands_dir: Path) -> str:
    """Generate an 8-hex-char command id, checked against existing files."""
    for _ in range(MAX_ID_ATTEMPTS):
        cmd_id = uuid.uuid4().hex[:8]
        if not (commands_dir / f"{cmd_id}.json").exists() and not (
            commands_dir / f"{cmd_id}.result"
        ).exists():
            return cmd_id
    return uuid.uuid4().hex  # astronomically unlikely; fall back to full hex


def _read_result_file(path: Path) -> dict:
    """Read and JSON-parse a result file; on parse failure re-read once."""
    raw = ""
    for attempt in range(2):
        try:
            raw = read_text_fallback(path).strip()
        except OSError as exc:  # file may still be locked by VBA
            time.sleep(0.3)
            if attempt == 1:
                return {"status": "error", "output": f"Result file read error: {exc}"}
            continue
        if raw:
            try:
                return json.loads(raw)
            except (json.JSONDecodeError, ValueError):
                time.sleep(0.3)  # VBA may still be flushing; re-read once
                continue
        else:
            time.sleep(0.3)
    if not raw:
        return {"status": "error", "output": "Result file is empty"}
    return {
        "status": "ok",
        "output": f"(result was not valid JSON; raw content follows)\n{raw}",
    }


def send_and_wait(macro: str, args: dict | None = None, timeout: float = DEFAULT_TIMEOUT) -> dict:
    """Send one macro command through the file bridge and wait for its result.

    Returns {"status": "ok" | "error" | "timeout", "output": str}.
    Creates the commands dir if missing, uses a collision-checked 8-hex id,
    polls every 0.5 s, and cleans up command/result files afterwards.
    """
    commands_dir = get_commands_dir()
    commands_dir.mkdir(parents=True, exist_ok=True)

    cmd_id = _new_command_id(commands_dir)
    cmd_path = commands_dir / f"{cmd_id}.json"
    result_path = commands_dir / f"{cmd_id}.result"

    payload = {"id": cmd_id, "macro": macro, "args": args or {}}
    cmd_path.write_text(json.dumps(payload), encoding="utf-8")

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if result_path.exists():
            result = _read_result_file(result_path)
            for stale in (cmd_path, result_path):
                try:
                    stale.unlink()
                except OSError:
                    pass
            return result
        time.sleep(POLL_INTERVAL)

    try:  # timeout — remove the command file so it is not executed later
        cmd_path.unlink()
    except OSError:
        pass
    return {
        "status": "timeout",
        "output": (
            f"No response from Outlook within {timeout:.0f} s for macro "
            f"'{macro}'. {TIMEOUT_HINT}"
        ),
    }


def run_macro(macro: str, args: dict | None = None, timeout: float = DEFAULT_TIMEOUT) -> str:
    """Run a macro via the bridge and format the outcome as tool-result text."""
    result = send_and_wait(macro, args, timeout)
    status = str(result.get("status", "error"))
    output = str(result.get("output", "")).strip()
    if status == "ok":
        return output or f"{macro} completed (no output returned)."
    if status == "timeout":
        return f"TIMEOUT: {output}"
    return f"ERROR from Outlook while running {macro}: {output or '(no details)'}"


# ---------------------------------------------------------------------------
# Bridge-backed tools (require Outlook running)
# ---------------------------------------------------------------------------


@mcp.tool()
def run_dry_run(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Preview how every Inbox email would be classified, without changing anything.

    Runs the FilterExistingDryRun macro in Outlook and returns the decision
    preview (KEEP / DELETE / REVIEW / LLM per email). Completely safe: no
    email is moved or deleted. Requires Outlook to be running locally.
    """
    return run_macro("FilterExistingDryRun", timeout=timeout)


@mcp.tool()
def filter_last_n_days(days: int, timeout: float = DEFAULT_TIMEOUT) -> str:
    """Classify AND act on Inbox emails received in the last N days (1-365).

    Runs the FilterLastNDays macro: emails classified DELETE are moved to
    Deleted Items, protected/ambiguous emails are moved to their folders.
    This MODIFIES the mailbox — use run_dry_run first when unsure. Deleted
    emails remain recoverable from Deleted Items. Requires Outlook running.
    """
    if not isinstance(days, int) or isinstance(days, bool) or not 1 <= days <= MAX_DAYS:
        return f"ERROR: days must be an integer between 1 and {MAX_DAYS} (got {days!r})."
    return run_macro("FilterLastNDays", {"days": str(days)}, timeout=timeout)


@mcp.tool()
def generate_classification_report(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Count how Inbox emails would be classified, without acting on any of them.

    Runs the GenerateClassificationReport macro and returns aggregate counts
    per decision (KEEP / DELETE / REVIEW / LLM). Read-only on the mailbox.
    Requires Outlook to be running locally.
    """
    return run_macro("GenerateClassificationReport", timeout=timeout)


@mcp.tool()
def generate_daily_digest(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Generate today's email digest and save it under <data_dir>/digests/.

    Runs the GenerateDailyDigest macro, which writes digest_YYYY-MM-DD.md
    summarizing recent email activity. Side effect: creates/overwrites
    today's digest file. Read the result afterwards with get_latest_digest.
    Requires Outlook to be running locally.
    """
    return run_macro("GenerateDailyDigest", timeout=timeout)


@mcp.tool()
def propose_rules(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Ask the agent to analyze recent activity and propose new filtering rules.

    Runs the ProposeRules macro and returns rule suggestions as text.
    Nothing is applied automatically — apply a suggestion explicitly with
    add_learned_sender_rule if you agree with it. Requires Outlook running.
    """
    return run_macro("ProposeRules", timeout=timeout)


@mcp.tool()
def summarize_selected_email(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Summarize the email currently selected in the Outlook window via LLM.

    Runs the SummarizeSelectedEmail macro. The user must have an email
    selected in Outlook, and the agent's LLM integration must be enabled
    (UseLLMAPI=True in settings.ini). Read-only. Requires Outlook running.
    """
    return run_macro("SummarizeSelectedEmail", timeout=timeout)


@mcp.tool()
def draft_reply_for_selected(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Draft a reply to the email(s) selected in Outlook, saved to Drafts only.

    Runs the DraftReplyForSelected macro, which uses the few-shot reply
    engine (learned_replies.txt style examples) to write a draft. The draft
    is placed in the Outlook Drafts folder and is NEVER sent automatically —
    the user reviews and sends it manually. Requires Outlook running with
    LLM integration enabled.
    """
    return run_macro("DraftReplyForSelected", timeout=timeout)


@mcp.tool()
def sync_learned_rules(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Sync learned rules bidirectionally with the configured cloud folder.

    Runs the SyncLearnedRules macro: merges learned_senders.txt,
    learned_subjects.txt and learned_replies.txt with the OneDrive copy,
    resolving conflicts by timestamp (later wins). Side effect: rewrites
    both local and cloud rule files with the merged result. Requires
    Outlook running and EnableCloudSync=True in settings.ini.
    """
    return run_macro("SyncLearnedRules", timeout=timeout)


@mcp.tool()
def reload_learned_rules(timeout: float = DEFAULT_TIMEOUT) -> str:
    """Force Outlook to reload learned sender rules from disk.

    Runs the ReloadLearnedSenders macro so that rule-file changes made
    outside Outlook (for example via add_learned_sender_rule) take effect
    immediately instead of at the next Outlook restart. Requires Outlook
    to be running locally.
    """
    return run_macro("ReloadLearnedSenders", timeout=timeout)


# ---------------------------------------------------------------------------
# File-backed tools (read-only, work without Outlook)
# ---------------------------------------------------------------------------


@mcp.tool()
def get_status() -> str:
    """Report agent status from local data files — works even if Outlook is closed.

    Reads settings.ini (version, LLM provider, whether LLM classification is
    enabled), counts deduplicated learned sender/subject rules, and shows the
    date of the latest daily digest. Read-only; no bridge round-trip.
    """
    data_dir = get_data_dir()
    lines = [
        "Outlook Email Agent — status",
        f"Data dir      : {data_dir} ({'exists' if data_dir.is_dir() else 'MISSING'})",
    ]

    settings = load_settings()
    if settings is None:
        lines.append(
            "settings.ini  : not found — the agent has not been initialized on "
            "this machine (run Outlook once with the VBA project loaded)."
        )
    else:
        version = settings.get("General", "Version", fallback="unknown")
        provider = settings.get("LLM", "Provider", fallback="unknown")
        use_llm = settings.get("LLM", "UseLLMAPI", fallback="unknown")
        lines.append(f"Version       : {version}")
        lines.append(f"LLM           : UseLLMAPI={use_llm}, Provider={provider}")

    senders = parse_pipe_rules(data_dir / "learned_senders.txt")
    keep_count = sum(1 for _, action, _ in senders.values() if action == "KEEP")
    delete_count = sum(1 for _, action, _ in senders.values() if action == "DELETE")
    subjects = parse_pipe_rules(data_dir / "learned_subjects.txt")
    lines.append(
        f"Learned rules : {len(senders)} sender rules "
        f"({keep_count} KEEP / {delete_count} DELETE), "
        f"{len(subjects)} subject rules"
    )

    digests = list_digest_files()
    if digests:
        latest = digests[-1]
        date_part = latest.stem.replace("digest_", "")
        lines.append(f"Latest digest : {date_part} ({latest.name})")
    else:
        lines.append("Latest digest : none yet (run generate_daily_digest)")
    return "\n".join(lines)


@mcp.tool()
def get_learned_rules(rule_type: str) -> str:
    """List learned filtering rules from disk. rule_type: 'senders' or 'subjects'.

    'senders' returns email → KEEP/DELETE rules from learned_senders.txt;
    'subjects' returns subject-fragment → DELETE rules from
    learned_subjects.txt. Duplicates are collapsed last-entry-wins, exactly
    as the VBA cache loads them. Read-only; works without Outlook.
    """
    kind = rule_type.strip().lower()
    filenames = {"senders": "learned_senders.txt", "subjects": "learned_subjects.txt"}
    if kind not in filenames:
        return "ERROR: rule_type must be 'senders' or 'subjects' (got {!r}).".format(rule_type)

    path = get_data_dir() / filenames[kind]
    rules = parse_pipe_rules(path)
    if not rules:
        return f"No learned {kind} rules yet ({path} is missing or empty)."

    header = f"Learned {kind[:-1]} rules — {len(rules)} unique (last entry wins):"
    body = [
        f"  {action:<7} {key}  ({timestamp or 'no timestamp'})"
        for key, action, timestamp in sorted(rules.values(), key=lambda r: r[0].lower())
    ]
    return "\n".join([header, *body])


@mcp.tool()
def get_latest_digest() -> str:
    """Return the full content of the newest daily digest markdown file.

    Reads the most recent <data_dir>/digests/digest_YYYY-MM-DD.md. Read-only;
    works without Outlook. If no digest exists yet, explains how to create
    one (generate_daily_digest, which needs Outlook running).
    """
    digests = list_digest_files()
    if not digests:
        return (
            "No digests found in {} — run generate_daily_digest() first "
            "(Outlook must be running) to create today's digest.".format(
                get_data_dir() / "digests"
            )
        )
    latest = digests[-1]
    try:
        content = read_text_fallback(latest)
    except OSError as exc:
        return f"ERROR: could not read {latest}: {exc}"
    return f"[{latest.name}]\n\n{content}"


@mcp.tool()
def get_recent_decisions(n: int = 50) -> str:
    """Show the last N classification decisions the agent logged (default 50).

    Parses <data_dir>/decision_log.txt, whose pipe-delimited lines are
    ``timestamp|sender|subject|source|action|confidence``, and formats them
    as a compact table (newest last). Read-only; works without Outlook.
    """
    if not isinstance(n, int) or isinstance(n, bool) or not 1 <= n <= MAX_LOG_LINES:
        return f"ERROR: n must be an integer between 1 and {MAX_LOG_LINES} (got {n!r})."

    path = get_data_dir() / "decision_log.txt"
    entries = read_data_lines(path)
    if not entries:
        return f"No decisions logged yet ({path} is missing or empty)."

    recent = entries[-n:]
    header = (
        f"Last {len(recent)} of {len(entries)} decisions (newest last):\n"
        f"{'TIMESTAMP':<20} {'ACTION':<8} {'CONF':<5} {'SENDER':<32} "
        f"{'SOURCE':<14} SUBJECT"
    )
    rows = []
    for line in recent:
        fields = [f.strip() for f in line.split("|")]
        fields += [""] * (6 - len(fields))  # pad malformed lines defensively
        timestamp, sender, subject, source, action, confidence = fields[:6]
        rows.append(
            f"{timestamp[:19]:<20} {action[:8]:<8} {confidence[:5]:<5} "
            f"{sender[:32]:<32} {source[:14]:<14} {subject[:60]}"
        )
    return "\n".join([header, *rows])


@mcp.tool()
def get_error_log(n: int = 50) -> str:
    """Show the last N lines of the agent's error log (default 50).

    Reads <data_dir>/error.log — pipe-delimited entries of the form
    ``timestamp|module.procedure|errNum|errDesc|context|Stack: ...`` written
    by the VBA LogError handler. Read-only; works without Outlook. Useful
    for diagnosing why a macro failed.
    """
    if not isinstance(n, int) or isinstance(n, bool) or not 1 <= n <= MAX_LOG_LINES:
        return f"ERROR: n must be an integer between 1 and {MAX_LOG_LINES} (got {n!r})."

    path = get_data_dir() / "error.log"
    if not path.is_file():
        return f"No error log found ({path}) — no errors recorded yet."
    lines = [ln for ln in read_text_fallback(path).splitlines() if ln.strip()]
    if not lines:
        return f"Error log is empty ({path})."
    recent = lines[-n:]
    return f"Last {len(recent)} of {len(lines)} error log lines (newest last):\n" + "\n".join(
        recent
    )


# ---------------------------------------------------------------------------
# Write-backed tools (mutating, with care)
# ---------------------------------------------------------------------------


@mcp.tool()
def add_learned_sender_rule(email: str, action: str) -> str:
    """Add a learned sender rule: always KEEP or always DELETE mail from a sender.

    Validates the input, then appends ``email|ACTION|timestamp`` to
    learned_senders.txt (append-only, last entry wins — exactly how the VBA
    agent learns). Afterwards it best-effort asks Outlook to reload rules;
    if Outlook is closed the rule is still saved and loads on next start.
    Side effect: future emails from this sender are auto-kept or auto-moved
    to Deleted Items by the filter. Never deletes existing mail by itself.
    """
    action_norm = action.strip().upper()
    if action_norm not in ("KEEP", "DELETE"):
        return f"ERROR: action must be 'KEEP' or 'DELETE' (got {action!r})."

    email_norm = email.strip()
    if not email_norm or "@" not in email_norm:
        return f"ERROR: {email!r} does not look like an email address (no '@')."
    if any(ch in email_norm for ch in ("|", "\r", "\n", "\x00")):
        return "ERROR: email must not contain pipe, newline, or null characters."
    if len(email_norm) > 320:
        return "ERROR: email address is too long (max 320 characters)."

    data_dir = get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    path = data_dir / "learned_senders.txt"
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"{email_norm}|{action_norm}|{timestamp}\n"
    try:
        with path.open("a", encoding="utf-8", newline="") as handle:
            handle.write(line)
    except OSError as exc:
        return f"ERROR: could not write to {path}: {exc}"

    saved_msg = f"Saved rule: {email_norm} -> {action_norm} (appended to {path.name})."

    # Best-effort live reload — never fail the tool if Outlook is closed.
    reload_result = send_and_wait("ReloadLearnedSenders", timeout=RELOAD_TIMEOUT)
    if reload_result.get("status") == "ok":
        return f"{saved_msg} Outlook reloaded the rules — the rule is active now."
    return (
        f"{saved_msg} Outlook did not confirm a live reload (it may be closed); "
        "the rule will load automatically the next time Outlook starts, or run "
        "reload_learned_rules once Outlook is open."
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Run the MCP server on stdio (the transport Claude Desktop/Code use)."""
    mcp.run("stdio")


if __name__ == "__main__":
    main()
