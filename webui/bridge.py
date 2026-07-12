r"""bridge.py — File-based command bridge between the Web UI and Outlook VBA.

VBA cannot be invoked from outside Outlook, so we use a file-based IPC:
  1. Server writes a JSON command file to %APPDATA%\OutlookEmailFilter\commands\
  2. VBA polls that directory every 2 seconds, picks up the file
  3. VBA executes the macro, writes a result JSON file alongside
  4. The browser polls the server for the result

Command file format:  commands/<id>.json   {"id":"...", "macro":"...", "args":{}}
Result file format:   commands/<id>.result {"id":"...", "status":"ok|error", "output":"..."}
"""

import json
import os
import time
import uuid
from typing import Optional

import config

STALE_COMMAND_SECONDS = 10   # unconsumed .json older than this ⇒ poller inactive
RESULT_SETTLE_SECONDS = 10   # unparseable .result younger than this ⇒ still being written
RESULT_REREAD_DELAY = 0.2    # re-read delay for truncated result files
STARTUP_CLEANUP_SECONDS = 3600
_ID_ALLOC_ATTEMPTS = 5


def _ensure_dir() -> None:
    os.makedirs(config.commands_dir(), exist_ok=True)


def _path(cmd_id: str, ext: str) -> str:
    return os.path.join(config.commands_dir(), f"{cmd_id}.{ext}")


# ---------------------------------------------------------------------------
# Send
# ---------------------------------------------------------------------------

def send_command(macro: str, args: Optional[dict] = None) -> str:
    """Write a command file and return the 8-hex-char command ID.

    Uses open(..., "x") so an ID collision with an existing file raises and a
    fresh ID is generated (max 5 attempts).
    """
    _ensure_dir()
    payload_args = dict(args or {})
    for _ in range(_ID_ALLOC_ATTEMPTS):
        cmd_id = uuid.uuid4().hex[:8]
        if os.path.exists(_path(cmd_id, "result")):
            continue
        try:
            with open(_path(cmd_id, "json"), "x", encoding="utf-8") as f:
                json.dump({"id": cmd_id, "macro": macro, "args": payload_args}, f)
            return cmd_id
        except FileExistsError:
            continue
    raise RuntimeError("Could not allocate a unique command id after 5 attempts")


# ---------------------------------------------------------------------------
# Receive
# ---------------------------------------------------------------------------

def get_result(cmd_id: str) -> Optional[dict]:
    """Check for a result file.

    Returns None while pending (no file, file locked, or file present but
    still being written — i.e. unparseable and younger than 10 s). Once a
    result is successfully returned, both <id>.result and any leftover
    <id>.json are deleted (best-effort).
    """
    result_path = _path(cmd_id, "result")
    if not os.path.exists(result_path):
        return None

    raw = _read_raw(result_path)
    if raw is None:
        return None  # locked by the writer — treat as pending

    data = _parse_result(raw)
    if data is None:
        # Possibly truncated mid-write: re-read once after a short delay.
        time.sleep(RESULT_REREAD_DELAY)
        raw2 = _read_raw(result_path)
        if raw2 is not None:
            raw = raw2
            data = _parse_result(raw)

    if data is None:
        if _file_age(result_path) <= RESULT_SETTLE_SECONDS:
            return None  # give the writer time to finish — still pending
        data = {
            "id": cmd_id,
            "status": "error",
            "output": raw.strip() or "Result file is empty or unparseable",
        }

    _cleanup_pair(cmd_id)
    return data


def _read_raw(path: str) -> Optional[str]:
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            return f.read()
    except (OSError, PermissionError):
        return None


def _parse_result(raw: str) -> Optional[dict]:
    stripped = raw.strip()
    if not stripped:
        return None
    try:
        data = json.loads(stripped)
    except (json.JSONDecodeError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def _file_age(path: str) -> float:
    try:
        return time.time() - os.path.getmtime(path)
    except OSError:
        return 0.0


def _cleanup_pair(cmd_id: str) -> None:
    for ext in ("result", "json"):
        try:
            os.remove(_path(cmd_id, ext))
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Health & housekeeping
# ---------------------------------------------------------------------------

def poller_health() -> dict:
    """Detect a dead Outlook poller: unconsumed .json files older than 10 s."""
    _ensure_dir()
    now = time.time()
    stale = 0
    for fname in os.listdir(config.commands_dir()):
        if not fname.endswith(".json"):
            continue
        try:
            if now - os.path.getmtime(os.path.join(config.commands_dir(), fname)) \
                    > STALE_COMMAND_SECONDS:
                stale += 1
        except OSError:
            pass
    return {"stale_commands": stale, "poller_responsive": stale == 0}


def cleanup_old_files(max_age_seconds: int = STARTUP_CLEANUP_SECONDS) -> int:
    """Remove .json and .result files older than max_age_seconds."""
    _ensure_dir()
    now = time.time()
    removed = 0
    for fname in os.listdir(config.commands_dir()):
        if not (fname.endswith(".json") or fname.endswith(".result")):
            continue
        path = os.path.join(config.commands_dir(), fname)
        try:
            if now - os.path.getmtime(path) > max_age_seconds:
                os.remove(path)
                removed += 1
        except OSError:
            pass
    return removed
