r"""bridge.py — File-based command bridge between the Web UI and Outlook VBA.

VBA cannot be invoked from outside Outlook, so we use a file-based IPC:
  1. Server writes a JSON command file to %APPDATA%\OutlookEmailFilter\commands\
  2. VBA polls that directory every 2 seconds, picks up the file
  3. VBA executes the macro, writes a result JSON file alongside
  4. Server polls for the result file (or times out)

Command file format:  commands/<id>.json   {"id":"...", "macro":"...", "args":{}}
Result file format:   commands/<id>.result {"id":"...", "status":"ok|error", "output":"..."}
"""

import json
import os
import time
import uuid
from typing import Optional

import settings_manager

COMMANDS_DIR = os.path.join(settings_manager.SETTINGS_DIR, "commands")
RESULT_TIMEOUT = 30  # seconds to wait for VBA result
POLL_INTERVAL = 0.5  # seconds between result polls


def _ensure_dir() -> None:
    os.makedirs(COMMANDS_DIR, exist_ok=True)


def send_command(macro: str, args: Optional[dict] = None) -> str:
    """Write a command file and return the command ID."""
    _ensure_dir()
    cmd_id = str(uuid.uuid4())[:8]
    payload = {"id": cmd_id, "macro": macro, "args": args or {}}
    path = os.path.join(COMMANDS_DIR, f"{cmd_id}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f)
    return cmd_id


def get_result(cmd_id: str) -> Optional[dict]:
    """Check for a result file. Returns None if not yet available."""
    path = os.path.join(COMMANDS_DIR, f"{cmd_id}.result")
    if not os.path.exists(path):
        return None

    # File exists — try to read it (with one retry for file-locking races)
    for attempt in range(2):
        try:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                raw = f.read()
        except PermissionError:
            # File may still be held by VBA — wait briefly and retry
            print(f"[bridge] PermissionError reading {cmd_id}.result (attempt {attempt+1})")
            time.sleep(0.3)
            continue
        except Exception as e:
            print(f"[bridge] Error reading {cmd_id}.result: {e}")
            return {"id": cmd_id, "status": "error", "output": f"Result file read error: {e}"}

        raw = raw.strip()
        if not raw:
            print(f"[bridge] Empty result file for {cmd_id}")
            return {"id": cmd_id, "status": "error", "output": "Result file is empty"}

        # Try JSON parse
        try:
            data = json.loads(raw)
            print(f"[bridge] Result OK for {cmd_id}: status={data.get('status')}, len={len(data.get('output', ''))}")
            return data
        except (json.JSONDecodeError, ValueError) as e:
            print(f"[bridge] JSON parse failed for {cmd_id}: {e}")
            print(f"[bridge] Raw content: {raw[:300]}")
            # Return raw content as-is rather than losing the result
            return {"id": cmd_id, "status": "ok", "output": raw}

    # Both read attempts failed (PermissionError)
    return {"id": cmd_id, "status": "error", "output": "Result file locked by another process"}


def wait_for_result(cmd_id: str, timeout: float = RESULT_TIMEOUT) -> dict:
    """Poll for a result file until timeout. Returns result dict or timeout error."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = get_result(cmd_id)
        if result is not None:
            # Clean up files
            try:
                os.remove(os.path.join(COMMANDS_DIR, f"{cmd_id}.json"))
                os.remove(os.path.join(COMMANDS_DIR, f"{cmd_id}.result"))
            except OSError:
                pass
            return result
        time.sleep(POLL_INTERVAL)
    # Timeout — clean up command file if still there
    try:
        os.remove(os.path.join(COMMANDS_DIR, f"{cmd_id}.json"))
    except OSError:
        pass
    return {"id": cmd_id, "status": "timeout", "output": "No response from Outlook within timeout"}


def list_pending_commands() -> list[str]:
    """List IDs of commands that have been sent but not yet picked up."""
    _ensure_dir()
    return [f[:-5] for f in os.listdir(COMMANDS_DIR) if f.endswith(".json")]


def cleanup_old_results(max_age_seconds: int = 300) -> int:
    """Remove .result files older than max_age_seconds. Returns count removed."""
    _ensure_dir()
    now = time.time()
    removed = 0
    for f in os.listdir(COMMANDS_DIR):
        if f.endswith(".result"):
            path = os.path.join(COMMANDS_DIR, f)
            try:
                if now - os.path.getmtime(path) > max_age_seconds:
                    os.remove(path)
                    removed += 1
            except OSError:
                pass
    return removed
