"""server.py — Flask Web UI for the Outlook Email Agent.

Run with: python server.py
Access at: http://localhost:5000

Routes:
  GET  /                           → serve SPA (index.html)
  GET  /api/settings               → read settings.ini as JSON
  POST /api/settings               → write {section, key, value}
  POST /api/settings/reload        → send ReinitializeFilter command via bridge
  GET  /api/learned/senders        → read learned_senders.txt
  GET  /api/learned/subjects       → read learned_subjects.txt
  GET  /api/learned/replies        → read learned_replies.txt
  GET  /api/errors                 → read error.log (last N lines)
  POST /api/errors/clear            → truncate error.log
  GET  /api/llm-debug-log           → read llm_debug.log (last N lines)
  POST /api/llm-debug-log/clear     → truncate llm_debug.log
  POST /api/command                → send macro command via bridge
  GET  /api/command/<id>/result    → poll for command result
  POST /api/chat                   → parse chat message → action
  GET  /api/status                 → version, provider, rule counts
"""

import os
import sys
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

import bridge
import chat
import settings_manager

app = Flask(__name__, static_folder="static")

DATA_DIR = settings_manager.SETTINGS_DIR


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _read_pipe_file(filename: str, max_lines: int = 500) -> list[dict]:
    """Read a pipe-delimited data file and return list of row dicts."""
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        return []
    rows = []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("|")
            rows.append(parts)
    return rows[-max_lines:]


def _read_log(filename: str, max_lines: int = 200) -> list[str]:
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()
    return [l.rstrip() for l in lines[-max_lines:]]


def _count_rules(filename: str) -> int:
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        return 0
    seen = set()
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                key = line.split("|")[0].lower()
                seen.add(key)
    return len(seen)


# ---------------------------------------------------------------------------
# SPA
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


# ---------------------------------------------------------------------------
# Settings API
# ---------------------------------------------------------------------------

@app.route("/api/settings", methods=["GET"])
def get_settings():
    return jsonify(settings_manager.read_all())


@app.route("/api/settings", methods=["POST"])
def post_setting():
    data = request.get_json()
    section = data.get("section", "")
    key = data.get("key", "")
    value = data.get("value", "")
    if not section or not key:
        return jsonify({"error": "section and key required"}), 400
    settings_manager.write_setting(section, key, value)
    return jsonify({"ok": True})


@app.route("/api/settings/section", methods=["POST"])
def post_section():
    data = request.get_json()
    section = data.get("section", "")
    values = data.get("values", {})
    if not section or not values:
        return jsonify({"error": "section and values required"}), 400
    settings_manager.write_section(section, values)
    return jsonify({"ok": True})


@app.route("/api/settings/reload", methods=["POST"])
def reload_settings():
    cmd_id = bridge.send_command("ReinitializeFilter")
    return jsonify({"ok": True, "command_id": cmd_id})


# ---------------------------------------------------------------------------
# Learned Rules API
# ---------------------------------------------------------------------------

@app.route("/api/learned/senders")
def learned_senders():
    rows = _read_pipe_file("learned_senders.txt")
    result = []
    seen = {}
    for parts in rows:
        if len(parts) >= 2:
            email = parts[0].strip()
            action = parts[1].strip().upper()
            ts = parts[2].strip() if len(parts) >= 3 else ""
            seen[email.lower()] = {"email": email, "action": action, "timestamp": ts}
    return jsonify(list(seen.values()))


@app.route("/api/learned/subjects")
def learned_subjects():
    rows = _read_pipe_file("learned_subjects.txt")
    result = []
    seen = {}
    for parts in rows:
        if len(parts) >= 2:
            subj = parts[0].strip()
            action = parts[1].strip().upper()
            ts = parts[2].strip() if len(parts) >= 3 else ""
            seen[subj.lower()] = {"subject": subj, "action": action, "timestamp": ts}
    return jsonify(list(seen.values()))


@app.route("/api/learned/replies")
def learned_replies():
    rows = _read_pipe_file("learned_replies.txt", max_lines=50)
    result = []
    for parts in rows:
        if len(parts) >= 4:
            result.append({
                "subject": parts[0].strip(),
                "from": parts[1].strip(),
                "original_snippet": parts[2].strip()[:200],
                "reply_snippet": parts[3].strip()[:300],
                "timestamp": parts[4].strip() if len(parts) >= 5 else "",
            })
    return jsonify(result)


# ---------------------------------------------------------------------------
# Error Log API
# ---------------------------------------------------------------------------

@app.route("/api/errors")
def get_errors():
    max_lines = int(request.args.get("n", 100))
    lines = _read_log("error.log", max_lines)
    return jsonify(lines)


@app.route("/api/llm-debug-log")
def get_llm_debug_log():
    max_lines = int(request.args.get("n", 200))
    lines = _read_log("llm_debug.log", max_lines)
    return jsonify(lines)


@app.route("/api/errors/clear", methods=["POST"])
def clear_errors():
    """Clear error.log contents (truncate file, don't delete)."""
    path = os.path.join(DATA_DIR, "error.log")
    try:
        open(path, "w").close()  # truncate to zero bytes
    except OSError:
        pass
    return jsonify({"ok": True})


@app.route("/api/llm-debug-log/clear", methods=["POST"])
def clear_llm_debug_log():
    """Clear llm_debug.log contents (truncate file, don't delete)."""
    path = os.path.join(DATA_DIR, "llm_debug.log")
    try:
        open(path, "w").close()
    except OSError:
        pass
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Command Bridge API
# ---------------------------------------------------------------------------

@app.route("/api/command", methods=["POST"])
def send_command():
    data = request.get_json()
    macro = data.get("macro", "")
    args = data.get("args", {})
    if not macro:
        return jsonify({"error": "macro name required"}), 400
    cmd_id = bridge.send_command(macro, args)
    return jsonify({"command_id": cmd_id, "status": "pending"})


@app.route("/api/command/<cmd_id>/result")
def get_command_result(cmd_id):
    result = bridge.get_result(cmd_id)
    if result is None:
        return jsonify({"status": "pending"})
    return jsonify(result)


@app.route("/api/command/debug")
def debug_commands():
    """Debug endpoint: list all files in the commands directory."""
    files = []
    if os.path.isdir(bridge.COMMANDS_DIR):
        for f in sorted(os.listdir(bridge.COMMANDS_DIR)):
            path = os.path.join(bridge.COMMANDS_DIR, f)
            size = os.path.getsize(path) if os.path.isfile(path) else 0
            preview = ""
            if os.path.isfile(path) and size < 5000:
                try:
                    with open(path, "r", encoding="utf-8", errors="replace") as fh:
                        preview = fh.read()[:500]
                except Exception:
                    preview = "(unreadable)"
            files.append({"name": f, "size": size, "preview": preview})
    return jsonify({"commands_dir": bridge.COMMANDS_DIR, "files": files})


# ---------------------------------------------------------------------------
# Chat API
# ---------------------------------------------------------------------------

@app.route("/api/chat", methods=["POST"])
def handle_chat():
    data = request.get_json()
    message = data.get("message", "").strip()
    if not message:
        return jsonify({"error": "message required"}), 400

    action = chat.parse(message)

    if action is None:
        return jsonify({
            "type": "unknown",
            "label": "I didn't understand that.",
            "output": f'Unknown command: "{message}"\n\n' + chat.help_text(),
        })

    if action["type"] == "help":
        return jsonify({
            "type": "help",
            "label": "Help",
            "output": chat.help_text(),
        })

    if action["type"] == "api":
        # The client will make the API call itself
        return jsonify(action)

    if action["type"] == "setting":
        settings_manager.write_setting(action["section"], action["key"], action["value"])
        return jsonify({
            "type": "setting",
            "label": action["label"],
            "output": f"Set [{action['section']}] {action['key']} = {action['value']}",
        })

    if action["type"] == "macro":
        cmd_id = bridge.send_command(action["macro"])
        return jsonify({
            "type": "macro",
            "label": action["label"],
            "command_id": cmd_id,
            "output": f"Command sent to Outlook (ID: {cmd_id}). Waiting for result...",
        })

    return jsonify({"type": "error", "output": "Unknown action type"})


# ---------------------------------------------------------------------------
# Status API
# ---------------------------------------------------------------------------

@app.route("/api/status")
def get_status():
    settings = settings_manager.read_all()
    general = settings.get("General", {})
    llm = settings.get("LLM", {})
    return jsonify({
        "version": general.get("Version", "unknown"),
        "settings_path": settings_manager.settings_path(),
        "llm_provider": llm.get("Provider", "azure"),
        "llm_enabled": llm.get("UseLLMAPI", "False"),
        "learned_senders": _count_rules("learned_senders.txt"),
        "learned_subjects": _count_rules("learned_subjects.txt"),
        "commands_dir": bridge.COMMANDS_DIR,
    })


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Clean up stale result files from previous sessions
    cleaned = bridge.cleanup_old_results(max_age_seconds=60)
    if cleaned:
        print(f"Cleaned up {cleaned} stale result file(s)")
    print(f"Outlook Email Agent Web UI")
    print(f"Open: http://localhost:{port}")
    print(f"Settings: {settings_manager.settings_path()}")
    app.run(debug=False, port=port, host="127.0.0.1")
