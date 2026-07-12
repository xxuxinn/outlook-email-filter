"""server.py — Flask Web UI for the Outlook Email Agent.

Run with: python server.py
Access at: http://localhost:5000

All /api/* routes require the X-Auth-Token header. The token is generated at
<DATA_DIR>/webui_token.txt and injected into index.html when it is served, so
opening the page in a browser is all that's needed. See README.md for the
full route table.
"""

import os
import re

from flask import Flask, Response, jsonify, request

import auth
import bridge
import chat
import config
import datafiles
import macros
import settings_manager

app = Flask(__name__, static_folder="static")

_CMD_ID_RE = re.compile(r"^[0-9a-f]{8}$")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _int_param(name: str, default: int, lo: int = 1, hi: int = 1000) -> int:
    """Parse an int query param defensively; clamp to [lo, hi]."""
    raw = request.args.get(name)
    if raw is None:
        return default
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, value))


def _json_body() -> dict:
    data = request.get_json(silent=True)
    return data if isinstance(data, dict) else {}


def _valid_cmd_id(cmd_id: str) -> bool:
    return bool(_CMD_ID_RE.match(cmd_id or ""))


def _send_bridge_command(macro_name: str, args: dict | None = None):
    """Send a bridge command; return (cmd_id, error_message)."""
    try:
        return bridge.send_command(macro_name, args), None
    except (OSError, RuntimeError) as e:
        return None, f"Could not write command file: {e}"


# ---------------------------------------------------------------------------
# Auth — every /api/* request must carry X-Auth-Token
# ---------------------------------------------------------------------------

@app.before_request
def _require_auth():
    if request.path.startswith("/api/"):
        if not auth.check_token(request.headers.get("X-Auth-Token")):
            return jsonify({"error": "unauthorized"}), 401
    return None


# ---------------------------------------------------------------------------
# SPA — index.html served with the auth token substituted in
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    path = os.path.join(app.static_folder, "index.html")
    try:
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
    except OSError:
        return Response("index.html not found", status=500, mimetype="text/plain")
    html = html.replace("__AUTH_TOKEN__", auth.get_token())
    return Response(html, mimetype="text/html")


# ---------------------------------------------------------------------------
# Settings API
# ---------------------------------------------------------------------------

@app.route("/api/settings", methods=["GET"])
def get_settings():
    return jsonify(settings_manager.read_all(mask_secrets=True))


@app.route("/api/settings", methods=["POST"])
def post_setting():
    data = _json_body()
    section = data.get("section", "")
    key = data.get("key", "")
    value = data.get("value", "")
    if not section or not key:
        return jsonify({"ok": False, "error": "section and key required"}), 400
    try:
        settings_manager.write_setting(section, key, value)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except OSError as e:
        return jsonify({"ok": False, "error": f"Could not write settings.ini: {e}"}), 500
    return jsonify({"ok": True})


@app.route("/api/settings/section", methods=["POST"])
def post_section():
    data = _json_body()
    section = data.get("section", "")
    values = data.get("values", {})
    if not section or not isinstance(values, dict) or not values:
        return jsonify({"ok": False, "error": "section and values required"}), 400
    try:
        settings_manager.write_section(section, values)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except OSError as e:
        return jsonify({"ok": False, "error": f"Could not write settings.ini: {e}"}), 500
    return jsonify({"ok": True})


@app.route("/api/settings/reload", methods=["POST"])
def reload_settings():
    cmd_id, err = _send_bridge_command("ReinitializeFilter")
    if err:
        return jsonify({"ok": False, "error": err}), 500
    return jsonify({"ok": True, "command_id": cmd_id})


# ---------------------------------------------------------------------------
# Macro manifest
# ---------------------------------------------------------------------------

@app.route("/api/macros")
def get_macros():
    return jsonify({"macros": macros.MACROS})


# ---------------------------------------------------------------------------
# Learned Rules API
# ---------------------------------------------------------------------------

@app.route("/api/learned/senders")
def learned_senders():
    rows = datafiles.read_pipe_file("learned_senders.txt")
    seen = {}
    for parts in rows:
        if len(parts) >= 2:
            email = parts[0].strip()
            seen[email.lower()] = {
                "email": email,
                "action": parts[1].strip().upper(),
                "timestamp": parts[2].strip() if len(parts) >= 3 else "",
            }
    return jsonify(list(seen.values()))


@app.route("/api/learned/subjects")
def learned_subjects():
    rows = datafiles.read_pipe_file("learned_subjects.txt")
    seen = {}
    for parts in rows:
        if len(parts) >= 2:
            subj = parts[0].strip()
            seen[subj.lower()] = {
                "subject": subj,
                "action": parts[1].strip().upper(),
                "timestamp": parts[2].strip() if len(parts) >= 3 else "",
            }
    return jsonify(list(seen.values()))


@app.route("/api/learned/replies")
def learned_replies():
    rows = datafiles.read_pipe_file("learned_replies.txt", max_lines=50)
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
# Logs API
# ---------------------------------------------------------------------------

@app.route("/api/errors")
def get_errors():
    return jsonify(datafiles.read_log("error.log", _int_param("n", 100)))


@app.route("/api/llm-debug-log")
def get_llm_debug_log():
    return jsonify(datafiles.read_log("llm_debug.log", _int_param("n", 200)))


def _truncate_log(filename: str):
    path = config.data_file(filename)
    try:
        open(path, "w", encoding="utf-8").close()
    except OSError as e:
        return jsonify({"ok": False, "error": f"Could not clear {filename}: {e}"}), 500
    return jsonify({"ok": True})


@app.route("/api/errors/clear", methods=["POST"])
def clear_errors():
    return _truncate_log("error.log")


@app.route("/api/llm-debug-log/clear", methods=["POST"])
def clear_llm_debug_log():
    return _truncate_log("llm_debug.log")


# ---------------------------------------------------------------------------
# Command Bridge API
# ---------------------------------------------------------------------------

@app.route("/api/command", methods=["POST"])
def send_command():
    data = _json_body()
    macro_name = data.get("macro", "")
    if not macro_name:
        return jsonify({"error": "macro name required"}), 400

    clean_args, err = macros.validate_command(macro_name, data.get("args", {}))
    if err:
        return jsonify({"error": err}), 400

    cmd_id, send_err = _send_bridge_command(macro_name, clean_args)
    if send_err:
        return jsonify({"error": send_err}), 500
    return jsonify({"command_id": cmd_id, "status": "pending"})


@app.route("/api/command/<cmd_id>/result")
def get_command_result(cmd_id):
    if not _valid_cmd_id(cmd_id):
        return jsonify({"error": "invalid command id"}), 400
    result = bridge.get_result(cmd_id)
    if result is None:
        return jsonify({"status": "pending"})
    return jsonify(result)


@app.route("/api/command/debug")
def debug_commands():
    """Debug endpoint: list all files in the commands directory."""
    files = []
    commands_dir = config.commands_dir()
    if os.path.isdir(commands_dir):
        for fname in sorted(os.listdir(commands_dir)):
            path = os.path.join(commands_dir, fname)
            size = os.path.getsize(path) if os.path.isfile(path) else 0
            preview = ""
            if os.path.isfile(path) and size < 5000:
                try:
                    with open(path, "r", encoding="utf-8", errors="replace") as fh:
                        preview = fh.read()[:500]
                except OSError:
                    preview = "(unreadable)"
            files.append({"name": fname, "size": size, "preview": preview})
    return jsonify({"commands_dir": commands_dir, "files": files})


@app.route("/api/bridge/health")
def bridge_health():
    """Report whether the Outlook command poller appears to be consuming files."""
    health = bridge.poller_health()
    return jsonify({"ok": True, **health})


# ---------------------------------------------------------------------------
# Digest API
# ---------------------------------------------------------------------------

@app.route("/api/digest")
def get_digest():
    digest = datafiles.latest_digest()
    if digest is None:
        return jsonify({"ok": False, "error": "No digest generated yet"}), 404
    return jsonify({"ok": True, "date": digest["date"], "content": digest["content"]})


@app.route("/api/digest/generate", methods=["POST"])
def generate_digest():
    cmd_id, err = _send_bridge_command("GenerateDailyDigest")
    if err:
        return jsonify({"ok": False, "error": err}), 500
    return jsonify({"ok": True, "command_id": cmd_id, "status": "pending"})


# ---------------------------------------------------------------------------
# Rule Proposals API
# ---------------------------------------------------------------------------

@app.route("/api/proposals")
def get_proposals():
    return jsonify(datafiles.read_proposals())


@app.route("/api/proposals/generate", methods=["POST"])
def generate_proposals():
    cmd_id, err = _send_bridge_command("ProposeRules")
    if err:
        return jsonify({"ok": False, "error": err}), 500
    return jsonify({"ok": True, "command_id": cmd_id, "status": "pending"})


@app.route("/api/proposals/<proposal_id>/approve", methods=["POST"])
def approve_proposal(proposal_id):
    if not _valid_cmd_id(proposal_id):
        return jsonify({"ok": False, "error": "invalid proposal id"}), 400
    proposal, err, status = datafiles.approve_proposal(proposal_id)
    if err:
        return jsonify({"ok": False, "error": err}), status

    # Tell Outlook to pick up the new rule: sender rules reload directly;
    # subject rules are only re-read on ReinitializeFilter.
    reload_macro = ("ReloadLearnedSenders" if proposal["type"] == "SENDER"
                    else "ReinitializeFilter")
    cmd_id, send_err = _send_bridge_command(reload_macro)
    response = {"ok": True, "proposal": proposal, "reload_macro": reload_macro}
    if send_err:
        return jsonify({**response, "reload_error": send_err})
    return jsonify({**response, "command_id": cmd_id})


@app.route("/api/proposals/<proposal_id>/reject", methods=["POST"])
def reject_proposal(proposal_id):
    if not _valid_cmd_id(proposal_id):
        return jsonify({"ok": False, "error": "invalid proposal id"}), 400
    proposal, err, status = datafiles.reject_proposal(proposal_id)
    if err:
        return jsonify({"ok": False, "error": err}), status
    return jsonify({"ok": True, "proposal": proposal})


# ---------------------------------------------------------------------------
# Decision Log API
# ---------------------------------------------------------------------------

@app.route("/api/decisions")
def get_decisions():
    return jsonify(datafiles.read_decisions(_int_param("n", 100)))


# ---------------------------------------------------------------------------
# Chat API
# ---------------------------------------------------------------------------

@app.route("/api/chat", methods=["POST"])
def handle_chat():
    data = _json_body()
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
        return jsonify({"type": "help", "label": "Help", "output": chat.help_text()})

    if action["type"] == "api":
        # The client will make the API call itself
        return jsonify(action)

    if action["type"] == "setting":
        try:
            settings_manager.write_setting(
                action["section"], action["key"], action["value"])
        except (ValueError, OSError) as e:
            return jsonify({"type": "error", "output": f"Could not save setting: {e}"})
        return jsonify({
            "type": "setting",
            "label": action["label"],
            "output": f"Set [{action['section']}] {action['key']} = {action['value']}",
        })

    if action["type"] == "macro":
        if macros.get_macro(action["macro"]) is None:
            return jsonify({"type": "error",
                            "output": f"Macro not allowed: {action['macro']}"})
        cmd_id, err = _send_bridge_command(action["macro"])
        if err:
            return jsonify({"type": "error", "output": err})
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
    health = bridge.poller_health()
    return jsonify({
        "version": general.get("Version", "unknown"),
        "settings_path": config.settings_path(),
        "llm_provider": llm.get("Provider", "azure"),
        "llm_enabled": llm.get("UseLLMAPI", "False"),
        "learned_senders": datafiles.count_rules("learned_senders.txt"),
        "learned_subjects": datafiles.count_rules("learned_subjects.txt"),
        "commands_dir": config.commands_dir(),
        "bridge_ok": health["poller_responsive"],
        "stale_commands": health["stale_commands"],
        "latest_digest": datafiles.latest_digest_date(),
    })


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    cleaned = bridge.cleanup_old_files()  # .json + .result older than 1 hour
    if cleaned:
        print(f"Cleaned up {cleaned} stale command/result file(s)")
    auth.get_token()  # ensure the token file exists before first request
    print("Outlook Email Agent Web UI")
    print(f"Open: http://localhost:{port}")
    print(f"Settings: {config.settings_path()}")
    print(f"Auth token file: {config.token_path()}")
    app.run(debug=False, port=port, host="127.0.0.1")
