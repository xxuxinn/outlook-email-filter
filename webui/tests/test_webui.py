"""Pytest suite for the Outlook Email Agent Web UI (Flask).

All tests run against a temporary data directory injected through the
OUTLOOK_FILTER_DATA_DIR environment variable — config.py resolves every path
at call time, so no module reloading is required.

Run:  pytest webui/tests/test_webui.py -v
"""

import json
import os
import sys
import time

import pytest

WEBUI_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if WEBUI_DIR not in sys.path:
    sys.path.insert(0, WEBUI_DIR)

import auth            # noqa: E402
import bridge          # noqa: E402
import chat            # noqa: E402
import config          # noqa: E402
import datafiles       # noqa: E402
import macros          # noqa: E402
import server          # noqa: E402
import settings_manager  # noqa: E402


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def data_dir(tmp_path, monkeypatch):
    """Temp data dir wired in through the env var config.py reads."""
    monkeypatch.setenv(config.ENV_DATA_DIR, str(tmp_path))
    (tmp_path / "commands").mkdir()
    return tmp_path


@pytest.fixture()
def client(data_dir):
    server.app.config["TESTING"] = True
    return server.app.test_client()


@pytest.fixture()
def headers(data_dir):
    return {"X-Auth-Token": auth.get_token()}


def _write_settings(data_dir, text, encoding="utf-8-sig"):
    (data_dir / "settings.ini").write_text(text, encoding=encoding)


def _command_macros(data_dir):
    """Parse all pending command .json files; return list of macro names."""
    found = []
    for fname in os.listdir(data_dir / "commands"):
        if fname.endswith(".json"):
            with open(data_dir / "commands" / fname, encoding="utf-8") as f:
                found.append(json.load(f))
    return found


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class TestAuth:
    def test_api_requires_token(self, client):
        r = client.get("/api/status")
        assert r.status_code == 401
        assert r.get_json()["error"] == "unauthorized"

    def test_api_rejects_wrong_token(self, client):
        r = client.get("/api/status", headers={"X-Auth-Token": "0" * 32})
        assert r.status_code == 401

    def test_api_accepts_valid_token(self, client, headers):
        r = client.get("/api/status", headers=headers)
        assert r.status_code == 200

    def test_index_open_and_token_substituted(self, client, headers):
        r = client.get("/")
        assert r.status_code == 200
        html = r.get_data(as_text=True)
        assert "__AUTH_TOKEN__" not in html
        assert headers["X-Auth-Token"] in html

    def test_token_file_created_32_hex(self, data_dir):
        token = auth.get_token()
        assert len(token) == 32
        assert all(c in "0123456789abcdef" for c in token)
        on_disk = (data_dir / "webui_token.txt").read_text().strip()
        assert on_disk == token


# ---------------------------------------------------------------------------
# Macro allowlist / arg validation
# ---------------------------------------------------------------------------

class TestMacroAllowlist:
    def test_unknown_macro_rejected(self, client, headers):
        r = client.post("/api/command", json={"macro": "DeleteEverything"},
                        headers=headers)
        assert r.status_code == 400

    def test_local_only_macros_not_in_manifest(self, client, headers):
        for name in ("EnableRealTimeFilter", "DisableRealTimeFilter"):
            r = client.post("/api/command", json={"macro": name}, headers=headers)
            assert r.status_code == 400

    def test_days_out_of_range_rejected(self, client, headers):
        for bad in ("0", "366", "-3", "abc"):
            r = client.post("/api/command",
                            json={"macro": "FilterLastNDays", "args": {"days": bad}},
                            headers=headers)
            assert r.status_code == 400, f"days={bad} should be rejected"

    def test_pattern_length_rejected(self, client, headers):
        for bad in ("ab", "x" * 101):
            r = client.post("/api/command",
                            json={"macro": "BulkDeleteBySender",
                                  "args": {"pattern": bad}},
                            headers=headers)
            assert r.status_code == 400

    def test_undeclared_arg_rejected(self, client, headers):
        r = client.post("/api/command",
                        json={"macro": "ShowVersionInfo", "args": {"evil": "1"}},
                        headers=headers)
        assert r.status_code == 400

    def test_missing_required_arg_rejected(self, client, headers):
        r = client.post("/api/command", json={"macro": "FilterLastNDays"},
                        headers=headers)
        assert r.status_code == 400

    def test_valid_command_writes_file(self, client, headers, data_dir):
        r = client.post("/api/command",
                        json={"macro": "FilterLastNDays", "args": {"days": "7"}},
                        headers=headers)
        assert r.status_code == 200
        cmd_id = r.get_json()["command_id"]
        assert len(cmd_id) == 8
        payloads = _command_macros(data_dir)
        assert len(payloads) == 1
        assert payloads[0]["macro"] == "FilterLastNDays"
        assert payloads[0]["args"] == {"days": "7"}
        assert payloads[0]["id"] == cmd_id

    def test_new_bridge_macros_in_manifest(self, client, headers):
        r = client.get("/api/macros", headers=headers)
        names = [m["name"] for m in r.get_json()["macros"]]
        assert "GenerateDailyDigest" in names
        assert "ProposeRules" in names
        assert "EnableRealTimeFilter" not in names
        assert "DisableRealTimeFilter" not in names


# ---------------------------------------------------------------------------
# Settings: masking + validation + encoding
# ---------------------------------------------------------------------------

class TestSettings:
    def test_get_masks_hardcoded_key(self, client, headers, data_dir):
        _write_settings(data_dir, "[LLM]\nAPIKeyHardcoded=sk-verysecret\nProvider=azure\n")
        r = client.get("/api/settings", headers=headers)
        llm = r.get_json()["LLM"]
        assert llm["APIKeyHardcoded"] == "__MASKED__"
        assert llm["Provider"] == "azure"

    def test_empty_secret_not_masked(self, client, headers, data_dir):
        _write_settings(data_dir, "[LLM]\nAPIKeyHardcoded=\n")
        r = client.get("/api/settings", headers=headers)
        assert r.get_json()["LLM"]["APIKeyHardcoded"] == ""

    def test_masked_write_is_noop(self, client, headers, data_dir):
        _write_settings(data_dir, "[LLM]\nAPIKeyHardcoded=sk-verysecret\n")
        r = client.post("/api/settings",
                        json={"section": "LLM", "key": "APIKeyHardcoded",
                              "value": "__MASKED__"},
                        headers=headers)
        assert r.status_code == 200
        assert r.get_json()["ok"] is True
        assert settings_manager.read_setting("LLM", "APIKeyHardcoded") == "sk-verysecret"

    def test_masked_write_in_section_is_noop(self, client, headers, data_dir):
        _write_settings(data_dir, "[LLM]\nAPIKeyHardcoded=sk-verysecret\nProvider=azure\n")
        r = client.post("/api/settings/section",
                        json={"section": "LLM",
                              "values": {"APIKeyHardcoded": "__MASKED__",
                                         "Provider": "claude"}},
                        headers=headers)
        assert r.status_code == 200
        assert settings_manager.read_setting("LLM", "APIKeyHardcoded") == "sk-verysecret"
        assert settings_manager.read_setting("LLM", "Provider") == "claude"

    def test_real_secret_write_goes_through(self, client, headers, data_dir):
        _write_settings(data_dir, "[LLM]\nAPIKeyHardcoded=old\n")
        client.post("/api/settings",
                    json={"section": "LLM", "key": "APIKeyHardcoded", "value": "new"},
                    headers=headers)
        assert settings_manager.read_setting("LLM", "APIKeyHardcoded") == "new"

    def test_unknown_section_rejected(self, client, headers, data_dir):
        r = client.post("/api/settings",
                        json={"section": "Hax", "key": "K", "value": "v"},
                        headers=headers)
        assert r.status_code == 400

    def test_bad_key_rejected(self, client, headers, data_dir):
        r = client.post("/api/settings",
                        json={"section": "General", "key": "bad key!", "value": "v"},
                        headers=headers)
        assert r.status_code == 400

    def test_newline_in_value_rejected(self, client, headers, data_dir):
        r = client.post("/api/settings",
                        json={"section": "General", "key": "LogLevel",
                              "value": "INFO\n[LLM]\nAPIKeyHardcoded=pwned"},
                        headers=headers)
        assert r.status_code == 400

    def test_write_is_utf8_sig(self, client, headers, data_dir):
        client.post("/api/settings",
                    json={"section": "General", "key": "LogLevel", "value": "INFO"},
                    headers=headers)
        raw = (data_dir / "settings.ini").read_bytes()
        assert raw.startswith(b"\xef\xbb\xbf")

    def test_encoding_fallback_cp950(self, data_dir):
        text = "[Patterns]\nDeleteSubjectPatterns=優惠,offer,digest\n"
        raw = text.encode("cp950")
        with pytest.raises(UnicodeDecodeError):
            raw.decode("utf-8")  # prove the fixture actually exercises fallback
        (data_dir / "settings.ini").write_bytes(raw)
        cfg = settings_manager.read_all()
        assert "優惠" in cfg["Patterns"]["DeleteSubjectPatterns"]

    def test_utf8_sig_read(self, data_dir):
        _write_settings(data_dir, "[Patterns]\nDeleteSubjectPatterns=優惠\n",
                        encoding="utf-8-sig")
        cfg = settings_manager.read_all()
        assert cfg["Patterns"]["DeleteSubjectPatterns"] == "優惠"


# ---------------------------------------------------------------------------
# Command result / cmd_id validation / bridge lifecycle
# ---------------------------------------------------------------------------

class TestCommandResult:
    def test_invalid_cmd_id_rejected(self, client, headers):
        for bad in ("ABCDEF12", "zzzzzzzz", "12345", "a" * 9):
            r = client.get(f"/api/command/{bad}/result", headers=headers)
            assert r.status_code == 400, f"cmd_id={bad!r} should be rejected"

    def test_traversal_cmd_id_never_reaches_handler(self, client, headers):
        # Flask's router rejects slash-containing ids before the handler runs,
        # so path traversal gets a 404 (route mismatch) rather than a 400.
        r = client.get("/api/command/../../etc/result", headers=headers)
        assert r.status_code in (400, 404)

    def test_missing_result_is_pending(self, client, headers):
        r = client.get("/api/command/deadbeef/result", headers=headers)
        assert r.status_code == 200
        assert r.get_json()["status"] == "pending"

    def test_result_read_deletes_files(self, client, headers, data_dir):
        cmd_dir = data_dir / "commands"
        (cmd_dir / "deadbeef.json").write_text('{"id":"deadbeef","macro":"X","args":{}}')
        (cmd_dir / "deadbeef.result").write_text(
            '{"id":"deadbeef","status":"ok","output":"Moved 3 emails"}')
        r = client.get("/api/command/deadbeef/result", headers=headers)
        body = r.get_json()
        assert body["status"] == "ok"
        assert body["output"] == "Moved 3 emails"
        assert not (cmd_dir / "deadbeef.result").exists()
        assert not (cmd_dir / "deadbeef.json").exists()

    def test_fresh_truncated_result_is_pending(self, client, headers, data_dir):
        (data_dir / "commands" / "deadbeef.result").write_text(
            '{"id":"deadbeef","status":"ok","outp')  # truncated mid-write
        r = client.get("/api/command/deadbeef/result", headers=headers)
        assert r.get_json()["status"] == "pending"
        # File must NOT have been deleted while pending
        assert (data_dir / "commands" / "deadbeef.result").exists()

    def test_old_truncated_result_is_error_with_raw(self, client, headers, data_dir):
        path = data_dir / "commands" / "deadbeef.result"
        path.write_text('{"broken json')
        old = time.time() - 60
        os.utime(path, (old, old))
        r = client.get("/api/command/deadbeef/result", headers=headers)
        body = r.get_json()
        assert body["status"] == "error"
        assert '{"broken json' in body["output"]

    def test_send_command_id_format(self, data_dir):
        cmd_id = bridge.send_command("ShowVersionInfo")
        assert len(cmd_id) == 8
        assert all(c in "0123456789abcdef" for c in cmd_id)

    def test_cleanup_old_files(self, data_dir):
        cmd_dir = data_dir / "commands"
        old_ts = time.time() - 7200
        for name in ("old1.json", "old2.result"):
            (cmd_dir / name).write_text("{}")
            os.utime(cmd_dir / name, (old_ts, old_ts))
        (cmd_dir / "new1.json").write_text("{}")
        removed = bridge.cleanup_old_files(max_age_seconds=3600)
        assert removed == 2
        assert (cmd_dir / "new1.json").exists()


class TestBridgeHealth:
    def test_healthy_when_no_stale_commands(self, client, headers):
        r = client.get("/api/bridge/health", headers=headers)
        body = r.get_json()
        assert body["poller_responsive"] is True
        assert body["stale_commands"] == 0

    def test_unhealthy_with_stale_command(self, client, headers, data_dir):
        path = data_dir / "commands" / "cafe0123.json"
        path.write_text("{}")
        old = time.time() - 60
        os.utime(path, (old, old))
        r = client.get("/api/bridge/health", headers=headers)
        body = r.get_json()
        assert body["poller_responsive"] is False
        assert body["stale_commands"] == 1

    def test_status_includes_bridge_and_digest(self, client, headers, data_dir):
        digests = data_dir / "digests"
        digests.mkdir()
        (digests / "digest_2026-07-12.md").write_text("# Digest", encoding="utf-8")
        r = client.get("/api/status", headers=headers)
        body = r.get_json()
        assert body["bridge_ok"] is True
        assert body["latest_digest"] == "2026-07-12"


# ---------------------------------------------------------------------------
# Digest
# ---------------------------------------------------------------------------

class TestDigest:
    def test_no_digest_404(self, client, headers):
        r = client.get("/api/digest", headers=headers)
        assert r.status_code == 404
        assert r.get_json()["ok"] is False

    def test_latest_digest_returned(self, client, headers, data_dir):
        digests = data_dir / "digests"
        digests.mkdir()
        (digests / "digest_2026-07-10.md").write_text("old", encoding="utf-8")
        (digests / "digest_2026-07-12.md").write_text(
            "# Daily Digest\n- 5 kept\n- 12 deleted\n", encoding="utf-8")
        r = client.get("/api/digest", headers=headers)
        body = r.get_json()
        assert r.status_code == 200
        assert body["date"] == "2026-07-12"
        assert "12 deleted" in body["content"]

    def test_generate_sends_bridge_command(self, client, headers, data_dir):
        r = client.post("/api/digest/generate", headers=headers)
        assert r.status_code == 200
        cmd_id = r.get_json()["command_id"]
        payloads = _command_macros(data_dir)
        assert [p["macro"] for p in payloads] == ["GenerateDailyDigest"]
        assert payloads[0]["id"] == cmd_id


# ---------------------------------------------------------------------------
# Rule proposals
# ---------------------------------------------------------------------------

PROPOSALS = (
    "a1b2c3d4|SENDER|spam@example.com|DELETE|Deleted 12 times|PENDING|2026-07-10 10:00:00\n"
    "b2c3d4e5|SUBJECT|Weekly Digest|DELETE|Repeated deletes|PENDING|2026-07-10 10:05:00\n"
    "c3d4e5f6|SENDER|boss@polyu.edu.hk|KEEP|Always kept|PENDING|2026-07-10 10:06:00\n"
)


class TestProposals:
    @pytest.fixture(autouse=True)
    def _proposals_file(self, data_dir):
        (data_dir / "rule_proposals.txt").write_text(PROPOSALS, encoding="utf-8")

    def test_list_proposals(self, client, headers):
        r = client.get("/api/proposals", headers=headers)
        rows = r.get_json()
        assert len(rows) == 3
        assert rows[0] == {
            "id": "a1b2c3d4", "type": "SENDER", "value": "spam@example.com",
            "action": "DELETE", "reason": "Deleted 12 times",
            "status": "PENDING", "timestamp": "2026-07-10 10:00:00",
        }

    def test_approve_sender_appends_rule_and_reloads(self, client, headers, data_dir):
        r = client.post("/api/proposals/a1b2c3d4/approve", headers=headers)
        assert r.status_code == 200
        body = r.get_json()
        assert body["ok"] is True
        assert body["reload_macro"] == "ReloadLearnedSenders"

        learned = (data_dir / "learned_senders.txt").read_text(encoding="utf-8")
        line = learned.strip().splitlines()[-1]
        parts = line.split("|")
        assert parts[0] == "spam@example.com"
        assert parts[1] == "DELETE"
        time.strptime(parts[2], "%Y-%m-%d %H:%M:%S")  # timestamp format

        proposals = (data_dir / "rule_proposals.txt").read_text(encoding="utf-8")
        assert "a1b2c3d4|SENDER|spam@example.com|DELETE|Deleted 12 times|APPROVED|" in proposals
        assert "b2c3d4e5|SUBJECT|Weekly Digest|DELETE|Repeated deletes|PENDING|" in proposals

        assert [p["macro"] for p in _command_macros(data_dir)] == ["ReloadLearnedSenders"]

    def test_approve_subject_appends_delete_and_reinitializes(self, client, headers, data_dir):
        r = client.post("/api/proposals/b2c3d4e5/approve", headers=headers)
        assert r.status_code == 200
        assert r.get_json()["reload_macro"] == "ReinitializeFilter"
        learned = (data_dir / "learned_subjects.txt").read_text(encoding="utf-8")
        assert learned.startswith("Weekly Digest|DELETE|")
        assert [p["macro"] for p in _command_macros(data_dir)] == ["ReinitializeFilter"]

    def test_approve_keep_sender(self, client, headers, data_dir):
        client.post("/api/proposals/c3d4e5f6/approve", headers=headers)
        learned = (data_dir / "learned_senders.txt").read_text(encoding="utf-8")
        assert learned.startswith("boss@polyu.edu.hk|KEEP|")

    def test_reject_updates_status_only(self, client, headers, data_dir):
        r = client.post("/api/proposals/a1b2c3d4/reject", headers=headers)
        assert r.status_code == 200
        proposals = (data_dir / "rule_proposals.txt").read_text(encoding="utf-8")
        assert "|REJECTED|" in proposals
        assert not (data_dir / "learned_senders.txt").exists()
        assert _command_macros(data_dir) == []

    def test_double_approve_rejected(self, client, headers):
        assert client.post("/api/proposals/a1b2c3d4/approve",
                           headers=headers).status_code == 200
        r = client.post("/api/proposals/a1b2c3d4/approve", headers=headers)
        assert r.status_code == 400

    def test_unknown_id_404(self, client, headers):
        r = client.post("/api/proposals/deadbeef/approve", headers=headers)
        assert r.status_code == 404

    def test_bad_id_format_400(self, client, headers):
        for bad in ("XYZ", "A1B2C3D4", "a1b2c3d"):
            r = client.post(f"/api/proposals/{bad}/approve", headers=headers)
            assert r.status_code == 400
        # Slash-containing ids never match the route (404 from the router).
        r = client.post("/api/proposals/../../x/approve", headers=headers)
        assert r.status_code in (400, 404)

    def test_value_sanitized_on_approve(self, client, headers, data_dir):
        (data_dir / "rule_proposals.txt").write_text(
            "d4e5f6a7|SUBJECT|Bad\tvalue with spaces|DELETE|r|PENDING|2026-07-10 10:00:00\n",
            encoding="utf-8")
        client.post("/api/proposals/d4e5f6a7/approve", headers=headers)
        learned = (data_dir / "learned_subjects.txt").read_text(encoding="utf-8")
        assert "|" not in learned.strip().splitlines()[-1].split("|")[0]

    def test_generate_sends_propose_rules(self, client, headers, data_dir):
        r = client.post("/api/proposals/generate", headers=headers)
        assert r.status_code == 200
        assert [p["macro"] for p in _command_macros(data_dir)] == ["ProposeRules"]


# ---------------------------------------------------------------------------
# Decisions + n-param clamping
# ---------------------------------------------------------------------------

class TestDecisionsAndClamping:
    def test_decisions_parsed(self, client, headers, data_dir):
        (data_dir / "decision_log.txt").write_text(
            "2026-07-12 09:00:00|spam@x.com|Buy now|RULE|DELETE|0.95\n"
            "2026-07-12 09:01:00|boss@polyu.edu.hk|Thesis|LLM|KEEP|0.80\n",
            encoding="utf-8")
        r = client.get("/api/decisions?n=100", headers=headers)
        rows = r.get_json()
        assert len(rows) == 2
        assert rows[1] == {
            "timestamp": "2026-07-12 09:01:00", "sender": "boss@polyu.edu.hk",
            "subject": "Thesis", "source": "LLM", "action": "KEEP",
            "confidence": "0.80",
        }

    def test_n_param_clamped_high(self, client, headers, data_dir):
        lines = "".join(f"2026-07-12 09:00:00|s{i}@x.com|S|RULE|DELETE|1.0\n"
                        for i in range(1500))
        (data_dir / "decision_log.txt").write_text(lines, encoding="utf-8")
        r = client.get("/api/decisions?n=999999", headers=headers)
        assert len(r.get_json()) == 1000  # clamped to max 1000

    def test_n_param_clamped_low(self, client, headers, data_dir):
        (data_dir / "error.log").write_text("a\nb\nc\n", encoding="utf-8")
        r = client.get("/api/errors?n=0", headers=headers)
        assert len(r.get_json()) == 1  # clamped to min 1

    def test_n_param_non_numeric_uses_default(self, client, headers, data_dir):
        (data_dir / "error.log").write_text("line\n" * 5, encoding="utf-8")
        r = client.get("/api/errors?n=abc", headers=headers)
        assert r.status_code == 200
        assert len(r.get_json()) == 5


# ---------------------------------------------------------------------------
# Log clear endpoints (honest failure)
# ---------------------------------------------------------------------------

class TestLogClear:
    def test_clear_truncates(self, client, headers, data_dir):
        (data_dir / "error.log").write_text("boom\n", encoding="utf-8")
        r = client.post("/api/errors/clear", headers=headers)
        assert r.get_json()["ok"] is True
        assert (data_dir / "error.log").read_text() == ""

    def test_clear_reports_oserror(self, client, headers, data_dir, monkeypatch):
        def boom(*args, **kwargs):
            raise OSError("disk on fire")
        monkeypatch.setattr("builtins.open", boom)
        r = client.post("/api/errors/clear", headers=headers)
        assert r.status_code == 500
        assert r.get_json()["ok"] is False
        assert "disk on fire" in r.get_json()["error"]


# ---------------------------------------------------------------------------
# Chat parser
# ---------------------------------------------------------------------------

class TestChat:
    @pytest.mark.parametrize("message,macro", [
        ("generate digest", "GenerateDailyDigest"),
        ("daily digest please", "GenerateDailyDigest"),
        ("digest", "GenerateDailyDigest"),
        ("propose rules", "ProposeRules"),
        ("suggest rules", "ProposeRules"),
        ("mine rules", "ProposeRules"),
        ("show version", "ShowVersionInfo"),
        ("status", "ShowVersionInfo"),
        ("draft replies", "DraftReplyForSelected"),
        ("sync rules", "SyncLearnedRules"),
    ])
    def test_macro_mappings(self, message, macro):
        action = chat.parse(message)
        assert action is not None
        assert action["type"] == "macro"
        assert action["macro"] == macro

    def test_info_alone_is_unknown(self):
        assert chat.parse("info") is None

    def test_all_chat_macros_in_manifest(self):
        for _, action in chat.COMMAND_MAP:
            if action["type"] == "macro":
                assert macros.get_macro(action["macro"]) is not None, action["macro"]

    def test_chat_endpoint_macro(self, client, headers, data_dir):
        r = client.post("/api/chat", json={"message": "generate digest"},
                        headers=headers)
        body = r.get_json()
        assert body["type"] == "macro"
        assert [p["macro"] for p in _command_macros(data_dir)] == ["GenerateDailyDigest"]

    def test_chat_endpoint_unknown(self, client, headers):
        r = client.post("/api/chat", json={"message": "frobnicate the mail"},
                        headers=headers)
        assert r.get_json()["type"] == "unknown"


# ---------------------------------------------------------------------------
# Macro manifest self-consistency
# ---------------------------------------------------------------------------

class TestManifest:
    def test_all_entries_have_required_fields(self):
        for m in macros.MACROS:
            for field in ("name", "label", "description", "category",
                          "args", "destructive"):
                assert field in m, f"{m.get('name')} missing {field}"

    def test_no_duplicate_names(self):
        names = [m["name"] for m in macros.MACROS]
        assert len(names) == len(set(names))
