"""Tests for the Outlook Email Agent MCP server (outlook_agent_mcp.py).

Covers: clean import, tool registration, file-backed tools against a fixture
data dir (via OUTLOOK_FILTER_DATA_DIR), bridge timeout behavior, the bridge
happy path (simulated VBA responder), and write-tool validation.
"""

from __future__ import annotations

import asyncio
import json
import sys
import threading
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import outlook_agent_mcp as srv  # noqa: E402  (path insert must come first)

EXPECTED_TOOLS = {
    "run_dry_run",
    "filter_last_n_days",
    "generate_classification_report",
    "generate_daily_digest",
    "propose_rules",
    "summarize_selected_email",
    "draft_reply_for_selected",
    "sync_learned_rules",
    "reload_learned_rules",
    "get_status",
    "get_learned_rules",
    "get_latest_digest",
    "get_recent_decisions",
    "get_error_log",
    "add_learned_sender_rule",
}


def call_tool(name: str, args: dict | None = None) -> str:
    """Invoke a registered tool through FastMCP and return its text output."""
    content, _structured = asyncio.run(srv.mcp.call_tool(name, args or {}))
    assert content, f"tool {name} returned no content"
    return content[0].text


@pytest.fixture()
def data_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Fixture data dir with sample agent files, wired via env override."""
    monkeypatch.setenv("OUTLOOK_FILTER_DATA_DIR", str(tmp_path))

    (tmp_path / "settings.ini").write_text(
        "[General]\n"
        "Version=3.0.0\n"
        "EnableLogging=True\n"
        "[LLM]\n"
        "UseLLMAPI=True\n"
        "Provider=claude          ; local | azure | claude | openai\n",
        encoding="utf-8",
    )

    (tmp_path / "learned_senders.txt").write_text(
        "# comment line to be skipped\n"
        "\n"
        "keep.me@example.com|KEEP|2026-07-01 10:00:00\n"
        "flip.me@example.com|KEEP|2026-07-01 11:00:00\n"
        "malformed-line-no-pipes\n"
        "FLIP.ME@example.com|DELETE|2026-07-02 09:00:00\n"
        "spam@junk.example|DELETE|2026-07-03 08:00:00\n",
        encoding="utf-8",
    )

    (tmp_path / "learned_subjects.txt").write_text(
        "Weekly Digest|DELETE|2026-07-01 10:00:00\n"
        "Funding Application Submitted|DELETE|2026-07-02 10:00:00\n",
        encoding="utf-8",
    )

    (tmp_path / "decision_log.txt").write_text(
        "2026-07-10 09:00:00|a@x.com|Hello there|Rule1|KEEP|1.0\n"
        "2026-07-10 09:05:00|b@y.com|Buy now!!|Rule8|DELETE|0.9\n"
        "malformed|only-two-fields\n"
        "2026-07-11 10:00:00|c@z.com|Thesis draft|LLM|KEEP|0.8\n",
        encoding="utf-8",
    )

    (tmp_path / "error.log").write_text(
        "2026-07-09 08:00:00|EmailFilter.ClassifyEmail|13|Type mismatch||Stack: A -> B\n"
        "2026-07-10 08:00:00|Utilities.CallLLM|-2147|Timeout||Stack: A -> C\n",
        encoding="utf-8",
    )

    digests = tmp_path / "digests"
    digests.mkdir()
    (digests / "digest_2026-07-10.md").write_text("# Digest OLD\n", encoding="utf-8")
    (digests / "digest_2026-07-11.md").write_text(
        "# Digest 2026-07-11\n\n- 12 kept, 30 deleted\n", encoding="utf-8"
    )
    return tmp_path


# ---------------------------------------------------------------------------
# (a) import + (b) tool registration
# ---------------------------------------------------------------------------


def test_module_imports_cleanly() -> None:
    assert hasattr(srv, "mcp")
    assert hasattr(srv, "main")
    assert srv.mcp.name == "outlook-email-agent"


def test_all_tools_registered_with_descriptions() -> None:
    tools = asyncio.run(srv.mcp.list_tools())
    names = {t.name for t in tools}
    assert names == EXPECTED_TOOLS
    for tool in tools:
        assert tool.description and tool.description.strip(), f"{tool.name} lacks docstring"


# ---------------------------------------------------------------------------
# (c) file-backed tools against the fixture data dir
# ---------------------------------------------------------------------------


def test_get_status(data_dir: Path) -> None:
    out = call_tool("get_status")
    assert "3.0.0" in out
    assert "Provider=claude" in out
    assert "UseLLMAPI=True" in out
    # 3 unique senders after last-entry-wins dedup (1 KEEP / 2 DELETE)
    assert "3 sender rules" in out
    assert "1 KEEP / 2 DELETE" in out
    assert "2 subject rules" in out
    assert "2026-07-11" in out


def test_get_learned_rules_senders_dedup(data_dir: Path) -> None:
    out = call_tool("get_learned_rules", {"rule_type": "senders"})
    assert "3 unique" in out
    # last entry wins, case-insensitively: flip.me ends up DELETE, listed once
    assert out.lower().count("flip.me@example.com") == 1
    assert "DELETE  FLIP.ME@example.com" in out
    assert "KEEP    keep.me@example.com" in out
    assert "malformed-line-no-pipes" not in out


def test_get_learned_rules_subjects(data_dir: Path) -> None:
    out = call_tool("get_learned_rules", {"rule_type": "subjects"})
    assert "2 unique" in out
    assert "Weekly Digest" in out


def test_get_learned_rules_invalid_type(data_dir: Path) -> None:
    out = call_tool("get_learned_rules", {"rule_type": "bogus"})
    assert out.startswith("ERROR")


def test_get_latest_digest(data_dir: Path) -> None:
    out = call_tool("get_latest_digest")
    assert "digest_2026-07-11.md" in out
    assert "12 kept, 30 deleted" in out
    assert "Digest OLD" not in out


def test_get_latest_digest_none(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("OUTLOOK_FILTER_DATA_DIR", str(tmp_path))
    out = call_tool("get_latest_digest")
    assert "No digests found" in out
    assert "generate_daily_digest" in out


def test_get_recent_decisions(data_dir: Path) -> None:
    out = call_tool("get_recent_decisions", {"n": 2})
    assert "Last 2 of 4 decisions" in out
    assert "c@z.com" in out and "Thesis draft" in out
    assert "a@x.com" not in out  # older than the last 2
    assert "malformed" in out  # defensive parse keeps the line visible


def test_get_recent_decisions_validation(data_dir: Path) -> None:
    assert call_tool("get_recent_decisions", {"n": 0}).startswith("ERROR")


def test_get_error_log(data_dir: Path) -> None:
    out = call_tool("get_error_log", {"n": 1})
    assert "Utilities.CallLLM" in out
    assert "ClassifyEmail" not in out


# ---------------------------------------------------------------------------
# (d) bridge behavior: fast graceful timeout + simulated happy path
# ---------------------------------------------------------------------------


def test_send_and_wait_times_out_gracefully(data_dir: Path) -> None:
    start = time.monotonic()
    result = srv.send_and_wait("FilterExistingDryRun", timeout=1.2)
    elapsed = time.monotonic() - start
    assert result["status"] == "timeout"
    assert "Outlook" in result["output"]
    assert elapsed < 5.0
    # command file cleaned up so a later Outlook start does not execute it
    assert not list((data_dir / "commands").glob("*.json"))


def test_bridge_tool_timeout_message(data_dir: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    out = call_tool("run_dry_run", {"timeout": 1.0})
    assert out.startswith("TIMEOUT")
    assert "Outlook" in out


def test_send_and_wait_happy_path(data_dir: Path) -> None:
    """Simulate the VBA side: answer the first command file that appears."""
    commands_dir = data_dir / "commands"

    def fake_vba() -> None:
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            pending = list(commands_dir.glob("*.json"))
            if pending:
                payload = json.loads(pending[0].read_text(encoding="utf-8"))
                assert payload["macro"] == "GenerateClassificationReport"
                assert payload["args"] == {}
                result = {"id": payload["id"], "status": "ok", "output": "KEEP: 5, DELETE: 3"}
                (commands_dir / f"{payload['id']}.result").write_text(
                    json.dumps(result), encoding="utf-8"
                )
                return
            time.sleep(0.1)

    commands_dir.mkdir(parents=True, exist_ok=True)
    responder = threading.Thread(target=fake_vba)
    responder.start()
    result = srv.send_and_wait("GenerateClassificationReport", timeout=10)
    responder.join()

    assert result["status"] == "ok"
    assert result["output"] == "KEEP: 5, DELETE: 3"
    # both bridge files cleaned up after a successful round-trip
    assert not list(commands_dir.iterdir())


def test_filter_last_n_days_validation(data_dir: Path) -> None:
    start = time.monotonic()
    assert call_tool("filter_last_n_days", {"days": 0}).startswith("ERROR")
    assert call_tool("filter_last_n_days", {"days": 400}).startswith("ERROR")
    assert time.monotonic() - start < 2.0  # validation must not hit the bridge


# ---------------------------------------------------------------------------
# write-backed tool
# ---------------------------------------------------------------------------


def test_add_learned_sender_rule_appends(data_dir: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(srv, "RELOAD_TIMEOUT", 0.6)  # keep best-effort reload fast
    out = call_tool(
        "add_learned_sender_rule", {"email": "new.sender@example.com", "action": "delete"}
    )
    assert "Saved rule: new.sender@example.com -> DELETE" in out
    assert "next time Outlook starts" in out  # no Outlook in tests → best-effort path

    lines = (data_dir / "learned_senders.txt").read_text(encoding="utf-8").splitlines()
    fields = lines[-1].split("|")
    assert fields[0] == "new.sender@example.com"
    assert fields[1] == "DELETE"
    assert len(fields) == 3 and len(fields[2]) == 19  # yyyy-mm-dd hh:mm:ss


def test_add_learned_sender_rule_validation(data_dir: Path) -> None:
    assert call_tool(
        "add_learned_sender_rule", {"email": "a@b.com", "action": "PURGE"}
    ).startswith("ERROR")
    assert call_tool(
        "add_learned_sender_rule", {"email": "not-an-email", "action": "KEEP"}
    ).startswith("ERROR")
    assert call_tool(
        "add_learned_sender_rule", {"email": "a|b@c.com", "action": "KEEP"}
    ).startswith("ERROR")
