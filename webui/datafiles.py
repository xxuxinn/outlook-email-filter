r"""datafiles.py — Read/write helpers for the agent's data files.

Covers pipe-delimited files (learned rules, decision log, rule proposals),
log tails, and digest markdown files under %APPDATA%\OutlookEmailFilter\.

rule_proposals.txt format:
    id|type|value|action|reason|status|timestamp
    type ∈ SENDER,SUBJECT; action ∈ KEEP,DELETE;
    status ∈ PENDING,APPROVED,REJECTED

decision_log.txt format:
    timestamp|senderEmail|subject|source|action|confidence
"""

import os
import re
from datetime import datetime
from typing import Optional

import config

TIMESTAMP_FMT = "%Y-%m-%d %H:%M:%S"
PROPOSAL_FIELDS = 7
_DIGEST_RE = re.compile(r"^digest_(\d{4}-\d{2}-\d{2})\.md$")

PROPOSAL_TYPES = {"SENDER", "SUBJECT"}
PROPOSAL_ACTIONS = {"KEEP", "DELETE"}


# ---------------------------------------------------------------------------
# Generic readers
# ---------------------------------------------------------------------------

def read_pipe_file(filename: str, max_lines: int = 500) -> list[list[str]]:
    """Read a pipe-delimited data file; return list of field lists."""
    path = config.data_file(filename)
    if not os.path.exists(path):
        return []
    rows = []
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            rows.append(stripped.split("|"))
    return rows[-max_lines:]


def read_log(filename: str, max_lines: int = 200) -> list[str]:
    path = config.data_file(filename)
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = f.readlines()
    return [l.rstrip() for l in lines[-max_lines:]]


def count_rules(filename: str) -> int:
    """Count distinct rule keys (first pipe field) in a learned-rules file."""
    path = config.data_file(filename)
    if not os.path.exists(path):
        return 0
    seen = set()
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                seen.add(stripped.split("|")[0].lower())
    return len(seen)


# ---------------------------------------------------------------------------
# Digest
# ---------------------------------------------------------------------------

def latest_digest() -> Optional[dict]:
    """Return {"date", "content"} for the newest digests/digest_*.md, or None."""
    directory = config.digests_dir()
    if not os.path.isdir(directory):
        return None
    dated = []
    for fname in os.listdir(directory):
        match = _DIGEST_RE.match(fname)
        if match:
            dated.append((match.group(1), fname))
    if not dated:
        return None
    date, fname = sorted(dated)[-1]  # YYYY-MM-DD sorts lexicographically
    try:
        with open(os.path.join(directory, fname), "r",
                  encoding="utf-8-sig", errors="replace") as f:
            content = f.read()
    except OSError:
        return None
    return {"date": date, "content": content}


def latest_digest_date() -> Optional[str]:
    digest = latest_digest()
    return digest["date"] if digest else None


# ---------------------------------------------------------------------------
# Decision log
# ---------------------------------------------------------------------------

def read_decisions(max_rows: int = 100) -> list[dict]:
    """Return the last N parsed decision_log.txt rows (newest last)."""
    rows = read_pipe_file("decision_log.txt", max_lines=max_rows)
    decisions = []
    for parts in rows:
        if len(parts) < 6:
            continue
        decisions.append({
            "timestamp": parts[0].strip(),
            "sender": parts[1].strip(),
            "subject": parts[2].strip(),
            "source": parts[3].strip(),
            "action": parts[4].strip(),
            "confidence": parts[5].strip(),
        })
    return decisions


# ---------------------------------------------------------------------------
# Rule proposals
# ---------------------------------------------------------------------------

def read_proposals() -> list[dict]:
    """Parse rule_proposals.txt into dicts (malformed lines skipped)."""
    proposals = []
    for parts in read_pipe_file("rule_proposals.txt", max_lines=1000):
        if len(parts) < PROPOSAL_FIELDS:
            continue
        proposals.append(_proposal_dict(parts))
    return proposals


def sanitize_rule_value(value: str) -> str:
    """Strip pipe, CR, LF and null from a rule value before file append."""
    cleaned = str(value)
    for ch in ("|", "\r", "\n", "\x00"):
        cleaned = cleaned.replace(ch, " ")
    return " ".join(cleaned.split())


def approve_proposal(proposal_id: str) -> tuple[Optional[dict], Optional[str], int]:
    """Approve a PENDING proposal: append the learned rule, mark APPROVED.

    Returns (proposal, error_message, http_status).
    """
    proposal = _find_proposal(proposal_id)
    if proposal is None:
        return None, "Proposal not found", 404
    if proposal["status"] != "PENDING":
        return None, f"Proposal is {proposal['status']}, not PENDING", 400
    if proposal["type"] not in PROPOSAL_TYPES:
        return None, f"Unknown proposal type: {proposal['type']}", 400

    value = sanitize_rule_value(proposal["value"])
    if not value:
        return None, "Proposal value is empty after sanitizing", 400
    timestamp = datetime.now().strftime(TIMESTAMP_FMT)

    if proposal["type"] == "SENDER":
        action = proposal["action"].upper()
        if action not in PROPOSAL_ACTIONS:
            return None, f"Invalid proposal action: {proposal['action']}", 400
        _append_line("learned_senders.txt", f"{value}|{action}|{timestamp}")
    else:  # SUBJECT — only DELETE is valid for subject rules
        _append_line("learned_subjects.txt", f"{value}|DELETE|{timestamp}")

    updated = _set_proposal_status(proposal_id, "APPROVED")
    if updated is None:
        return None, "Failed to update proposal status", 500
    return updated, None, 200


def reject_proposal(proposal_id: str) -> tuple[Optional[dict], Optional[str], int]:
    """Mark a PENDING proposal as REJECTED. Returns (proposal, error, status)."""
    proposal = _find_proposal(proposal_id)
    if proposal is None:
        return None, "Proposal not found", 404
    if proposal["status"] != "PENDING":
        return None, f"Proposal is {proposal['status']}, not PENDING", 400
    updated = _set_proposal_status(proposal_id, "REJECTED")
    if updated is None:
        return None, "Failed to update proposal status", 500
    return updated, None, 200


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _proposal_dict(parts: list[str]) -> dict:
    return {
        "id": parts[0].strip(),
        "type": parts[1].strip().upper(),
        "value": parts[2].strip(),
        "action": parts[3].strip().upper(),
        "reason": parts[4].strip(),
        "status": parts[5].strip().upper(),
        "timestamp": parts[6].strip(),
    }


def _find_proposal(proposal_id: str) -> Optional[dict]:
    for proposal in read_proposals():
        if proposal["id"] == proposal_id:
            return proposal
    return None


def _set_proposal_status(proposal_id: str, new_status: str) -> Optional[dict]:
    """Rewrite rule_proposals.txt with the given proposal's status changed."""
    path = config.data_file("rule_proposals.txt")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = f.read().splitlines()

    updated: Optional[dict] = None
    new_lines = []
    for line in lines:
        parts = line.split("|")
        if len(parts) >= PROPOSAL_FIELDS and parts[0].strip() == proposal_id:
            parts = [*parts[:5], new_status, *parts[6:]]
            updated = _proposal_dict(parts)
            new_lines.append("|".join(parts))
        else:
            new_lines.append(line)

    if updated is None:
        return None
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(new_lines) + "\n")
    return updated


def _append_line(filename: str, line: str) -> None:
    path = config.data_file(filename)
    os.makedirs(config.data_dir(), exist_ok=True)
    with open(path, "a", encoding="utf-8", newline="") as f:
        f.write(line + "\n")
