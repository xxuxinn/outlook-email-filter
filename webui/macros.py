"""macros.py — Server-side macro allowlist (single source of truth).

Every macro the Web UI may send over the command bridge is declared here.
server.py rejects anything not in this manifest; chat.py validates its
keyword mappings against it; the SPA renders its buttons from GET /api/macros.

Arg spec fields: name, type ("int" | "string"), required, plus constraints
(min/max for int, min_length/max_length for string).

Note: EnableRealTimeFilter / DisableRealTimeFilter are deliberately NOT here.
They live in ThisOutlookSession (a document module) and must be run in the
VBA Immediate Window; the SPA keeps them as local-only informational hints.
"""

import re
from typing import Any, Optional

_DAYS_ARG = {"name": "days", "type": "int", "required": True, "min": 1, "max": 365}
_PATTERN_ARG = {
    "name": "pattern", "type": "string", "required": True,
    "min_length": 3, "max_length": 100,
}

_FORBIDDEN_STRING_CHARS = ("\r", "\n", "\x00")
_INT_RE = re.compile(r"^-?\d+$")


def _m(name: str, label: str, description: str, category: str,
       args: Optional[list] = None, destructive: bool = False) -> dict:
    return {
        "name": name,
        "label": label,
        "description": description,
        "category": category,
        "args": args or [],
        "destructive": destructive,
    }


MACROS: list[dict] = [
    # --- Version ---
    _m("ShowVersionInfo", "Show Version Info",
       "Show version and status", "Version"),

    # --- Filtering ---
    _m("FilterExistingDryRun", "Dry Run",
       "Preview filter decisions (no changes)", "Filtering"),
    _m("FilterExistingEmails", "Filter Inbox",
       "Filter all Inbox emails", "Filtering", destructive=True),
    _m("FilterAllFolders", "Filter All Folders",
       "Filter Inbox + Other + PST archives", "Filtering", destructive=True),
    _m("FilterSelectedEmail", "Test Selected Email",
       "Test classification on one selected email", "Filtering"),
    _m("FilterSelectedEmails", "Filter Selected",
       "Filter selected email(s) with confirmation", "Filtering", destructive=True),
    _m("FilterCurrentFolder", "Filter Current Folder",
       "Filter current folder with confirmation", "Filtering", destructive=True),
    _m("FilterLastNDays", "Filter Last N Days",
       "Filter emails from the last N days", "Filtering",
       args=[_DAYS_ARG], destructive=True),
    _m("GenerateClassificationReport", "Classification Report",
       "Count classifications without acting", "Filtering"),
    _m("BulkDeleteBySender", "Bulk Delete By Sender",
       "Delete all from matching senders", "Filtering",
       args=[_PATTERN_ARG], destructive=True),
    _m("MoveProtectedSources", "Move Protected Sources",
       "Move protected domain emails to Protected folder", "Filtering",
       destructive=True),

    # --- Digest & Proposals ---
    _m("GenerateDailyDigest", "Generate Daily Digest",
       "Generate today's digest markdown file now", "Digest & Proposals"),
    _m("ProposeRules", "Propose Rules",
       "LLM proposes new rules from the decision log", "Digest & Proposals"),

    # --- Agent Tools ---
    _m("GenerateAddressingPatterns", "Generate Addressing Patterns",
       "LLM-generate name/greeting patterns", "Agent Tools"),
    _m("ScanSentForReplyPatterns", "Scan Sent Items",
       "Learn reply style from Sent Items", "Agent Tools"),
    _m("DraftReplyForSelected", "Draft Reply For Selected",
       "Draft few-shot replies for selected email(s)", "Agent Tools"),
    _m("ShowLearnedRepliesSummary", "Learned Replies Summary",
       "Show learned reply pair count", "Agent Tools"),

    # --- LLM Tools ---
    _m("SummarizeSelectedEmail", "Summarize Selected Email",
       "Summarize selected email using LLM", "LLM Tools"),
    _m("DraftReplyToSelected", "Draft Reply To Selected",
       "Draft reply to selected email (few-shot)", "LLM Tools"),

    # --- Learned Rules ---
    _m("ShowLearnedSenders", "Learned Senders Count",
       "Show learned sender rule count", "Learned Rules"),
    _m("ShowLearnedSendersList", "Dump Sender Rules",
       "Dump sender rules to Immediate Window", "Learned Rules"),
    _m("ReloadLearnedSenders", "Reload Learned Senders",
       "Force reload learned rules from file", "Learned Rules"),
    _m("CleanLearnedSendersFile", "Clean Senders File",
       "Remove duplicate sender entries", "Learned Rules", destructive=True),
    _m("ImportExistingLearnedFolders", "Import Learned Folders",
       "Bulk import from LearnKeep/LearnDelete", "Learned Rules"),
    _m("ShowLearnedSubjectsList", "Dump Subject Rules",
       "Dump subject rules to Immediate Window", "Learned Rules"),
    _m("CleanLearnedSubjectsFile", "Clean Subjects File",
       "Remove duplicate subject entries", "Learned Rules", destructive=True),
    _m("ImportExistingLearnedSubjectFolder", "Import Subject Folder",
       "Bulk import from LearnSubjectDelete", "Learned Rules"),

    # --- Cloud Sync ---
    _m("SyncLearnedRules", "Sync Learned Rules",
       "Bidirectional sync of learned rules with cloud folder", "Cloud Sync"),

    # --- Server Rules ---
    _m("ImportServerRules", "Import Server Rules",
       "Import server rules as learned DELETE rules", "Server Rules"),
    _m("ExportLearnedRulesToServer", "Export Rules To Server",
       "Push DELETE rules to Exchange server", "Server Rules", destructive=True),

    # --- Undo / Recovery ---
    _m("RestoreFromReview", "Restore From Review",
       "Move Review folder emails back to Inbox", "Undo / Recovery",
       destructive=True),
    _m("RestoreDeletedKeepEmails", "Restore Deleted KEEP Emails",
       "Rescue wrongly deleted KEEP emails", "Undo / Recovery",
       destructive=True),

    # --- Migration & System ---
    _m("DetectAndMigrateOldFolders", "Migrate Old Folders",
       "Rename v1.x folders to v2.0 names", "Migration & System",
       destructive=True),
    _m("ReinitializeFilter", "Reinitialize Filter",
       "Restart event handlers, command poller, reload settings",
       "Migration & System"),
]

MACROS_BY_NAME: dict[str, dict] = {m["name"]: m for m in MACROS}


def get_macro(name: str) -> Optional[dict]:
    return MACROS_BY_NAME.get(name)


def validate_command(name: str, args: Any) -> tuple[Optional[dict], Optional[str]]:
    """Validate a macro name + args against the manifest.

    Returns (clean_args, None) on success or (None, error_message) on failure.
    Validated values are serialized back to strings for the command file so
    the wire format matches what the VBA bridge has always received.
    """
    macro = get_macro(name)
    if macro is None:
        return None, f"Unknown or disallowed macro: {name}"

    if args is None:
        args = {}
    if not isinstance(args, dict):
        return None, "args must be a JSON object"

    declared = {spec["name"]: spec for spec in macro["args"]}
    unknown = sorted(k for k in args if k not in declared)
    if unknown:
        return None, f"Undeclared argument(s): {', '.join(unknown)}"

    clean: dict[str, str] = {}
    for spec in macro["args"]:
        arg_name = spec["name"]
        if arg_name not in args or args[arg_name] in (None, ""):
            if spec.get("required"):
                return None, f"Missing required argument: {arg_name}"
            continue
        value = args[arg_name]
        validated, err = _validate_value(spec, value)
        if err:
            return None, err
        clean[arg_name] = validated
    return clean, None


def _validate_value(spec: dict, value: Any) -> tuple[str, Optional[str]]:
    name = spec["name"]
    if spec["type"] == "int":
        if isinstance(value, bool):
            return "", f"Argument '{name}' must be an integer"
        if isinstance(value, int):
            num = value
        elif isinstance(value, str) and _INT_RE.match(value.strip()):
            num = int(value.strip())
        else:
            return "", f"Argument '{name}' must be an integer"
        lo, hi = spec.get("min"), spec.get("max")
        if lo is not None and num < lo or hi is not None and num > hi:
            return "", f"Argument '{name}' must be between {lo} and {hi}"
        return str(num), None

    # string
    if not isinstance(value, str):
        return "", f"Argument '{name}' must be a string"
    if any(ch in value for ch in _FORBIDDEN_STRING_CHARS):
        return "", f"Argument '{name}' must not contain control characters"
    lo = spec.get("min_length", 0)
    hi = spec.get("max_length", 10_000)
    if not (lo <= len(value) <= hi):
        return "", f"Argument '{name}' must be {lo}-{hi} characters"
    return value, None
