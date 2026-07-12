"""chat.py — Keyword-based command parser for the chat interface.

Maps natural language phrases to actions without requiring an LLM.
Returns an action descriptor that server.py dispatches. Every macro action
is validated against the macros.py manifest before being returned.
"""

from typing import Optional

import macros

# Map of (keywords tuple) → action descriptor
# Checked in order; first match wins.
COMMAND_MAP = [
    # --- Dry run / preview ---
    (("dry run", "preview", "what would happen", "test run"),
     {"type": "macro", "macro": "FilterExistingDryRun", "label": "Running dry run..."}),

    # --- Filtering ---
    (("filter inbox", "filter emails", "run filter", "filter all"),
     {"type": "macro", "macro": "FilterExistingEmails", "label": "Filtering inbox..."}),

    # --- Daily digest ---
    (("daily digest", "generate digest", "digest"),
     {"type": "macro", "macro": "GenerateDailyDigest", "label": "Generating daily digest..."}),

    # --- Rule proposals ---
    (("propose rules", "suggest rules", "mine rules"),
     {"type": "macro", "macro": "ProposeRules", "label": "Proposing rules from decision history..."}),

    # --- Version / status ---
    (("show version", "version", "status"),
     {"type": "macro", "macro": "ShowVersionInfo", "label": "Getting version info..."}),

    # --- Learned senders ---
    (("show senders", "learned senders", "list senders", "sender rules"),
     {"type": "api", "endpoint": "/api/learned/senders", "label": "Fetching learned senders..."}),

    # --- Learned subjects ---
    (("show subjects", "learned subjects", "subject rules"),
     {"type": "api", "endpoint": "/api/learned/subjects", "label": "Fetching learned subjects..."}),

    # --- Learned replies ---
    (("show replies", "learned replies", "reply examples"),
     {"type": "api", "endpoint": "/api/learned/replies", "label": "Fetching learned replies..."}),

    # --- Decisions ---
    (("show decisions", "decision log", "decisions"),
     {"type": "api", "endpoint": "/api/decisions", "label": "Fetching decision log..."}),

    # --- Errors / logs ---
    (("show errors", "error log", "errors", "logs"),
     {"type": "api", "endpoint": "/api/errors", "label": "Fetching error log..."}),

    # --- Settings: reload ---
    (("reload settings", "refresh settings", "reload config"),
     {"type": "macro", "macro": "ReinitializeFilter", "label": "Reloading settings..."}),

    # --- Settings: change provider ---
    (("provider local", "use local", "switch local", "change to local"),
     {"type": "setting", "section": "LLM", "key": "Provider", "value": "local",
      "label": "Switching LLM provider to local..."}),

    (("provider azure", "use azure", "switch azure", "change to azure"),
     {"type": "setting", "section": "LLM", "key": "Provider", "value": "azure",
      "label": "Switching LLM provider to azure..."}),

    (("provider claude", "use claude", "switch claude", "change to claude"),
     {"type": "setting", "section": "LLM", "key": "Provider", "value": "claude",
      "label": "Switching LLM provider to claude..."}),

    (("provider openai", "use openai", "switch openai", "use openrouter", "switch openrouter"),
     {"type": "setting", "section": "LLM", "key": "Provider", "value": "openai",
      "label": "Switching LLM provider to OpenAI-compatible..."}),

    # --- Enable / disable LLM ---
    (("enable llm", "turn on llm", "llm on"),
     {"type": "setting", "section": "LLM", "key": "UseLLMAPI", "value": "True",
      "label": "Enabling LLM..."}),

    (("disable llm", "turn off llm", "llm off"),
     {"type": "setting", "section": "LLM", "key": "UseLLMAPI", "value": "False",
      "label": "Disabling LLM..."}),

    # --- Scan sent items ---
    (("scan sent", "learn from sent", "scan replies"),
     {"type": "macro", "macro": "ScanSentForReplyPatterns", "label": "Scanning sent items..."}),

    # --- Draft replies for selected ---
    (("draft replies", "draft reply", "draft for selected"),
     {"type": "macro", "macro": "DraftReplyForSelected", "label": "Drafting replies for selected email(s)..."}),

    # --- Cloud sync ---
    (("sync rules", "cloud sync", "sync learned", "sync senders", "sync subjects"),
     {"type": "macro", "macro": "SyncLearnedRules", "label": "Syncing learned rules with cloud..."}),

    # --- Restore ---
    (("restore review", "restore from review"),
     {"type": "macro", "macro": "RestoreFromReview", "label": "Restoring from Review..."}),

    # --- Help ---
    (("help", "what can you do", "commands", "?"),
     {"type": "help", "label": "Showing help..."}),
]

HELP_TEXT = """Available commands:
• **dry run** / preview — preview filter decisions (no changes)
• **filter inbox** — filter all inbox emails
• **generate digest** / daily digest — generate today's digest now
• **propose rules** / suggest rules — mine the decision log for new rules
• **show version** / status — show version and config info
• **show senders** — view learned sender rules
• **show subjects** — view learned subject rules
• **show replies** — view learned reply examples
• **show decisions** — view the recent decision log
• **show errors** — view recent error log
• **reload settings** — reload settings.ini
• **provider local/azure/claude/openai** — switch LLM provider
• **enable/disable llm** — toggle LLM on/off
• **scan sent** — scan Sent Items for reply patterns
• **draft replies** — draft few-shot replies for selected email(s)
• **sync rules** — sync learned rules with cloud (OneDrive)
• **restore review** — move Review folder emails back to Inbox"""


def parse(message: str) -> Optional[dict]:
    """Parse a user message and return an action descriptor, or None if unknown.

    Macro actions are only returned when the macro exists in the server-side
    manifest (macros.py) — anything else is treated as unknown.
    """
    msg = message.lower().strip()
    if not msg:
        return None
    for keywords, action in COMMAND_MAP:
        if any(kw in msg for kw in keywords):
            if action["type"] == "macro" and macros.get_macro(action["macro"]) is None:
                return None
            return action
    return None


def help_text() -> str:
    return HELP_TEXT
