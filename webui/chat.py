"""chat.py — Keyword-based command parser for the chat interface.

Maps natural language phrases to actions without requiring an LLM.
Returns an action descriptor that server.py dispatches.
"""

from typing import Optional

# Map of (keywords tuple) → action descriptor
# Checked in order; first match wins.
COMMAND_MAP = [
    # --- Dry run / preview ---
    (("dry run", "preview", "what would happen", "test run"),
     {"type": "macro", "macro": "FilterExistingDryRun", "label": "Running dry run..."}),

    # --- Filtering ---
    (("filter inbox", "filter emails", "run filter", "filter all"),
     {"type": "macro", "macro": "FilterExistingEmails", "label": "Filtering inbox..."}),

    # --- Version / status ---
    (("version", "status", "show version", "info"),
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
• **show version** / status — show version and config info
• **show senders** — view learned sender rules
• **show subjects** — view learned subject rules
• **show replies** — view learned reply examples
• **show errors** — view recent error log
• **reload settings** — reload settings.ini
• **provider local/azure/claude/openai** — switch LLM provider
• **enable/disable llm** — toggle LLM on/off
• **scan sent** — scan Sent Items for reply patterns
• **draft replies** — draft few-shot replies for selected email(s)
• **restore review** — move Review folder emails back to Inbox"""


def parse(message: str) -> Optional[dict]:
    """Parse a user message and return an action descriptor, or None if unknown."""
    msg = message.lower().strip()
    for keywords, action in COMMAND_MAP:
        if any(kw in msg for kw in keywords):
            return action
    return None


def help_text() -> str:
    return HELP_TEXT
