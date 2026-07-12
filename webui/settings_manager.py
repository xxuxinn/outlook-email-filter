"""settings_manager.py — Read/write settings.ini for the Outlook Email Agent.

Encoding: the VBA side writes settings.ini as UTF-8 with BOM. Legacy files may
still be in the Windows ANSI codepage (cp950 on Traditional-Chinese systems),
so reads try utf-8-sig → cp950 → latin-1 and log which decoder was used.
Writes are always utf-8-sig.
"""

import configparser
import logging
import os
import re
from typing import Any

import config

logger = logging.getLogger(__name__)

KNOWN_SECTIONS = frozenset(
    {"General", "Folders", "Patterns", "LLM", "Agent", "Sync", "Digest"}
)
SECRET_KEYS = frozenset({"APIKeyHardcoded"})
MASK_VALUE = "__MASKED__"

_KEY_RE = re.compile(r"^[A-Za-z0-9_]{1,64}$")
_READ_ENCODINGS = ("utf-8-sig", "cp950", "latin-1")


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_setting(section: str, key: str, value: Any) -> None:
    """Raise ValueError if a section/key/value triple is not writable."""
    if section not in KNOWN_SECTIONS:
        raise ValueError(f"Unknown section: {section!r}")
    if not isinstance(key, str) or not _KEY_RE.match(key):
        raise ValueError("Key must match ^[A-Za-z0-9_]{1,64}$")
    text = str(value)
    if "\r" in text or "\n" in text:
        raise ValueError("Value must not contain newlines")


def is_masked_secret_write(key: str, value: Any) -> bool:
    """True when a write should be silently skipped (masked secret echo)."""
    return key in SECRET_KEYS and str(value) == MASK_VALUE


# ---------------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------------

def _read_text_with_fallback(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    for encoding in _READ_ENCODINGS:
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding == "utf-8-sig":
            logger.debug("settings.ini decoded as utf-8-sig")
        else:
            logger.info("settings.ini decoded with fallback encoding %s", encoding)
        return text
    # latin-1 never raises, so this is unreachable; kept for safety.
    return raw.decode("latin-1", errors="replace")


def _load() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.optionxform = str  # preserve key case
    path = config.settings_path()
    if os.path.exists(path):
        try:
            cfg.read_string(_read_text_with_fallback(path))
        except (OSError, configparser.Error) as e:
            logger.error("Failed to parse settings.ini: %s", e)
    return cfg


def read_all(mask_secrets: bool = False) -> dict:
    """Return all settings as {section: {key: value}}.

    With mask_secrets=True, non-empty secret values (APIKeyHardcoded) are
    replaced by the literal "__MASKED__" so they never reach the browser.
    """
    cfg = _load()
    result = {}
    for section in cfg.sections():
        values = dict(cfg[section])
        if mask_secrets:
            values = {
                key: (MASK_VALUE if key in SECRET_KEYS and value else value)
                for key, value in values.items()
            }
        result[section] = values
    return result


def read_setting(section: str, key: str, default: str = "") -> str:
    cfg = _load()
    return cfg.get(section, key, fallback=default)


# ---------------------------------------------------------------------------
# Write (always utf-8-sig, matching the VBA side)
# ---------------------------------------------------------------------------

def _save(cfg: configparser.ConfigParser) -> None:
    os.makedirs(config.data_dir(), exist_ok=True)
    with open(config.settings_path(), "w", encoding="utf-8-sig") as f:
        cfg.write(f)


def write_setting(section: str, key: str, value: Any) -> None:
    """Write a single key/value (read-modify-write). Raises ValueError."""
    validate_setting(section, key, value)
    if is_masked_secret_write(key, value):
        return  # no-op: the client echoed the mask back
    cfg = _load()
    if not cfg.has_section(section):
        cfg.add_section(section)
    cfg.set(section, key, str(value))
    _save(cfg)


def write_section(section: str, data: dict) -> None:
    """Write all key/value pairs in a section. Validates all pairs first."""
    for key, value in data.items():
        validate_setting(section, key, value)
    writable = {
        key: value for key, value in data.items()
        if not is_masked_secret_write(key, value)
    }
    cfg = _load()
    if not cfg.has_section(section):
        cfg.add_section(section)
    for key, value in writable.items():
        cfg.set(section, key, str(value))
    _save(cfg)


# ---------------------------------------------------------------------------
# Paths (delegated to config for backwards compatibility)
# ---------------------------------------------------------------------------

def settings_path() -> str:
    return config.settings_path()


def data_dir() -> str:
    return config.data_dir()
