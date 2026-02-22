"""settings_manager.py — Read/write settings.ini for the Outlook Email Agent."""

import os
import configparser
from typing import Any

SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", ""), "OutlookEmailFilter")
SETTINGS_PATH = os.path.join(SETTINGS_DIR, "settings.ini")

SECTIONS = ["General", "Folders", "Patterns", "LLM", "Agent"]


def _load() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.optionxform = str  # preserve key case
    if os.path.exists(SETTINGS_PATH):
        cfg.read(SETTINGS_PATH, encoding="utf-8")
    return cfg


def read_all() -> dict:
    """Return all settings as a nested dict {section: {key: value}}."""
    cfg = _load()
    result = {}
    for section in cfg.sections():
        result[section] = dict(cfg[section])
    return result


def read_setting(section: str, key: str, default: str = "") -> str:
    cfg = _load()
    return cfg.get(section, key, fallback=default)


def write_setting(section: str, key: str, value: Any) -> None:
    """Write a single key/value to settings.ini (read-modify-write)."""
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    cfg = _load()
    if not cfg.has_section(section):
        cfg.add_section(section)
    cfg.set(section, key, str(value))
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)


def write_section(section: str, data: dict) -> None:
    """Write all key/value pairs in a section (read-modify-write)."""
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    cfg = _load()
    if not cfg.has_section(section):
        cfg.add_section(section)
    for key, value in data.items():
        cfg.set(section, key, str(value))
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)


def settings_path() -> str:
    return SETTINGS_PATH


def data_dir() -> str:
    return SETTINGS_DIR
