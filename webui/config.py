r"""config.py — Path configuration for the Web UI.

All paths are resolved at call time (functions, not constants) so tests can
inject a temporary data directory via the OUTLOOK_FILTER_DATA_DIR environment
variable without reloading modules.

Default data directory: %APPDATA%\OutlookEmailFilter\
"""

import os

ENV_DATA_DIR = "OUTLOOK_FILTER_DATA_DIR"


def data_dir() -> str:
    """Root data directory (settings, learned rules, logs, commands)."""
    override = os.environ.get(ENV_DATA_DIR, "")
    if override:
        return override
    return os.path.join(os.environ.get("APPDATA", ""), "OutlookEmailFilter")


def settings_path() -> str:
    return os.path.join(data_dir(), "settings.ini")


def commands_dir() -> str:
    return os.path.join(data_dir(), "commands")


def digests_dir() -> str:
    return os.path.join(data_dir(), "digests")


def token_path() -> str:
    return os.path.join(data_dir(), "webui_token.txt")


def data_file(name: str) -> str:
    """Absolute path of a file inside the data directory."""
    return os.path.join(data_dir(), name)
