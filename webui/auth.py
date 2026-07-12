"""auth.py — Token authentication for the Web UI.

A random 32-hex-char token is stored at <DATA_DIR>/webui_token.txt (created
on first use, 0600 permissions best-effort). The SPA receives the token via
template substitution when index.html is served; every /api/* request must
carry it in the X-Auth-Token header.
"""

import hmac
import os
import re
import secrets

import config

_TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")

# Cache keyed by token file path so tests with different data dirs
# each get their own token.
_token_cache: dict[str, str] = {}


def get_token() -> str:
    """Load the auth token from disk, creating it if missing/invalid."""
    path = config.token_path()
    cached = _token_cache.get(path)
    if cached:
        return cached

    token = _read_token(path)
    if token is None:
        token = secrets.token_hex(16)  # 32 hex chars
        _write_token(path, token)

    _token_cache[path] = token
    return token


def check_token(provided: str | None) -> bool:
    """Constant-time comparison of a provided token against the real one."""
    if not provided:
        return False
    return hmac.compare_digest(str(provided), get_token())


def _read_token(path: str) -> str | None:
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="ascii", errors="strict") as f:
            candidate = f.read().strip()
    except (OSError, UnicodeDecodeError):
        return None
    if _TOKEN_RE.match(candidate):
        return candidate
    return None


def _write_token(path: str, token: str) -> None:
    os.makedirs(config.data_dir(), exist_ok=True)
    with open(path, "w", encoding="ascii") as f:
        f.write(token)
    try:
        os.chmod(path, 0o600)  # best-effort; no-op semantics on Windows
    except OSError:
        pass
