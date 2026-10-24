"""Checks that every VBA module imports cleanly through the VBA Editor.

The VBA Editor reads .bas files in the Windows ANSI code page and names each
module from its `Attribute VB_Name` line, so sources must be ASCII-only and
carry that header (ThisOutlookSession.bas is pasted, not imported).
"""
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
BAS_FILES = sorted((ROOT / "src").glob("*.bas"))
PASTED_MODULES = {"ThisOutlookSession"}


@pytest.mark.parametrize("path", BAS_FILES, ids=lambda p: p.name)
def test_bas_is_ascii_only(path: Path) -> None:
    data = path.read_bytes()
    bad = [i for i, b in enumerate(data) if b > 127]
    assert not bad, f"{path.name}: non-ASCII byte at offset {bad[0]}"


@pytest.mark.parametrize(
    "path", [p for p in BAS_FILES if p.stem not in PASTED_MODULES], ids=lambda p: p.name
)
def test_bas_has_vb_name_header(path: Path) -> None:
    first = path.read_text(encoding="ascii").splitlines()[0]
    assert first == f'Attribute VB_Name = "{path.stem}"'


def test_bas_checked_out_with_crlf() -> None:
    attrs = (ROOT / ".gitattributes").read_text(encoding="utf-8")
    assert "*.bas text eol=crlf" in attrs
