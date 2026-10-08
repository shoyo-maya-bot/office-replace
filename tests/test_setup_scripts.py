"""セットアップのスクリプトが constraints.txt で依存の版をそろえる（CI と同じ版を入れる）."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_setup_sh_uses_constraints():
    text = (ROOT / "setup.sh").read_text(encoding="utf-8")
    install = [ln for ln in text.splitlines() if "pip install" in ln and "-e" in ln]
    assert install and all("$C" in ln for ln in install)
    assert '[ -f constraints.txt ] && C="-c constraints.txt"' in text


def test_setup_cmd_uses_constraints():
    raw = (ROOT / "setup.cmd").read_bytes()
    assert b"\r\n" in raw  # Windows の改行を保つ
    text = raw.decode("cp932")
    install = [ln for ln in text.splitlines() if "pip install" in ln and "-e" in ln]
    assert install and all("%C%" in ln for ln in install)
    assert "if exist constraints.txt set C=-c constraints.txt" in text
