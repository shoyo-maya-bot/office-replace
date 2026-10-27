"""README・--help・セットアップの案内が、コードの実物と一致することを確かめる（ドキュメントと動きのずれを見つける）.

後退: README のオプションの表に --rules・--from・--to・--target・--log-level・--version が無かった、
setup.cmd の案内が「Python 3.12」で requires-python（3.12 以上）と言い方が違っていた、など.
"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

from devtools_common.cli import ExitCode

from office_replace import cli, engine
from office_replace.rules import LONG_MATCH

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")


def _section(heading: str) -> str:
    m = re.search(rf"^## {re.escape(heading)}\n(.*?)(?=^## |\Z)", README, re.M | re.S)
    assert m, f"README に「## {heading}」がありません"
    return m.group(1)


def _norm(s: str) -> str:
    return re.sub(r"[\s`]+", "", s)


def test_every_option_is_in_readme_table():
    table = _section("オプション")
    # オプションごとに、どれか 1 つの書き方（-o / --out など）が表にあること
    actions = [a.option_strings for a in cli.build_parser()._actions if a.option_strings and "--help" not in a.option_strings]
    missing = sorted(" / ".join(opts) for opts in actions if not any(re.search(rf"`{re.escape(o)}(?![\w-])", table) for o in opts))
    assert not missing, f"README の「オプション」の表に無いオプション: {missing}"


def test_exit_codes_match_readme_and_help():
    line = next(ln for ln in _section("オプション").splitlines() if ln.startswith("終了コード:"))
    readme = dict(re.findall(r"`(\d+)` ([^／]+)", line))
    m = re.search(r"終了コード: (.*?)\n(?!\s+/)", cli.__doc__, re.S)
    assert m
    help_ = dict(re.findall(r"(\d+) ([^/]+)", m.group(1)))
    codes = {str(int(c)) for c in ExitCode}
    assert set(readme) == codes, line
    assert set(help_) == codes, m.group(1)
    for code in codes:
        assert _norm(readme[code]) == _norm(help_[code]), code


def test_zip_limit_defaults_match_readme():
    lim = engine.ZipLimits()
    row = next(ln for ln in _section("オプション").splitlines() if ln.startswith("| `--max-unzip-mb`"))
    assert f"既定 {lim.max_total // 1024**2} / {lim.max_entries} / {lim.max_ratio:g}" in row, row
    assert f"{lim.ratio_min_size // 1024**2}MB を超える部品だけ" in row, row


def test_long_match_matches_readme():
    assert f"`*` の一致が {LONG_MATCH} 文字を超えたとき" in README


def test_python_version_matches_requires_python():
    spec = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["requires-python"]
    m = re.fullmatch(r">=\s*(\d+\.\d+)", spec)
    assert m, spec
    version = m.group(1)
    assert f"Python {version} 以上" in README
    setup_cmd = (ROOT / "setup.cmd").read_bytes().decode("cp932")
    assert f"Python {version} 以上" in setup_cmd
    # CI は対応する最も古い版で動かす
    ci = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
    assert set(re.findall(r'python-version: "([\d.]+)"', ci)) == {version}


def test_version_is_dev_while_changelog_has_unreleased_changes():
    """CHANGELOG の「未リリース」に変更があるなら、版は .devN（リリース済みの版のままにしない）.

    後退: 0.2.0 をリリースした後も pyproject の版が 0.2.0 のままで、未リリースの変更が入った版と区別できなかった.
    """
    version = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    released = re.findall(r"^## (\d+\.\d+\.\d+)", changelog, re.M)
    assert released, "CHANGELOG にリリース済みの版（## 0.2.0 など）がありません"
    assert version not in released, f"{version} はリリース済みです。次の .devN にしてください"
    m = re.search(r"^## 未リリース\n(.*?)(?=^## \d)", changelog, re.M | re.S)
    if m and re.search(r"^- ", m.group(1), re.M):
        assert re.fullmatch(r"\d+\.\d+\.\d+\.dev\d+", version), f"未リリースの変更があるので .devN にしてください: {version}"
        assert f"`{version}`" in m.group(1), "「未リリース」の冒頭に今の版を書いてください"
    # .dev の版は、最新のリリースより後の版であること
    base = tuple(int(x) for x in version.split(".dev")[0].split("."))
    assert base > max(tuple(int(x) for x in r.split(".")) for r in released), version


def test_setup_scripts_find_supported_python():
    """setup.sh は requires-python と同じ版を確かめ、Git Bash（.venv/Scripts）でも動く."""
    spec = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["requires-python"]
    major, minor = re.fullmatch(r">=\s*(\d+)\.(\d+)", spec).groups()
    sh = (ROOT / "setup.sh").read_text(encoding="utf-8")
    assert f"sys.version_info >= ({major}, {minor})" in sh
    assert f"Python {major}.{minor} 以上" in sh
    assert ".venv/Scripts/python" in sh and ".venv/bin/python" in sh


def test_agent_guides_exit_codes_match_readme():
    """AGENTS.md と copilot-instructions の終了コードの説明が README と同じ."""
    line = next(ln for ln in _section("オプション").splitlines() if ln.startswith("終了コード:"))
    readme = {c: _norm(t) for c, t in re.findall(r"`(\d+)` ([^／]+)", line)}
    for name in ("AGENTS.md", ".github/copilot-instructions.md"):
        text = (ROOT / name).read_text(encoding="utf-8")
        m = re.search(r"終了コード: (.*)", text)
        assert m, name
        guide = {c: _norm(t) for c, t in re.findall(r"(\d+) ([^/]+?)(?=\s*/ \d |。|$)", m.group(1))}
        assert guide == readme, (name, guide, readme)
    assert (ROOT / "scripts" / "verify_gates.py").is_file()
