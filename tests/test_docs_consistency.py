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
