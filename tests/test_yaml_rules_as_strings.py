"""YAML の置換表: 値を書いたとおりの文字列で読む（0123 を 83、no を False にしない）."""

from __future__ import annotations

from pathlib import Path

from office_replace.rules import load


def test_yaml_values_are_read_as_written(tmp_path: Path):
    p = tmp_path / "r.yaml"
    p.write_text(
        "rules:\n"
        "  - 置換前: 0123\n    置換後: no\n"
        "  - before: yes\n    after: 1.50\n    targets: [text, sheet]\n"
        "  - before: 2026-01-01\n    after: 0x1F\n    完全一致: yes\n"
        "  - before: 削除\n    after:\n"
        "  - before: 消す\n    after: ~\n",
        encoding="utf-8",
    )
    rules = load(p)
    assert [(r.before, r.after) for r in rules] == [
        ("0123", "no"),
        ("yes", "1.50"),
        ("2026-01-01", "0x1F"),
        ("削除", ""),
        ("消す", ""),  # 空・~ は空の置換後（削除）
    ]
    assert rules[1].targets == ("text", "sheet")
    assert rules[2].whole


def test_yaml_with_bom_and_top_level_list(tmp_path: Path):
    p = tmp_path / "r.yml"
    p.write_text("- {before: '007', after: 008}\n", encoding="utf-8-sig")
    assert [(r.before, r.after) for r in load(p)] == [("007", "008")]
