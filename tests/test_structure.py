"""processors の分割（形式ごとのモジュール）. 振る舞いは既存のテストで確かめ、ここでは構成と公開名を確かめる."""

from __future__ import annotations

import importlib
from pathlib import Path

import office_replace.processors as processors

PKG = Path(processors.__file__).resolve().parent
MODULES = ("common", "docprops", "spreadsheet", "formula", "wordprocessing", "presentation")


def test_processors_is_split_by_format():
    assert PKG.name == "processors" and (PKG / "__init__.py").is_file()
    for name in MODULES:
        mod = importlib.import_module(f"office_replace.processors.{name}")
        lines = Path(mod.__file__).read_text(encoding="utf-8").count("\n")
        assert lines < 400, f"{name}.py が {lines} 行（形式ごとに分けて保つ）"


def test_public_names_are_kept():
    # engine とテストが使う名前は、分割前と同じく office_replace.processors から使える
    for name in ("excel", "word", "powerpoint", "chart_caches", "properties", "rename_table_columns", "replace_literals"):
        assert callable(getattr(processors, name))
