"""ファイルの種類ごとの置換（Excel / Word / PowerPoint）.

形式ごとのモジュールに分けている（振る舞いは分割前の processors.py と同じ）:

- common: 段落・値の置換など共通の部品
- docprops: 文書のプロパティ・グラフのデータ（3 形式共通）
- spreadsheet: Excel（セル・図形・コメント・ヘッダー／フッター・テーブル・シート名）
- formula: Excel の数式（数式を持つ場所・構造化参照・数式の中の文字列）
- wordprocessing: Word
- presentation: PowerPoint
"""

from __future__ import annotations

from .common import MC_FALLBACK, Q
from .docprops import APP_FIELDS, CORE_FIELDS, chart_caches, properties
from .formula import SHEET_FORMULA_XPATHS, rename_table_columns, replace_literals
from .presentation import powerpoint
from .spreadsheet import excel
from .wordprocessing import WORD_PARTS, word

__all__ = [
    "APP_FIELDS",
    "CORE_FIELDS",
    "MC_FALLBACK",
    "Q",
    "SHEET_FORMULA_XPATHS",
    "WORD_PARTS",
    "chart_caches",
    "excel",
    "powerpoint",
    "properties",
    "rename_table_columns",
    "replace_literals",
    "word",
]
