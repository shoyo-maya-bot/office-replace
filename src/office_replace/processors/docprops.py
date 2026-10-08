"""文書のプロパティ・グラフのデータ（Excel / Word / PowerPoint 共通）."""

from __future__ import annotations

import posixpath

from ..ooxml import NS, Change, Package
from ..rules import Rule
from .common import _parts, _q, _replace_value

CORE_FIELDS = {
    ("dc", "title"): "タイトル",
    ("dc", "subject"): "件名",
    ("dc", "creator"): "作成者",
    ("cp", "keywords"): "キーワード",
    ("dc", "description"): "コメント",
    ("cp", "lastModifiedBy"): "最終更新者",
    ("cp", "category"): "分類",
}
APP_FIELDS = {"Company": "会社", "Manager": "管理者"}


def properties(pkg: Package, rules: list[Rule], titles: bool) -> list[Change]:
    """文書のプロパティ（タイトル・作成者・会社など）. titles: 「パーツのタイトル」（スライドの題など）も置換する."""
    changes: list[Change] = []
    if pkg.has("docProps/core.xml"):
        root = pkg.xml("docProps/core.xml").getroot()
        for (prefix, tag), label in CORE_FIELDS.items():
            for el in root.iterfind(f"{prefix}:{tag}", NS):
                changes += _replace_value(pkg, "docProps/core.xml", el, rules, f"文書のプロパティ（{label}）")
    if pkg.has("docProps/app.xml"):
        root = pkg.xml("docProps/app.xml").getroot()
        for tag, label in APP_FIELDS.items():
            for el in root.iterfind(f"ep:{tag}", NS):
                changes += _replace_value(pkg, "docProps/app.xml", el, rules, f"文書のプロパティ（{label}）")
        if titles:  # Excel ではシート名の一覧なので、シート名の変更に合わせて別に直す
            for el in root.iterfind("ep:TitlesOfParts/vt:vector/vt:lpstr", NS):
                changes += _replace_value(pkg, "docProps/app.xml", el, rules, "文書のプロパティ（パーツのタイトル）")
    if pkg.has("docProps/custom.xml"):
        for prop in pkg.xml("docProps/custom.xml").getroot().iterfind("op:property", NS):
            for el in prop.iterfind("vt:lpwstr", NS):
                changes += _replace_value(
                    pkg, "docProps/custom.xml", el, rules, f"文書のプロパティ（ユーザー設定「{prop.get('name', '')}」）"
                )
    return changes


_STR_CACHES = {_q("c", "strCache"), _q("c", "multiLvlStrCache")}


def chart_caches(pkg: Package, rules: list[Rule]) -> list[Change]:
    """グラフに保存されている項目名・系列名（c:strCache / c:tx の c:v）. 数値（c:numCache）は変えない."""
    changes: list[Change] = []
    for part in _parts(pkg, r"(xl|word|ppt)/charts/chart[^/]*\.xml"):
        label = f"グラフ {posixpath.basename(part).removesuffix('.xml')} の項目名・系列名"
        for v in pkg.xml(part).getroot().iter(_q("c", "v")):
            parent = v.getparent()
            in_cache = any(a.tag in _STR_CACHES for a in v.iterancestors())
            if in_cache or (parent is not None and parent.tag == _q("c", "tx")):
                changes += _replace_value(pkg, part, v, rules, label)
    return changes
