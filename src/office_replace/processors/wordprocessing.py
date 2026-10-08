"""Word: 本文・脚注・コメント・ヘッダー／フッター・図形・グラフ."""

from __future__ import annotations

import posixpath
import re

from lxml import etree

from ..ooxml import NS, Change, Package
from ..rules import Rule
from .common import _parts, _q, _replace_paragraphs, _replace_value

WORD_PARTS = {
    "word/document.xml": "本文",
    "word/footnotes.xml": "脚注",
    "word/endnotes.xml": "文末脚注",
    "word/comments.xml": "コメント",
}


def word(pkg: Package, text_rules: list[Rule]) -> list[Change]:
    changes: list[Change] = []
    parts = dict(WORD_PARTS)
    for n in pkg.names:
        m = re.fullmatch(r"word/(header|footer)(\d+)\.xml", n)
        if m:
            parts[n] = ("ヘッダー" if m.group(1) == "header" else "フッター") + m.group(2)
    for part, label in parts.items():
        if not pkg.has(part):
            continue
        changes += _replace_paragraphs(
            pkg,
            part,
            _q("w", "p"),
            _q("w", "t"),
            text_rules,
            lambda p, n, lb=label: f"{lb} 段落 {n}" + ("（テキストボックス）" if _in_textbox(p) else ""),
        )
        # 図形（DrawingML）の中の文字
        changes += _replace_paragraphs(pkg, part, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, lb=label: f"{lb} 図形")
    for part in _parts(pkg, r"word/charts/chart[^/]*\.xml"):
        label = posixpath.basename(part).removesuffix(".xml")
        changes += _replace_paragraphs(pkg, part, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, lb=label: f"グラフ {lb}")
    # コメントの作成者
    if pkg.has("word/comments.xml"):
        for c in pkg.xml("word/comments.xml").getroot().iterfind("w:comment", NS):
            changes += _replace_value(pkg, "word/comments.xml", c, text_rules, "コメントの作成者", attr=_q("w", "author"))
    if pkg.has("word/people.xml"):
        for person in pkg.xml("word/people.xml").getroot().iterfind("w15:person", NS):
            changes += _replace_value(pkg, "word/people.xml", person, text_rules, "コメントの作成者", attr=_q("w15", "author"))
    return changes


def _in_textbox(p: etree._Element) -> bool:
    return any(a.tag == _q("w", "txbxContent") for a in p.iterancestors())
