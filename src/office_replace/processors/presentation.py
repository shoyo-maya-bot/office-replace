"""PowerPoint: スライド・ノート・グラフ・コメント・マスター."""

from __future__ import annotations

import posixpath
import re

from ..ooxml import NS, Change, Package
from ..rules import Rule
from .common import _parts, _q, _related, _replace_paragraphs, _replace_value, _root_is, _shape_name


def _slide_order(pkg: Package) -> list[str]:
    pres = pkg.xml("ppt/presentation.xml").getroot()
    rels = pkg.rels("ppt/presentation.xml")
    return [rels.get(s.get(_q("r", "id")), "") for s in pres.iterfind("p:sldIdLst/p:sldId", NS)]


def powerpoint(pkg: Package, text_rules: list[Rule], include_masters: bool = False) -> list[Change]:
    changes: list[Change] = []
    for no, slide in enumerate(_slide_order(pkg), start=1):
        if not pkg.has(slide):
            continue
        changes += _replace_paragraphs(
            pkg, slide, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, no=no: f"スライド {no} 図形「{_shape_name(p)}」"
        )
        for target in pkg.rels(slide).values():
            if pkg.has(target) and ("/notesSlides/" in target or "/charts/" in target):
                label = "ノート" if "/notesSlides/" in target else "グラフ"
                changes += _replace_paragraphs(
                    pkg, target, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, no=no, lb=label: f"スライド {no} {lb}"
                )
    changes += _ppt_comments(pkg, text_rules)
    if include_masters:
        for part in sorted(n for n in pkg.names if re.fullmatch(r"ppt/(slideMasters/slideMaster|slideLayouts/slideLayout)\d+\.xml", n)):
            label = posixpath.basename(part).removesuffix(".xml")
            changes += _replace_paragraphs(
                pkg, part, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, lb=label: f"{lb} 図形「{_shape_name(p)}」"
            )
    return changes


def _ppt_comments(pkg: Package, rules: list[Rule]) -> list[Change]:
    """スライドのコメント（従来の形式と、新しい形式の返信を含む）とその作成者."""
    changes: list[Change] = []
    owner = _related(pkg, [(f"スライド {no}", s) for no, s in enumerate(_slide_order(pkg), start=1)], "comment")
    for part in _parts(pkg, r"ppt/comments/[^/]*\.xml"):
        label = (owner.get(part, "") + " コメント").strip()
        if _root_is(pkg, part, "p", "cmLst"):
            for text in pkg.xml(part).getroot().iterfind("p:cm/p:text", NS):
                changes += _replace_value(pkg, part, text, rules, label)
        elif _root_is(pkg, part, "p188", "cmLst"):
            changes += _replace_paragraphs(pkg, part, _q("a", "p"), _q("a", "t"), rules, lambda p, n, lb=label: lb)
    if pkg.has("ppt/commentAuthors.xml"):
        for a in pkg.xml("ppt/commentAuthors.xml").getroot().iterfind("p:cmAuthor", NS):
            changes += _replace_value(pkg, "ppt/commentAuthors.xml", a, rules, "コメントの作成者", attr="name")
    if pkg.has("ppt/authors.xml"):
        for a in pkg.xml("ppt/authors.xml").getroot().iterfind("p188:author", NS):
            changes += _replace_value(pkg, "ppt/authors.xml", a, rules, "コメントの作成者", attr="name")
    return changes
