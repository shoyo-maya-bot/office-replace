"""各形式の置換で共通の部品（段落・値の置換、部品の探し方）."""

from __future__ import annotations

import re
from collections.abc import Callable

from lxml import etree

from ..ooxml import BREAKS, NS, Change, Package, is_break, paragraphs, set_text
from ..rules import Rule
from ..textruns import replace_all

Q = "{%s}%s"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _q(prefix: str, tag: str) -> str:
    return Q % (NS[prefix], tag)


def _short(s: str, n: int = 60) -> str:
    s = s.replace("\n", " ")
    return s if len(s) <= n else s[: n - 1] + "…"


def _replace_paragraphs(
    pkg: Package, part: str, para_tag: str, text_tag: str, rules: list[Rule], where: Callable[[etree._Element, int], str]
) -> list[Change]:
    changes = []
    root = pkg.xml(part).getroot()
    for n, (p, ts) in enumerate(paragraphs(root, para_tag, text_tag, breaks=True), start=1):
        # 改行・タブは "\n"・"\t" として段落の文字列に入れ、書き換えない（一致もまたがない）
        fixed = {k for k, t in enumerate(ts) if is_break(t)}
        pieces = [BREAKS[t.tag] if k in fixed else (t.text or "") for k, t in enumerate(ts)]
        new, count = replace_all(pieces, rules, fixed)
        if count:
            for k, (t, v) in enumerate(zip(ts, new, strict=True)):
                if k not in fixed:
                    set_text(t, v)
            pkg.touch(part)
            if not any(a.tag == MC_FALLBACK for a in p.iterancestors()):  # 互換用の複製は数えない
                changes.append(Change(pkg.path.name, "text", where(p, n), _short("".join(pieces)), _short("".join(new)), count))
    return changes


def _shape_name(el: etree._Element) -> str:
    for anc in el.iterancestors():
        for tag in ("xdr:cNvPr", "p:cNvPr"):
            prefix, local = tag.split(":")
            nv = anc.find(f".//{_q(prefix, local)}")
            if nv is not None and anc.tag.split("}")[1] in ("sp", "graphicFrame", "grpSp", "cxnSp"):
                return nv.get("name", "")
    return ""


def _replace_value(pkg: Package, part: str, el: etree._Element, rules: list[Rule], location: str, attr: str | None = None) -> list[Change]:
    """要素の文字（attr を指定したときは属性の値）を 1 つの文字列として置換する."""
    old = el.get(attr) if attr else el.text
    if not old:
        return []
    (new,), count = replace_all([old], rules)
    if not count:
        return []
    if attr:
        el.set(attr, new)
    else:
        set_text(el, new)
    pkg.touch(part)
    return [Change(pkg.path.name, "text", location, _short(old), _short(new), count)]


def _parts(pkg: Package, pattern: str) -> list[str]:
    return sorted(n for n in pkg.names if re.fullmatch(pattern, n))


def _root_is(pkg: Package, part: str, prefix: str, tag: str) -> bool:
    try:
        return pkg.xml(part).getroot().tag == _q(prefix, tag)
    except etree.XMLSyntaxError:
        return False


def _related(pkg: Package, sources: list[tuple[str, str]], word: str) -> dict[str, str]:
    """（ラベル, 部品）の一覧から、名前に word を含む関係先の部品 → ラベル."""
    out: dict[str, str] = {}
    for label, part in sources:
        for target in pkg.rels(part).values():
            if word in target.lower():
                out.setdefault(target, label)
    return out


def _apply_simple(rules: list[Rule], text: str) -> tuple[str, int]:
    total = 0
    for r in rules:
        text, n = r.apply(text)
        total += n
    return text, total
