"""ファイルの種類ごとの置換（Excel / Word / PowerPoint）."""

from __future__ import annotations

import posixpath
import re
from collections.abc import Callable

from lxml import etree

from .ooxml import NS, Change, Package, paragraphs, set_text
from .rules import Rule
from .sheetname import check_sheet_name
from .textruns import replace_all

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
    for n, (p, ts) in enumerate(paragraphs(root, para_tag, text_tag), start=1):
        pieces = [t.text or "" for t in ts]
        new, count = replace_all(pieces, rules)
        if count:
            for t, v in zip(ts, new, strict=True):
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


# --- Excel --------------------------------------------------------------------------------------


def _sheets(pkg: Package) -> list[tuple[str, str, etree._Element]]:
    """(シート名, 部品のパス, workbook.xml の sheet 要素)."""
    wb = pkg.xml("xl/workbook.xml").getroot()
    rels = pkg.rels("xl/workbook.xml")
    out = []
    for s in wb.iterfind("s:sheets/s:sheet", NS):
        rid = s.get(_q("r", "id"))
        out.append((s.get("name", ""), rels.get(rid, ""), s))
    return out


def excel(pkg: Package, text_rules: list[Rule], sheet_rules: list[Rule]) -> tuple[list[Change], list[str]]:
    changes: list[Change] = []
    errors: list[str] = []
    sheets = _sheets(pkg)
    # 共有文字列がどのセルで使われているか（レポートの場所用）
    used: dict[int, list[str]] = {}
    for name, part, _ in sheets:
        if not pkg.has(part):
            continue
        for c in pkg.xml(part).getroot().iterfind(".//s:sheetData/s:row/s:c[@t='s']", NS):
            v = c.find("s:v", NS)
            if v is not None and (v.text or "").isdigit():
                used.setdefault(int(v.text), []).append(f"{name}!{c.get('r', '')}")

    if text_rules:
        if pkg.has("xl/sharedStrings.xml"):
            root = pkg.xml("xl/sharedStrings.xml").getroot()
            for idx, si in enumerate(root.iterfind("s:si", NS)):
                ts = [t for t in si.iter(_q("s", "t")) if not any(a.tag == _q("s", "rPh") for a in t.iterancestors())]
                pieces = [t.text or "" for t in ts]
                new, count = replace_all(pieces, text_rules)
                if count:
                    for t, v in zip(ts, new, strict=True):
                        set_text(t, v)
                    pkg.touch("xl/sharedStrings.xml")
                    cells = used.get(idx, [])
                    loc = (
                        "、".join(cells[:3]) + (f" ほか {len(cells) - 3} セル" if len(cells) > 3 else "") if cells else "（未使用の文字列）"
                    )
                    changes.append(
                        Change(pkg.path.name, "text", loc, _short("".join(pieces)), _short("".join(new)), count * max(1, len(cells)))
                    )
        for name, part, _ in sheets:
            if not pkg.has(part):
                continue
            root = pkg.xml(part).getroot()
            for c in root.iterfind(".//s:sheetData/s:row/s:c[@t='inlineStr']", NS):
                ts = list(c.iter(_q("s", "t")))
                pieces = [t.text or "" for t in ts]
                new, count = replace_all(pieces, text_rules)
                if count:
                    for t, v in zip(ts, new, strict=True):
                        set_text(t, v)
                    pkg.touch(part)
                    changes.append(
                        Change(pkg.path.name, "text", f"{name}!{c.get('r', '')}", _short("".join(pieces)), _short("".join(new)), count)
                    )
            for target in pkg.rels(part).values():
                if "/drawings/" in target and pkg.has(target):
                    changes += _replace_paragraphs(
                        pkg, target, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, nm=name: f"{nm} 図形「{_shape_name(p)}」"
                    )
                    for t2 in pkg.rels(target).values():
                        if "/charts/" in t2 and pkg.has(t2):
                            changes += _replace_paragraphs(
                                pkg, t2, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, nm=name: f"{nm} グラフ"
                            )
    if sheet_rules:
        c2, e2 = _rename_sheets(pkg, sheets, sheet_rules)
        changes += c2
        errors += e2
    return changes, errors


def _apply_simple(rules: list[Rule], text: str) -> tuple[str, int]:
    total = 0
    for r in rules:
        text, n = r.apply(text)
        total += n
    return text, total


def _ref(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_぀-ヿ㐀-鿿][\w.぀-ヿ㐀-鿿]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def _rename_sheets(pkg: Package, sheets, rules: list[Rule]) -> tuple[list[Change], list[str]]:
    changes, errors = [], []
    names = [n for n, _, _ in sheets]
    renames: dict[str, str] = {}
    for i, (old, _, el) in enumerate(sheets):
        new, count = _apply_simple(rules, old)
        if not count or new == old:
            continue
        others = [n for j, n in enumerate(names) if j != i]
        problem = check_sheet_name(new, others)
        if problem:
            errors.append(f"{pkg.path.name}: シート名「{old}」→「{new}」は変更しません（{problem}）")
            continue
        el.set("name", new)
        names[i] = new
        renames[old] = new
        pkg.touch("xl/workbook.xml")
        changes.append(Change(pkg.path.name, "sheet", "シート名", old, new, 1))
    if not renames:
        return changes, errors
    # 数式・名前の定義・グラフの参照を新しいシート名に合わせる（'旧名'! と 旧名! の両方。新しい名前は引用符で囲む）
    mapping: dict[str, str] = {}
    for o, n in renames.items():
        new_ref = "'" + n.replace("'", "''") + "'!"
        mapping["'" + o.replace("'", "''") + "'!"] = new_ref
        if _ref(o) == o:
            mapping[o + "!"] = new_ref
    pattern = re.compile(r"(?<![\w'.])(?:" + "|".join(re.escape(k) for k in sorted(mapping, key=len, reverse=True)) + ")")

    def fix(text: str) -> str:
        return pattern.sub(lambda m: mapping[m.group(0)], text)

    targets = [("xl/workbook.xml", ".//s:definedNames/s:definedName")]
    targets += [(part, ".//s:c/s:f") for _, part, _ in sheets if pkg.has(part)]
    targets += [(n, ".//c:f") for n in pkg.names if n.startswith("xl/charts/chart") and n.endswith(".xml")]
    fixed = 0
    for part, xpath in targets:
        for el in pkg.xml(part).getroot().iterfind(xpath, NS):
            if el.text and pattern.search(el.text):
                el.text = fix(el.text)
                pkg.touch(part)
                fixed += 1
    if fixed:
        changes.append(Change(pkg.path.name, "sheet", "数式・名前の定義・グラフの参照", "", f"{fixed} か所を新しいシート名に更新", fixed))
    return changes, errors


# --- Word ---------------------------------------------------------------------------------------

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
    return changes


def _in_textbox(p: etree._Element) -> bool:
    return any(a.tag == _q("w", "txbxContent") for a in p.iterancestors())


# --- PowerPoint ---------------------------------------------------------------------------------


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
    if include_masters:
        for part in sorted(n for n in pkg.names if re.fullmatch(r"ppt/(slideMasters/slideMaster|slideLayouts/slideLayout)\d+\.xml", n)):
            label = posixpath.basename(part).removesuffix(".xml")
            changes += _replace_paragraphs(
                pkg, part, _q("a", "p"), _q("a", "t"), text_rules, lambda p, n, lb=label: f"{lb} 図形「{_shape_name(p)}」"
            )
    return changes
