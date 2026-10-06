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


# --- 文書のプロパティ・グラフのデータ（Excel / Word / PowerPoint 共通） ----------------------------------

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


def _drop_phonetic(container: etree._Element) -> None:
    """ふりがな（rPh）を外す. 置換した文字列に元の読みが残ると、固有名が読めてしまうため."""
    for rph in container.findall("s:rPh", NS):
        container.remove(rph)


def _not_phonetic(t: etree._Element) -> bool:
    return not any(a.tag == _q("s", "rPh") for a in t.iterancestors())


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
                ts = [t for t in si.iter(_q("s", "t")) if _not_phonetic(t)]
                pieces = [t.text or "" for t in ts]
                new, count = replace_all(pieces, text_rules)
                if count:
                    for t, v in zip(ts, new, strict=True):
                        set_text(t, v)
                    _drop_phonetic(si)
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
                ts = [t for t in c.iter(_q("s", "t")) if _not_phonetic(t)]
                pieces = [t.text or "" for t in ts]
                new, count = replace_all(pieces, text_rules)
                if count:
                    for t, v in zip(ts, new, strict=True):
                        set_text(t, v)
                    for is_ in c.iterfind("s:is", NS):
                        _drop_phonetic(is_)
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
        changes += _excel_comments(pkg, sheets, text_rules)
    if sheet_rules:
        c2, e2 = _rename_sheets(pkg, sheets, sheet_rules)
        changes += c2
        errors += e2
    return changes, errors


def _excel_comments(pkg: Package, sheets, rules: list[Rule]) -> list[Change]:
    """セルのコメント（従来のメモ・スレッド形式）とその作成者."""
    changes: list[Change] = []
    owner = _related(pkg, [(name, part) for name, part, _ in sheets], "comment")
    for part in _parts(pkg, r"(?i)xl/.*comment[^/]*\.xml"):
        sheet = owner.get(part, "")
        if _root_is(pkg, part, "s", "comments"):
            root = pkg.xml(part).getroot()
            for c in root.iterfind("s:commentList/s:comment", NS):
                ts = [t for t in c.iter(_q("s", "t")) if _not_phonetic(t)]
                pieces = [t.text or "" for t in ts]
                new, count = replace_all(pieces, rules)
                if count:
                    for t, v in zip(ts, new, strict=True):
                        set_text(t, v)
                    for text in c.iterfind("s:text", NS):
                        _drop_phonetic(text)
                    pkg.touch(part)
                    loc = f"{sheet}!{c.get('ref', '')} コメント"
                    changes.append(Change(pkg.path.name, "text", loc, _short("".join(pieces)), _short("".join(new)), count))
            for a in root.iterfind("s:authors/s:author", NS):
                changes += _replace_value(pkg, part, a, rules, f"{sheet} コメントの作成者")
        elif _root_is(pkg, part, "tc", "ThreadedComments"):
            for c in pkg.xml(part).getroot().iterfind("tc:threadedComment", NS):
                for text in c.iterfind("tc:text", NS):
                    changes += _replace_value(pkg, part, text, rules, f"{sheet}!{c.get('ref', '')} コメント（スレッド）")
    for part in _parts(pkg, r"xl/persons/[^/]*\.xml"):
        for person in pkg.xml(part).getroot().iterfind("tc:person", NS):
            changes += _replace_value(pkg, part, person, rules, "コメントの作成者（スレッド）", attr="displayName")
    return changes


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

    # 入力規則（dataValidation の formula1/2）・条件付き書式（cfRule の formula）は、Excel 2010 以降の拡張（extLst の xm:f）にもある
    sheet_xpaths = [
        ".//s:c/s:f",
        ".//s:dataValidations/s:dataValidation/s:formula1",
        ".//s:dataValidations/s:dataValidation/s:formula2",
        ".//s:conditionalFormatting/s:cfRule/s:formula",
        ".//xm:f",
    ]
    targets = [("xl/workbook.xml", ".//s:definedNames/s:definedName", None)]
    for _, part, _ in sheets:
        if pkg.has(part):
            targets += [(part, x, None) for x in sheet_xpaths]
            targets += [(part, ".//s:conditionalFormatting/s:cfRule/s:cfvo[@type='formula']", "val")]
            targets += [(part, ".//s:hyperlinks/s:hyperlink", "location")]  # ブックの中へのリンク
    targets += [(n, ".//c:f", None) for n in pkg.names if n.startswith("xl/charts/chart") and n.endswith(".xml")]
    fixed = 0
    for part, xpath, attr in targets:
        for el in pkg.xml(part).getroot().iterfind(xpath, NS):
            value = el.get(attr) if attr else el.text
            if value and pattern.search(value):
                if attr:
                    el.set(attr, fix(value))
                else:
                    el.text = fix(value)
                pkg.touch(part)
                fixed += 1
    # 名前そのものを持つところ（ピボットテーブルの元データ・文書のプロパティのシート名の一覧）
    exact = [(n, ".//s:cacheSource/s:worksheetSource", "sheet") for n in _parts(pkg, r"xl/pivotCache/pivotCacheDefinition[^/]*\.xml")]
    if pkg.has("docProps/app.xml"):
        exact.append(("docProps/app.xml", "ep:TitlesOfParts/vt:vector/vt:lpstr", None))
    for part, xpath, attr in exact:
        for el in pkg.xml(part).getroot().iterfind(xpath, NS):
            value = el.get(attr) if attr else el.text
            if value in renames:
                if attr:
                    el.set(attr, renames[value])
                else:
                    el.text = renames[value]
                pkg.touch(part)
                fixed += 1
    if fixed:
        changes.append(
            Change(
                pkg.path.name,
                "sheet",
                "数式・名前の定義・入力規則・条件付き書式・グラフの参照",
                "",
                f"{fixed} か所を新しいシート名に更新",
                fixed,
            )
        )
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
