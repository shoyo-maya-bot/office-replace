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


def excel(
    pkg: Package, text_rules: list[Rule], sheet_rules: list[Rule], formula_rules: list[Rule] | None = None
) -> tuple[list[Change], list[str]]:
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
        changes += _excel_header_footer(pkg, sheets, text_rules)
        c2, e2 = _excel_tables(pkg, sheets, text_rules)
        changes += c2
        errors += e2
    if formula_rules:
        changes += _formula_literals(pkg, sheets, formula_rules)
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


# 数式を持つ場所. 入力規則（dataValidation の formula1/2）・条件付き書式（cfRule の formula）は、
# Excel 2010 以降の拡張（extLst の xm:f）にもある
SHEET_FORMULA_XPATHS = [
    ".//s:c/s:f",
    ".//s:dataValidations/s:dataValidation/s:formula1",
    ".//s:dataValidations/s:dataValidation/s:formula2",
    ".//s:conditionalFormatting/s:cfRule/s:formula",
    ".//xm:f",
]


def _formula_elements(pkg: Package, sheets, *, references: bool = True) -> list[tuple[str, etree._Element, str | None]]:
    """数式を持つ要素の一覧 (部品, 要素, 属性名 or None). references: 参照だけを持つ場所（名前の定義・グラフ・リンク）も含める."""
    targets: list[tuple[str, str, str | None]] = []
    if references:
        targets.append(("xl/workbook.xml", ".//s:definedNames/s:definedName", None))
    for _, part, _ in sheets:
        if pkg.has(part):
            targets += [(part, x, None) for x in SHEET_FORMULA_XPATHS]
            targets += [(part, ".//s:conditionalFormatting/s:cfRule/s:cfvo[@type='formula']", "val")]
            if references:
                targets += [(part, ".//s:hyperlinks/s:hyperlink", "location")]  # ブックの中へのリンク
    if references:
        targets += [(n, ".//c:f", None) for n in pkg.names if n.startswith("xl/charts/chart") and n.endswith(".xml")]
        for part in _table_parts(pkg, sheets):
            targets += [(part, "s:tableColumns/s:tableColumn/s:calculatedColumnFormula", None)]
            targets += [(part, "s:tableColumns/s:tableColumn/s:totalsRowFormula", None)]
    out = []
    for part, xpath, attr in targets:
        for el in pkg.xml(part).getroot().iterfind(xpath, NS):
            out.append((part, el, attr))
    return out


# --- Excel のヘッダー・フッター --------------------------------------------------------------------

_HF = {
    "oddHeader": "ヘッダー",
    "oddFooter": "フッター",
    "evenHeader": "偶数ページのヘッダー",
    "evenFooter": "偶数ページのフッター",
    "firstHeader": "先頭ページのヘッダー",
    "firstFooter": "先頭ページのフッター",
}


def _excel_header_footer(pkg: Package, sheets, rules: list[Rule]) -> list[Change]:
    """印刷のヘッダー・フッター（&L&C&R などの書式コードはそのまま、文字だけを置換する）."""
    changes: list[Change] = []
    for name, part, _ in sheets:
        if not pkg.has(part):
            continue
        for tag, label in _HF.items():
            for el in pkg.xml(part).getroot().iterfind(f"s:headerFooter/s:{tag}", NS):
                changes += _replace_value(pkg, part, el, rules, f"{name} {label}")
    return changes


# --- Excel のテーブル（ListObject） ----------------------------------------------------------------


def _table_parts(pkg: Package, sheets) -> list[str]:
    out = []
    for _, part, _ in sheets:
        if not pkg.has(part):
            continue
        for target in pkg.rels(part).values():
            if "/tables/" in target and pkg.has(target) and _root_is(pkg, target, "s", "table"):
                out.append(target)
    return out


def _excel_tables(pkg: Package, sheets, rules: list[Rule]) -> tuple[list[Change], list[str]]:
    """テーブルの列名（tableColumn@name）を、同じルールで置換した見出しセルに合わせる.

    列名と見出しセルが食い違うと、Excel で開いたときに「修復」になる。列名を変えたら、
    構造化参照（テーブル1[旧列名]・テーブル1[[#This Row],[旧列名]] など）を使う数式・名前の定義も直す。
    """
    changes: list[Change] = []
    errors: list[str] = []
    renames: dict[str, dict[str, str]] = {}  # テーブル名（casefold）→ {旧列名（casefold）: 新列名}
    owner = {t: name for name, part, _ in sheets if pkg.has(part) for t in pkg.rels(part).values()}
    for part in _table_parts(pkg, sheets):
        root = pkg.xml(part).getroot()
        tname = root.get("displayName") or root.get("name", "")
        sheet = owner.get(part, "")
        cols = root.findall("s:tableColumns/s:tableColumn", NS)
        mapping: dict[str, str] = {}
        for col in cols:
            old = col.get("name", "")
            new, count = _apply_simple(rules, old)
            if count and new != old:
                col.set("name", new)
                mapping[old.casefold()] = new
                pkg.touch(part)
                changes.append(Change(pkg.path.name, "text", f"{sheet} テーブル「{tname}」の列名", _short(old), _short(new), count))
            label = col.get("totalsRowLabel")
            if label:
                changes += _replace_value(pkg, part, col, rules, f"{sheet} テーブル「{tname}」の集計行", attr="totalsRowLabel")
        seen: dict[str, str] = {}
        for col in cols:
            key = (col.get("name") or "").casefold()
            if key in seen:
                errors.append(
                    f"{pkg.path.name}: テーブル「{tname}」の列名が置換で重複します（{col.get('name')}）. "
                    "Excel で開くと修復が必要になります。置換表を見直してください"
                )
            seen[key] = col.get("name", "")
        if mapping:
            renames[tname.casefold()] = mapping
    if renames:
        fixed = 0
        for part, el, attr in _formula_elements(pkg, sheets):
            value = el.get(attr) if attr else el.text
            if not value or "[" not in value:
                continue
            new = rename_table_columns(value, renames)
            if new != value:
                if attr:
                    el.set(attr, new)
                else:
                    el.text = new
                pkg.touch(part)
                fixed += 1
        if fixed:
            changes.append(
                Change(
                    pkg.path.name,
                    "text",
                    "数式・名前の定義のテーブルの列の参照（構造化参照）",
                    "",
                    f"{fixed} か所を新しい列名に更新",
                    fixed,
                )
            )
    return changes, errors


_SPECIAL = "[]#'"


def _unescape_col(raw: str) -> str:
    out, i = [], 0
    while i < len(raw):
        if raw[i] == "'" and i + 1 < len(raw):
            out.append(raw[i + 1])
            i += 2
            continue
        out.append(raw[i])
        i += 1
    return "".join(out)


def _escape_col(name: str) -> str:
    return "".join("'" + ch if ch in _SPECIAL else ch for ch in name)


def _skip_quoted(f: str, i: int, quote: str) -> int:
    """f[i] の引用符で始まる文字列（"..." や '...'。2 つ重ねて 1 文字）の次の位置."""
    j = i + 1
    while j < len(f):
        if f[j] == quote:
            if j + 1 < len(f) and f[j + 1] == quote:
                j += 2
                continue
            return j + 1
        j += 1
    return len(f)


def _bracket_end(f: str, i: int) -> int | None:
    depth, j = 0, i
    while j < len(f):
        c = f[j]
        if c == "'":
            j += 2
            continue
        if c == "[":
            depth += 1
        elif c == "]":
            depth -= 1
            if depth == 0:
                return j + 1
        j += 1
    return None


def _rewrite_spec(seg: str, mapping: dict[str, str]) -> str:
    """[列名] / [[#Headers],[列名]] / [[列1]:[列2]] の、いちばん内側の [ ] の列名を置き換える."""
    edits: list[tuple[int, int, str]] = []
    stack: list[list] = []  # [開き括弧の位置, 子があるか]
    j = 0
    while j < len(seg):
        c = seg[j]
        if c == "'":
            j += 2
            continue
        if c == "[":
            if stack:
                stack[-1][1] = True
            stack.append([j, False])
        elif c == "]" and stack:
            start, has_child = stack.pop()
            if not has_child:
                raw = seg[start + 1 : j]
                at = "@" if raw.startswith("@") else ""
                body = raw[len(at) :]
                if not body.startswith("#"):
                    new = mapping.get(_unescape_col(body).strip().casefold())
                    if new is not None:
                        edits.append((start + 1, j, at + _escape_col(new)))
        j += 1
    for s, e, text in reversed(edits):
        seg = seg[:s] + text + seg[e:]
    return seg


def rename_table_columns(formula: str, tables: dict[str, dict[str, str]]) -> str:
    """数式の構造化参照（テーブル名[...]）の列名を置き換える. 文字列（"..."）とシート名（'...'）の中は変えない.

    tables: テーブル名（casefold）→ {旧列名（casefold）: 新列名}. ファイルの中では [@列] も テーブル名[[#This Row],[列]] の形で保存される。
    """
    out: list[str] = []
    i = 0
    while i < len(formula):
        ch = formula[i]
        if ch in "\"'":
            j = _skip_quoted(formula, i, ch)
            out.append(formula[i:j])
            i = j
            continue
        if ch == "[":
            end = _bracket_end(formula, i)
            if end is None:
                out.append(formula[i:])
                break
            prefix = "".join(out)
            m = re.search(r"[\w.\\]+$", prefix)
            mapping = tables.get(m.group(0).casefold()) if m else None
            seg = formula[i:end]
            out.append(_rewrite_spec(seg, mapping) if mapping else seg)
            i = end
            continue
        out.append(ch)
        i += 1
    return "".join(out)


# --- Excel の数式の中の文字列（対象「数式」を指定したときだけ） ---------------------------------------

_LITERAL = re.compile(r'"((?:[^"]|"")*)"')


def replace_literals(formula: str, rules: list[Rule]) -> tuple[str, int]:
    """数式の中の文字列（"..."）だけを置換する. セル参照・関数名・シート名は変えない."""
    total = 0

    def sub(m: re.Match[str]) -> str:
        nonlocal total
        new, n = _apply_simple(rules, m.group(1).replace('""', '"'))
        total += n
        return '"' + new.replace('"', '""') + '"'

    return _LITERAL.sub(sub, formula), total


def _drop_cached(c: etree._Element) -> None:
    """数式セルの計算結果（v）を消す. 置換前の文字列の計算結果が残らないよう、開いたときに計算し直させる."""
    for v in c.findall("s:v", NS):
        c.remove(v)
    if "t" in c.attrib:
        del c.attrib["t"]


def _full_calc_on_load(pkg: Package) -> None:
    root = pkg.xml("xl/workbook.xml").getroot()
    calc = root.find("s:calcPr", NS)
    if calc is None:
        calc = etree.Element(_q("s", "calcPr"))
        anchor = None
        for tag in ("sheets", "functionGroups", "externalReferences", "definedNames"):
            el = root.find(f"s:{tag}", NS)
            if el is not None:
                anchor = el
        if anchor is None:
            root.append(calc)
        else:
            anchor.addnext(calc)
    calc.set("fullCalcOnLoad", "1")
    pkg.touch("xl/workbook.xml")


def _formula_literals(pkg: Package, sheets, rules: list[Rule]) -> list[Change]:
    changes: list[Change] = []
    for name, part, _ in sheets:
        if not pkg.has(part):
            continue
        root = pkg.xml(part).getroot()
        shared: set[str] = set()
        for c in root.iterfind(".//s:sheetData/s:row/s:c", NS):
            f = c.find("s:f", NS)
            if f is None or not f.text:
                continue
            new, count = replace_literals(f.text, rules)
            if count and new != f.text:
                changes.append(Change(pkg.path.name, "formula", f"{name}!{c.get('r', '')}", _short(f.text), _short(new), count))
                f.text = new
                _drop_cached(c)
                if f.get("t") == "shared" and f.get("si") is not None:
                    shared.add(f.get("si", ""))
                pkg.touch(part)
        if shared:  # 共有数式の子のセル（数式の文字は親だけが持つ）の計算結果も消す
            for c in root.iterfind(".//s:sheetData/s:row/s:c", NS):
                f = c.find("s:f", NS)
                if f is not None and f.get("t") == "shared" and f.get("si") in shared:
                    _drop_cached(c)
        for p2, el, attr in _formula_elements(pkg, [(name, part, None)], references=False):
            if el.tag == _q("s", "f") and el.getparent() is not None and el.getparent().tag == _q("s", "c"):
                continue  # セルの数式は上で済み
            value = el.get(attr) if attr else el.text
            if not value:
                continue
            new, count = replace_literals(value, rules)
            if count and new != value:
                if attr:
                    el.set(attr, new)
                else:
                    el.text = new
                pkg.touch(p2)
                changes.append(Change(pkg.path.name, "formula", f"{name} 入力規則・条件付き書式", _short(value), _short(new), count))
    if changes:
        _full_calc_on_load(pkg)
    return changes


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

    fixed = 0
    for part, el, attr in _formula_elements(pkg, sheets):
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
