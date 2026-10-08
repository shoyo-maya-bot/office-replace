"""Excel の数式: 数式を持つ場所、構造化参照の列名の書き換え、数式の中の文字列の置換."""

from __future__ import annotations

import re

from lxml import etree

from ..ooxml import NS, Change, Package
from ..rules import Rule
from .common import _apply_simple, _q, _root_is, _short

_UNQUOTED_REF = re.compile(r"(\[[^\]'!]*\])?([\w.]+)(?::([\w.]+))?!")


def _quote_sheets(names: list[str]) -> str:
    return "'" + ":".join(names).replace("'", "''") + "'!"


def rename_sheet_refs(formula: str, renames: dict[str, str]) -> tuple[str, int]:
    """数式のシートの参照（シート名!）を新しいシート名にする → (新しい数式, 書き換えた数).

    - 'シート名'! と シート名! の両方. 書き換えた参照は引用符で囲む（'新名'!）
    - 3D 参照（集計:別!A1 / '集計:別'!A1）は、両端のシート名をそれぞれ直し、全体を 1 つの引用符で囲む（'集計:新名'!A1）
    - 外部のブックの参照（[1]集計!A1 / '[1]集計'!A1）は書き換えない（別のブックのシートなので）
    - 文字列（"..."）の中は書き換えない
    """
    out: list[str] = []
    n = 0
    i = 0
    while i < len(formula):
        ch = formula[i]
        # 名前の途中（直前が文字・数字・_ . ' ]）からは参照を始めない
        boundary = i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "_.']")
        if ch == '"':  # 文字列（"..."）の中は変えない（Excel もシート名の変更で INDIRECT の文字列は変えない）
            j = _skip_quoted(formula, i, '"')
            out.append(formula[i:j])
            i = j
            continue
        if ch == "'" and boundary:
            j = _skip_quoted(formula, i, "'")
            if formula[j : j + 1] == "!" and formula[j - 1 : j] == "'":
                body = formula[i + 1 : j - 1].replace("''", "'")
                if not body.startswith("["):  # 外部のブック
                    names = body.split(":")  # シート名には : を使えない
                    new = [renames.get(x, x) for x in names]
                    if new != names:
                        out.append(_quote_sheets(new))
                        n += 1
                        i = j + 1
                        continue
                out.append(formula[i : j + 1])
                i = j + 1
                continue
        elif boundary:
            m = _UNQUOTED_REF.match(formula, i)
            if m:
                if not m.group(1):  # [1]シート名! は外部のブック
                    names = [x for x in (m.group(2), m.group(3)) if x is not None]
                    new = [renames.get(x, x) for x in names]
                    if new != names:
                        out.append(_quote_sheets(new))
                        n += 1
                        i = m.end()
                        continue
                out.append(m.group(0))
                i = m.end()
                continue
        out.append(ch)
        i += 1
    return "".join(out), n


def sheet_ref_spans(formula: str) -> list[tuple[int, int]]:
    """数式の中のシートの参照（'シート名'! / シート名! / 3D 参照）の範囲 (始め, 終わり) の一覧. 外部のブックの参照・文字列の中は除く.

    rename_sheet_refs と同じ読み方. 残りの確認で、置換前がシート名のせいで数式に残っているかを見分けるのに使う.
    """
    spans: list[tuple[int, int]] = []
    i = 0
    while i < len(formula):
        ch = formula[i]
        boundary = i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "_.']")
        if ch == '"':
            i = _skip_quoted(formula, i, '"')
            continue
        if ch == "'" and boundary:
            j = _skip_quoted(formula, i, "'")
            if formula[j : j + 1] == "!" and formula[j - 1 : j] == "'":
                if not formula[i + 1 : j - 1].startswith("["):
                    spans.append((i, j + 1))
                i = j + 1
                continue
        elif boundary:
            m = _UNQUOTED_REF.match(formula, i)
            if m:
                if not m.group(1):
                    spans.append((m.start(), m.end()))
                i = m.end()
                continue
        i += 1
    return spans


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


def _table_parts(pkg: Package, sheets) -> list[str]:
    out = []
    for _, part, _ in sheets:
        if not pkg.has(part):
            continue
        for target in pkg.rels(part).values():
            if "/tables/" in target and pkg.has(target) and _root_is(pkg, target, "s", "table"):
                out.append(target)
    return out


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
