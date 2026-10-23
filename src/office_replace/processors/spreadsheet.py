"""Excel: セル・図形・コメント・ヘッダー／フッター・テーブルの列名・シート名."""

from __future__ import annotations

import re

from lxml import etree

from ..ooxml import NS, BrokenOutputError, Change, Package, set_text
from ..rules import Rule
from ..sheetname import check_sheet_name
from ..textruns import replace_all
from .common import _apply_simple, _parts, _q, _related, _replace_paragraphs, _replace_value, _root_is, _shape_name, _short
from .formula import _formula_elements, _formula_literals, _table_parts, rename_sheet_refs, rename_table_columns


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


# --- Excel のヘッダー・フッター --------------------------------------------------------------------

_HF = {
    "oddHeader": "ヘッダー",
    "oddFooter": "フッター",
    "evenHeader": "偶数ページのヘッダー",
    "evenFooter": "偶数ページのフッター",
    "firstHeader": "先頭ページのヘッダー",
    "firstFooter": "先頭ページのフッター",
}
HEADER_FOOTER_TAGS = frozenset(_HF)


# 書式コード（& の後ろ）. && は文字の & なので、ここには入れない（hf_pieces で文字として扱う）
_HF_CODE = re.compile(
    r"&(?:"
    r'"[^"]*"?'  # &"フォント名,スタイル"
    r"|K(?:[0-9A-Fa-f]{6}|\d{2}[ST+\-]\d{3})"  # &K 文字の色（RGB / テーマの色）
    r"|[PN][+\-]\d+"  # &P+1 など（ページ番号の足し引き）
    r"|\d{1,3}"  # &12 文字の大きさ
    r"|[LCRPNDTZFAGBIUESXYOH]"  # 位置・ページ番号・日付・ファイル名・シート名・図・太字などの切り替え
    r"|.?"  # 知らない書式（壊さずにそのまま残す）
    r")",
    re.S | re.I,
)


def hf_pieces(value: str) -> tuple[list[str], set[int]]:
    """ヘッダー・フッターの文字列 → (文字と書式コードの並び, 書式コードの番号). 文字の && は & にする."""
    pieces: list[str] = []
    fixed: set[int] = set()
    buf: list[str] = []
    i = 0
    while i < len(value):
        if value.startswith("&&", i):
            buf.append("&")
            i += 2
        elif value[i] == "&":
            m = _HF_CODE.match(value, i)  # 末尾の .? があるので必ず一致する
            code = m.group() if m else value[i]
            if buf:
                pieces.append("".join(buf))
                buf = []
            fixed.add(len(pieces))
            pieces.append(code)
            i += len(code)
        else:
            buf.append(value[i])
            i += 1
    if buf:
        pieces.append("".join(buf))
    return pieces, fixed


def _replace_header_footer(pkg: Package, part: str, el: etree._Element, rules: list[Rule], location: str) -> list[Change]:
    """書式コード（&L・&P・&"ＭＳ ゴシック,太字"・&12 など）は残し、間の文字だけを置換する.

    一致は書式コードをまたがない（* も）. 置換後の文字の & は && にする（&P などの書式として読まれないように）.
    """
    old = el.text
    if not old:
        return []
    pieces, fixed = hf_pieces(old)
    new_pieces, count = replace_all(pieces, rules, fixed)
    if not count:
        return []
    new = "".join(t if k in fixed else t.replace("&", "&&") for k, t in enumerate(new_pieces))
    set_text(el, new)
    pkg.touch(part)
    return [Change(pkg.path.name, "text", location, _short(old), _short(new), count)]


def _excel_header_footer(pkg: Package, sheets, rules: list[Rule]) -> list[Change]:
    """印刷のヘッダー・フッター（&L&C&R・&P などの書式コードはそのまま、文字だけを置換する）."""
    changes: list[Change] = []
    for name, part, _ in sheets:
        if not pkg.has(part):
            continue
        for tag, label in _HF.items():
            for el in pkg.xml(part).getroot().iterfind(f"s:headerFooter/s:{tag}", NS):
                changes += _replace_header_footer(pkg, part, el, rules, f"{name} {label}")
    return changes


# --- Excel のテーブル（ListObject） ----------------------------------------------------------------


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
        # 置換後の列名を先に求め、重複する（Excel が修復を求める壊れたファイルになる）なら、このファイルは出力しない
        planned: list[tuple[etree._Element, str, str, int]] = []
        for col in cols:
            old = col.get("name", "")
            new, count = _apply_simple(rules, old)
            if count and new != old:
                planned.append((col, old, new, count))
        if planned:
            new_names = {id(c): n for c, _, n, _ in planned}
            seen: set[str] = set()
            for col in cols:
                name = new_names.get(id(col), col.get("name") or "")
                if name.casefold() in seen:
                    raise BrokenOutputError(
                        f"テーブル「{tname}」の列名が置換で重複します（{name}）. "
                        "Excel で開くと修復が必要なファイルになるため、このファイルは出力しません。置換表を見直してください"
                    )
                seen.add(name.casefold())
        mapping: dict[str, str] = {}
        for col, old, new, count in planned:
            col.set("name", new)
            mapping[old.casefold()] = new
            pkg.touch(part)
            changes.append(Change(pkg.path.name, "text", f"{sheet} テーブル「{tname}」の列名", _short(old), _short(new), count))
        for col in cols:
            if col.get("totalsRowLabel"):
                changes += _replace_value(pkg, part, col, rules, f"{sheet} テーブル「{tname}」の集計行", attr="totalsRowLabel")
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
    # 数式・名前の定義・グラフの参照を新しいシート名に合わせる（'旧名'! と 旧名!、3D 参照の両端. 外部のブックの参照は変えない）
    fixed = 0
    for part, el, attr in _formula_elements(pkg, sheets):
        value = el.get(attr) if attr else el.text
        if not value or "!" not in value:
            continue
        new_value, n = rename_sheet_refs(value, renames)
        if n:
            if attr:
                el.set(attr, new_value)
            else:
                el.text = new_value
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
