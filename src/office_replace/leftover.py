"""書き出したファイルに、置換前の文字列が残っていないかを調べる（固有名を伏せる用途の確認）.

ツールが置換しない場所（数式の中の文字列・変更履歴・グラフの元データの埋め込みブックなど）に残った文字列を、
注意（rule_id: replace.leftover）としてレポートに出す。残った理由（reason）も付ける。対象は「文字列」が対象の、正規表現でないルール。

調べないもの（文字列ではない値で、置換前に当たっても固有名ではないもの）:
- 文書のプロパティ（docProps/）の日付（作成日時・更新日時・印刷日時・ユーザー設定の日付）. 年度の置換（2025 → 2026）で毎回当たるため
- Excel の共有文字列の番号（t="s" のセルの v. 文字ではなく番号）
- 数値の構造の値を持つ要素の文字（図形のアンカーの行・列・オフセット、Word の図の位置、文書のプロパティの統計（ページ数・編集時間など）・版数）
- XML の構造用の属性（色・大きさ・座標・ID・番号・書式など）. 調べる属性は、利用者が付けた名前や画面に出る文字を持つものだけ
  （TEXT_ATTRS / ELEMENT_TEXT_ATTRS: シート名・名前の定義・テーブル名と列名・図形の名前と代替テキスト・作成者・
  入力規則のメッセージ・ヒント・外部へのリンク先など）. `--from 50` が色（FF5050）や座標（off@x="5050"）に当たらないように

Excel の数値のセル（数式の結果が数値のものも）の v は、値の全体が置換前に一致するときだけ数える（部分一致は数えない）.
「202」の置換で 2025 が残りにならないように。文字列のセル・文字列の数式の結果（t="str"）は部分一致も数える.
"""

from __future__ import annotations

import html
import io
import zipfile
from dataclasses import dataclass

from lxml import etree

from .ooxml import _PARSER, UnsafeZipError, ZipLimits, check_zip
from .processors.formula import sheet_ref_spans
from .processors.spreadsheet import HEADER_FOOTER_TAGS, hf_pieces
from .rules import Rule

RULE_ID = "replace.leftover"
XML_SUFFIXES = (".xml", ".rels", ".vml")
OFFICE_SUFFIXES = (".xlsx", ".xlsm", ".docx", ".docm", ".pptx", ".pptm")
MAX_EMBEDDED = 50 * 1024 * 1024  # 埋め込みのファイルはこの大きさまで調べる
MAX_DEPTH = 2
V_FORMULA = "v（数式の結果）"
V_VALUE = "v（数値など）"
SHEET_REF = "（シート名の参照）"  # 数式の中の、シートの参照（シート名!）に当たったもの. 場所の名前の後ろに付ける
# 数式を持つ要素（セル・グラフの f、入力規則・条件付き書式の formula、名前の定義）
FORMULA_TAGS = ("f", "formula", "formula1", "formula2", "definedName")
# docProps の日付の要素（dcterms:created / modified、cp:lastPrinted、ユーザー設定の日付 vt:filetime / vt:date）
DOCPROPS_DATES = {"created", "modified", "lastPrinted", "filetime", "date"}

# 調べる属性（ローカル名）. 利用者が付けた名前・画面に出る文字を持つもの. ほかの属性（色・座標・ID・書式など）は調べない
TEXT_ATTRS = {
    "name",  # シート名・名前の定義・テーブル名・列名・図形の名前・ピボットの項目名など
    "displayName",  # テーブル名
    "descr",  # 図形・画像の代替テキスト
    "title",  # 図形の代替テキストのタイトル・VML の題名
    "alt",  # VML の代替テキスト
    "tooltip",  # リンクのヒント
    "display",  # Excel のリンクの表示文字
    "location",  # Excel のリンク先（シート名!セル）
    "author",  # 変更履歴・コメントの作成者
    "initials",  # コメントの作成者の頭文字
    "userId",  # 変更履歴の作成者の ID（メールアドレスなど）
    "prompt",  # 入力規則のメッセージ
    "promptTitle",
    "error",
    "errorTitle",
    "totalsRowLabel",  # テーブルの集計行の文字
    "sheet",  # ピボットテーブルの元データのシート名
    "caption",  # ピボットの見出し
    "text",  # 条件付き書式の「次の文字列を含む」
    "string",  # VML の文字（textpath）
}
# 要素によって意味が変わる属性（val など）. 要素名と属性名の組で調べる
ELEMENT_TEXT_ATTRS = {
    ("alias", "val"),  # Word のコンテンツコントロールのタイトル
    ("tag", "val"),  # Word のコンテンツコントロールのタグ
    ("docVar", "val"),  # Word の文書変数
    ("listEntry", "displayText"),  # Word のドロップダウンの選択肢
    ("listEntry", "value"),
    ("filter", "val"),  # Excel のオートフィルターの条件
    ("customFilter", "val"),
    ("sheetName", "val"),  # Excel の外部のブックのシート名
    ("s", "v"),  # ピボットのキャッシュの文字列の項目
    ("item", "n"),  # ピボットの項目の名前
}

# 文字（テキストノード）が数値の構造の値で、調べない要素（ローカル名）
STRUCT_TEXT = {"col", "colOff", "row", "rowOff", "posOffset"}
# docProps/ の統計など（ページ数・編集時間・版数など. 文字ではない値）
DOCPROPS_STATS = {
    "TotalTime",
    "Pages",
    "Words",
    "Characters",
    "CharactersWithSpaces",
    "Lines",
    "Paragraphs",
    "Slides",
    "Notes",
    "HiddenSlides",
    "MMClips",
    "DocSecurity",
    "AppVersion",
    "revision",
    "i4",  # 見出しの組（HeadingPairs）の数など
}


def _text_attr(element: str, attr: str, el: etree._Element) -> bool:
    """属性を残りの確認で調べるか. 外部へのリンク（TargetMode="External"）のリンク先も調べる（ファイルのパス・URL）."""
    if attr in TEXT_ATTRS or (element, attr) in ELEMENT_TEXT_ATTRS:
        return True
    return element == "Relationship" and attr == "Target" and el.get("TargetMode") == "External"


@dataclass
class Leftover:
    rule_id: str
    part: str  # zip の中の部品（埋め込みは「外の部品!中の部品」）
    where: str  # 要素名・属性名（最大 3 つ）
    before: str  # ルールの置換前
    count: int
    reason: str = ""  # 残った理由と対処


FILE_NAME_PART = "（出力ファイル名）"
FILE_NAME_WHERE = "ファイル名・フォルダ名"


def name_rules(rules: list[Rule]) -> list[Rule]:
    """名前を調べるルール. 対象に「文字列」か「ファイル名」を含む、正規表現でないもの（置換後に置換前が含まれるものは除く）."""
    return [r for r in rules if not r.regex and {"text", "file"} & set(r.targets) and not r.compiled().search(r.after)]


def scan_name(rel_dest: str, rules: list[Rule]) -> list[Leftover]:
    """出力ファイルの名前（出力フォルダからの相対パス）に残った置換前の文字列. rules は name_rules で選んだもの."""
    out = []
    for r in rules:
        n = _count(r.compiled(), rel_dest)
        if n:
            out.append(
                Leftover(RULE_ID, f"{FILE_NAME_PART} {rel_dest}", FILE_NAME_WHERE, r.before, n, reason(FILE_NAME_PART, FILE_NAME_WHERE))
            )
    return out


# 残った理由（部品・要素名から判断する）. 上から順に当てはめる
def reason(part: str, where: str) -> str:
    el = where.split("、", 1)[0]
    tag = el.split("@", 1)[0]
    if part.startswith(FILE_NAME_PART):
        return (
            "出力ファイルの名前（またはフォルダ名）。ファイル名は対象に「ファイル名」（--target file）を指定すると置換できます。"
            "フォルダ名は置換しないので、出力フォルダで名前を変えてください"
        )
    if "!" in part:
        return "埋め込みファイル（グラフの元データのブックなど）の中。このツールは置換しません。元のアプリで開いて直してください"
    if part.startswith("xl/") and tag.endswith(SHEET_REF):
        return (
            "Excel の数式・名前の定義の中のシート名の参照（シート名!）。シート名に置換前が入っているためです。"
            "対象に「シート名」（--target sheet）を指定すると、シート名と参照をあわせて置換できます"
        )
    if part == "xl/workbook.xml" and el == "sheet@name":
        return "Excel のシート名。対象に「シート名」（--target sheet）を指定すると置換できます（数式の参照もあわせて直します）"
    if part == "docProps/app.xml" and tag == "lpstr":
        return "文書のプロパティのシート名の一覧。対象に「シート名」（--target sheet）を指定するとシート名とあわせて置換できます"
    if part.startswith("xl/") and tag in FORMULA_TAGS:
        return "Excel の数式の中の文字列。対象に「数式」（--target formula）を指定すると置換できます"
    if part.startswith("xl/worksheets/") and tag == V_FORMULA:
        return (
            "Excel の数式の計算結果（ファイルに保存されている値）。このツールは計算結果を書き換えません。"
            "数式の中の文字列が元なら --target formula で置換すると計算結果も消えます。"
            "置換したセルを参照している数式なら、Excel で開いて再計算して保存すると直ります"
        )
    if part.startswith("xl/worksheets/") and tag == V_VALUE:
        return "Excel の数値などのセルの値（文字列ではない値）。このツールは数値を置換しません。Excel で直してください"
    if part.startswith("xl/pivotCache/"):
        return "ピボットテーブルのキャッシュ。Excel で開いて「更新」すると元のデータに合わせて置き換わります"
    if part.startswith("xl/externalLinks/"):
        return "外部のブックへのリンク。Excel の「リンクの編集」で直してください"
    if tag in ("delText", "delInstrText") or tag in ("ins", "del", "moveFrom", "moveTo") or "@author" in el:
        return "変更履歴（削除した文字・作成者）やその他の作成者名。Word などで変更履歴を承諾・削除してください"
    if tag == "instrText":
        return "Word のフィールドコード。Word で開いてフィールドを直してください"
    if part.endswith(".vml") or tag.startswith("（"):
        return "古い形式の図形（VML）など、このツールが置換しない部品"
    return "このツールが置換しない場所。Office で開いて確認してください"


def scan_rules(rules: list[Rule]) -> list[Rule]:
    """調べるルール. 置換後に置換前が含まれるルール（集計 → 集計_2027 など）は、置換後の文字に当たってしまうため除く."""
    return [r for r in rules if not r.regex and "text" in r.targets and not r.compiled().search(r.after)]


def _local(tag) -> str:
    return tag.split("}", 1)[1] if isinstance(tag, str) and "}" in tag else str(tag)


class _Number(str):
    """Excel の数値のセルの値（文字列ではない値）. 残りの確認では値の全体の一致だけを数える."""


# セルの型（c@t）のうち、値が文字列のもの. 無し・n（数値）・b（真偽）・e（エラー）・d（日付）は文字列ではない
_TEXT_TYPES = {"str", "inlineStr"}


def _v_name(el: etree._Element) -> str | None:
    """Excel のセルの v の場所の名前. 共有文字列の番号（t="s"）は文字ではないので None（調べない）."""
    c = el.getparent()
    if c is None or _local(c.tag) != "c":
        return "v"
    if c.get("t") == "s":
        return None
    if any(_local(x.tag) == "f" for x in c if isinstance(x.tag, str)) or c.get("t") == "str":
        return V_FORMULA
    return V_VALUE


def _count(rx, text: str) -> int:
    """text の中の一致の数. 数値のセル（_Number）は値の全体が一致するときだけ 1（部分一致は数えない）."""
    if isinstance(text, _Number):
        return 1 if rx.fullmatch(text) else 0
    return sum(1 for m in rx.finditer(text) if m.end() > m.start())


def _split_sheet_refs(rx, part: str, where: str, text: str) -> list[tuple[str, int]]:
    """一致の数を、数式の中のシート名の参照に当たったもの（where + SHEET_REF）とそれ以外に分ける."""
    if not (part.startswith("xl/") and where in FORMULA_TAGS) or "!" not in text:
        return [(where, _count(rx, text))]
    spans = sheet_ref_spans(text)
    in_ref = other = 0
    for m in rx.finditer(text):
        if m.end() <= m.start():
            continue
        if any(a <= m.start() and m.end() <= b for a, b in spans):
            in_ref += 1
        else:
            other += 1
    return [(where + SHEET_REF, in_ref), (where, other)]


def _strings(data: bytes, tree: etree._ElementTree | None, part: str = "") -> list[tuple[str, str]]:
    """部品の中の（場所, 文字列）. 文字・属性の値ごとに調べる."""
    skip = DOCPROPS_DATES | DOCPROPS_STATS if part.startswith("docProps/") else STRUCT_TEXT
    sheet = part.startswith("xl/worksheets/")
    if tree is None:
        try:
            tree = etree.ElementTree(etree.fromstring(data, _PARSER))
        except etree.XMLSyntaxError:  # 古い VML など、整形式でないもの
            return [("（XML として読めない部品）", html.unescape(data.decode("utf-8", "replace")))]
    out: list[tuple[str, str]] = []
    for el in tree.getroot().iter():
        name = _local(el.tag) if isinstance(el.tag, str) else "comment"
        if name in skip:
            continue
        number = False
        if sheet and name == "v":
            name = _v_name(el)
            if name is None:
                continue
            number = el.getparent() is not None and el.getparent().get("t") not in _TEXT_TYPES
        if sheet and name in HEADER_FOOTER_TAGS and el.text:
            # ヘッダー・フッターは書式コード（&P・&D など）を除いた文字だけを調べる（置換と同じ区切り. && は & として）
            pieces, fixed = hf_pieces(el.text)
            out += [(name, t) for k, t in enumerate(pieces) if k not in fixed]
        elif el.text:
            out.append((name, _Number(el.text) if number else el.text))
        if el.tail and el.tail.strip():
            out.append((name, el.tail))
        for k, v in el.attrib.items():
            attr = _local(k)
            if _text_attr(name, attr, el):
                out.append((f"{name}@{attr}", v))
    return out


def scan_zip(
    zf: zipfile.ZipFile,
    rules: list[Rule],
    trees: dict[str, etree._ElementTree] | None = None,
    prefix: str = "",
    depth: int = 0,
    limits: ZipLimits | None = None,
) -> list[Leftover]:
    """zip のすべての XML 部品（と埋め込みの Office ファイル）を調べる. trees: 書き換え済みでまだ保存していない部品.

    埋め込みの Office ファイルも zip 爆弾の上限（limits）で確かめ、超えるものは調べずに飛ばす.
    """
    trees = trees or {}
    limits = limits or ZipLimits()
    found: list[Leftover] = []
    for info in zf.infolist():
        name = info.filename
        lower = name.lower()
        if lower.endswith(XML_SUFFIXES):
            tree = trees.get(name)
            strings = _strings(b"" if tree is not None else zf.read(name), tree, name)
            for r in rules:
                rx = r.compiled()
                hits: dict[str, int] = {}
                for where, text in strings:
                    for w, n in _split_sheet_refs(rx, name, where, text):
                        if n:
                            hits[w] = hits.get(w, 0) + n
                if hits:
                    where = "、".join(list(hits)[:3])
                    found.append(Leftover(RULE_ID, prefix + name, where, r.before, sum(hits.values()), reason(prefix + name, where)))
        elif lower.endswith(OFFICE_SUFFIXES) and depth < MAX_DEPTH and info.file_size <= MAX_EMBEDDED:
            try:
                with zipfile.ZipFile(io.BytesIO(zf.read(name))) as inner:
                    check_zip(inner, limits)
                    found += scan_zip(inner, rules, None, f"{prefix}{name}!", depth + 1, limits)
            except (zipfile.BadZipFile, UnsafeZipError):
                continue
    return found
