"""書き出したファイルに、置換前の文字列が残っていないかを調べる（固有名を伏せる用途の確認）.

ツールが置換しない場所（数式の中の文字列・変更履歴・グラフの元データの埋め込みブックなど）に残った文字列を、
注意（rule_id: replace.leftover）としてレポートに出す。対象は「文字列」が対象の、正規表現でないルール。
"""

from __future__ import annotations

import html
import io
import zipfile
from dataclasses import dataclass

from lxml import etree

from .ooxml import _PARSER
from .rules import Rule

RULE_ID = "replace.leftover"
XML_SUFFIXES = (".xml", ".rels", ".vml")
OFFICE_SUFFIXES = (".xlsx", ".xlsm", ".docx", ".docm", ".pptx", ".pptm")
MAX_EMBEDDED = 50 * 1024 * 1024  # 埋め込みのファイルはこの大きさまで調べる
MAX_DEPTH = 2


@dataclass
class Leftover:
    rule_id: str
    part: str  # zip の中の部品（埋め込みは「外の部品!中の部品」）
    where: str  # 要素名・属性名（最大 3 つ）
    before: str  # ルールの置換前
    count: int


def scan_rules(rules: list[Rule]) -> list[Rule]:
    """調べるルール. 置換後に置換前が含まれるルール（集計 → 集計_2027 など）は、置換後の文字に当たってしまうため除く."""
    return [r for r in rules if not r.regex and "text" in r.targets and not r.compiled().search(r.after)]


def _local(tag) -> str:
    return tag.split("}", 1)[1] if isinstance(tag, str) and "}" in tag else str(tag)


def _strings(data: bytes, tree: etree._ElementTree | None) -> list[tuple[str, str]]:
    """部品の中の（場所, 文字列）. 文字・属性の値ごとに調べる."""
    if tree is None:
        try:
            tree = etree.ElementTree(etree.fromstring(data, _PARSER))
        except etree.XMLSyntaxError:  # 古い VML など、整形式でないもの
            return [("（XML として読めない部品）", html.unescape(data.decode("utf-8", "replace")))]
    out: list[tuple[str, str]] = []
    for el in tree.getroot().iter():
        name = _local(el.tag) if isinstance(el.tag, str) else "comment"
        if el.text:
            out.append((name, el.text))
        if el.tail and el.tail.strip():
            out.append((name, el.tail))
        for k, v in el.attrib.items():
            out.append((f"{name}@{_local(k)}", v))
    return out


def scan_zip(
    zf: zipfile.ZipFile, rules: list[Rule], trees: dict[str, etree._ElementTree] | None = None, prefix: str = "", depth: int = 0
) -> list[Leftover]:
    """zip のすべての XML 部品（と埋め込みの Office ファイル）を調べる. trees: 書き換え済みでまだ保存していない部品."""
    trees = trees or {}
    found: list[Leftover] = []
    for info in zf.infolist():
        name = info.filename
        lower = name.lower()
        if lower.endswith(XML_SUFFIXES):
            tree = trees.get(name)
            strings = _strings(b"" if tree is not None else zf.read(name), tree)
            for r in rules:
                rx = r.compiled()
                hits: dict[str, int] = {}
                for where, text in strings:
                    n = sum(1 for m in rx.finditer(text) if m.end() > m.start())
                    if n:
                        hits[where] = hits.get(where, 0) + n
                if hits:
                    found.append(Leftover(RULE_ID, prefix + name, "、".join(list(hits)[:3]), r.before, sum(hits.values())))
        elif lower.endswith(OFFICE_SUFFIXES) and depth < MAX_DEPTH and info.file_size <= MAX_EMBEDDED:
            try:
                with zipfile.ZipFile(io.BytesIO(zf.read(name))) as inner:
                    found += scan_zip(inner, rules, None, f"{prefix}{name}!", depth + 1)
            except zipfile.BadZipFile:
                continue
    return found
