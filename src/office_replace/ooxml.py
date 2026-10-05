"""Office Open XML（.xlsx / .docx / .pptx は XML の入った zip）の読み書き.

変更した XML だけを書き直し、それ以外の部品（画像・マクロ・書式など）はバイト単位でそのまま写す。
XML は lxml で読み書きし、名前空間の接頭辞を保つ（Office は接頭辞が変わると開けないことがある）。
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
_PARSER = etree.XMLParser(remove_blank_text=False, huge_tree=True, resolve_entities=False, no_network=True)


@dataclass
class Change:
    file: str
    kind: str  # text / sheet / file
    location: str
    before: str
    after: str
    count: int = 1


class Package:
    def __init__(self, path: Path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.names = self.zip.namelist()
        self._xml: dict[str, etree._ElementTree] = {}
        self.dirty: set[str] = set()

    def close(self) -> None:
        self.zip.close()

    def has(self, name: str) -> bool:
        return name in self.names

    def xml(self, name: str) -> etree._ElementTree:
        if name not in self._xml:
            self._xml[name] = etree.ElementTree(etree.fromstring(self.zip.read(name), _PARSER))
        return self._xml[name]

    def touch(self, name: str) -> None:
        self.dirty.add(name)

    def rels(self, part: str) -> dict[str, str]:
        """部品の関係（rId → 部品のパス）."""
        d, f = posixpath.split(part)
        rel = posixpath.join(d, "_rels", f + ".rels")
        if rel not in self.names:
            return {}
        out = {}
        for r in self.xml(rel).getroot().findall("rel:Relationship", NS):
            if r.get("TargetMode") == "External":
                continue
            target = r.get("Target", "")
            out[r.get("Id", "")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(d, target))
        return out

    def save(self, dest: Path) -> None:
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_name(dest.name + ".tmp")
        try:
            with zipfile.ZipFile(tmp, "w") as out:
                for info in self.zip.infolist():
                    if info.filename in self.dirty:
                        data = etree.tostring(self._xml[info.filename], xml_declaration=True, encoding="UTF-8", standalone=True)
                    else:
                        data = self.zip.read(info.filename)
                    out.writestr(info, data, compress_type=info.compress_type)
            tmp.replace(dest)
        except BaseException:
            tmp.unlink(missing_ok=True)  # 書きかけの一時ファイルを出力フォルダに残さない
            raise


def paragraphs(root: etree._Element, para_tag: str, text_tag: str) -> list[tuple[etree._Element, list[etree._Element]]]:
    """段落ごとの文字要素. 段落の中に入れ子の段落（テキストボックス）があっても、いちばん近い段落に属させる."""
    groups: dict[etree._Element, list[etree._Element]] = {}
    order: list[etree._Element] = []
    for t in root.iter(text_tag):
        p = next((a for a in t.iterancestors() if a.tag == para_tag), None)
        if p is None:
            continue
        if p not in groups:
            groups[p] = []
            order.append(p)
        groups[p].append(t)
    return [(p, groups[p]) for p in order]


def set_text(t: etree._Element, value: str) -> None:
    t.text = value
    if t.tag in (f"{{{NS['w']}}}t", f"{{{NS['s']}}}t") and value != value.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
