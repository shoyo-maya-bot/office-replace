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
    # 文書のプロパティ
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "op": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    # Excel の拡張（入力規則・条件付き書式）とスレッド形式のコメント
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
    # PowerPoint の新しいコメント、Word のコメントの作成者
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}
_PARSER = etree.XMLParser(remove_blank_text=False, huge_tree=True, resolve_entities=False, no_network=True)


class UnsafeZipError(Exception):
    """展開すると大きすぎる・部品が多すぎる・圧縮率が高すぎる zip（zip 爆弾の疑い）."""


class BrokenOutputError(Exception):
    """置換すると壊れたファイル（Office が修復を求めるもの）になるため、このファイルは出力しない."""


@dataclass(frozen=True)
class ZipLimits:
    """zip 爆弾対策の上限. 既定値は普通の Office ファイルには十分な大きさ. CLI の --max-* で変えられる."""

    max_total: int = 2 * 1024**3  # 展開後の合計（バイト）
    max_entries: int = 10_000  # 部品の数
    max_ratio: float = 200.0  # 部品ごとの圧縮率（展開後 / 圧縮後）
    ratio_min_size: int = 1024 * 1024  # 圧縮率はこの大きさ（展開後）を超える部品だけ見る（小さい XML は圧縮率が高くて普通）


def check_zip(zf: zipfile.ZipFile, limits: ZipLimits) -> None:
    """中身を読む前に、zip の目録（申告されたサイズ）で上限を確かめる. 超えれば UnsafeZipError.

    zipfile は目録のサイズより多くは展開しない（超える分は読まず、CRC が合わなければ BadZipFile）ので、目録で確かめれば足りる.
    """
    infos = zf.infolist()
    if len(infos) > limits.max_entries:
        raise UnsafeZipError(f"部品の数が {len(infos)} 個で、上限 {limits.max_entries} 個を超えています")
    total = 0
    for info in infos:
        total += info.file_size
        if total > limits.max_total:
            raise UnsafeZipError(f"展開後の大きさが上限 {limits.max_total:,} バイトを超えます")
        if info.file_size > limits.ratio_min_size:
            ratio = info.file_size / max(info.compress_size, 1)
            if ratio > limits.max_ratio:
                raise UnsafeZipError(f"部品 {info.filename} の圧縮率が {ratio:.0f} 倍で、上限 {limits.max_ratio:g} 倍を超えています")


@dataclass
class Change:
    file: str
    kind: str  # text / sheet / file
    location: str
    before: str
    after: str
    count: int = 1


class Package:
    def __init__(self, path: Path, limits: ZipLimits | None = None):
        self.path = path
        self.limits = limits or ZipLimits()
        self.zip = zipfile.ZipFile(path)
        try:
            check_zip(self.zip, self.limits)
        except UnsafeZipError:
            self.zip.close()
            raise
        self.names = self.zip.namelist()
        self._xml: dict[str, etree._ElementTree] = {}
        self.dirty: set[str] = set()

    def close(self) -> None:
        self.zip.close()

    def has(self, name: str) -> bool:
        return name in self.names

    def xml(self, name: str) -> etree._ElementTree:
        if name not in self._xml:
            self._xml[name] = etree.ElementTree(etree.fromstring(self.read(name), _PARSER))
        return self._xml[name]

    def read(self, name: str) -> bytes:
        return self.zip.read(name)

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
                        data = self.read(info.filename)
                    out.writestr(info, data, compress_type=info.compress_type)
            tmp.replace(dest)
        except BaseException:
            tmp.unlink(missing_ok=True)  # 書きかけの一時ファイルを出力フォルダに残さない
            raise


# 段落の中の改行・タブの要素 → 段落の文字列での表し方. 置換は改行・タブをまたがない（textruns.replace_runs の fixed）
BREAKS = {
    f"{{{NS['w']}}}br": "\n",  # Word の改行（Shift+Enter. ページ・段の区切りも）
    f"{{{NS['w']}}}cr": "\n",
    f"{{{NS['w']}}}tab": "\t",  # Word のタブ（w:r の子. 段落の書式のタブ位置 w:tabs/w:tab は文字ではない）
    f"{{{NS['a']}}}br": "\n",  # PowerPoint・図形の改行（DrawingML のタブは a:t の中の文字）
}
_RUN_TAGS = {f"{{{NS['w']}}}r", f"{{{NS['a']}}}p"}  # 改行・タブとして数える要素の親（w:br・w:tab は w:r、a:br は a:p の子）


def is_break(el: etree._Element) -> bool:
    return el.tag in BREAKS


def paragraphs(
    root: etree._Element, para_tag: str, text_tag: str, *, breaks: bool = False
) -> list[tuple[etree._Element, list[etree._Element]]]:
    """段落ごとの文字要素（文書の順）. 段落の中に入れ子の段落（テキストボックス）があっても、いちばん近い段落に属させる.

    breaks: 改行・タブの要素（BREAKS）も並びに入れる. 文字列にするときは BREAKS の文字にし、置換で書き換えない.
    """
    tags = [text_tag] + ([t for t in BREAKS if t.split("}")[0] == text_tag.split("}")[0]] if breaks else [])
    groups: dict[etree._Element, list[etree._Element]] = {}
    order: list[etree._Element] = []
    for t in root.iter(*tags):
        if t.tag != text_tag:
            parent = t.getparent()
            if parent is None or parent.tag not in _RUN_TAGS:
                continue
        p = next((a for a in t.iterancestors() if a.tag == para_tag), None)
        if p is None:
            continue
        if p not in groups:
            groups[p] = []
            order.append(p)
        groups[p].append(t)
    # 改行・タブしかない段落は置換するものがない
    return [(p, groups[p]) for p in order if any(not is_break(t) for t in groups[p])]


def set_text(t: etree._Element, value: str) -> None:
    t.text = value
    if t.tag in (f"{{{NS['w']}}}t", f"{{{NS['s']}}}t") and value != value.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
