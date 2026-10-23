"""Excel のヘッダー・フッターの書式コード（&L・&P・&"フォント"・&12・&K 色・&&）を壊さないこと.

後退: ヘッダー・フッターの文字列を書式コードごと 1 つの文字列として置換していたため、
`P` → `Q` で `&R&P`（右にページ番号）が `&R&Q` になる、`社外秘*` の `*` が `&R&P` まで飲み込んでページ番号が消える、
置換後の `&` が書式コードとして読まれる（`A&B` が「A」+ 太字の切り替え）、などが起きていた.
"""

from __future__ import annotations

import html
from pathlib import Path

import pytest
from lxml import etree
from openpyxl import Workbook
from test_office_replace import _patch_zip, _read, run

from office_replace.processors.spreadsheet import hf_pieces

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _book(path: Path, header: str) -> None:
    """oddHeader に書式コードそのままの文字列を入れたブックを作る（openpyxl を通すと書式コードを組み立て直すため、XML を直接書く）."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "本文"
    ws.oddHeader.center.text = "PLACEHOLDER"
    wb.save(path)
    _patch_zip(path, {"xl/worksheets/sheet1.xml": lambda b: b.replace(b"&amp;CPLACEHOLDER", html.escape(header, quote=False).encode())})
    assert header in html.unescape(_read(path, "xl/worksheets/sheet1.xml"))


def _header(path: Path) -> str:
    root = etree.fromstring(_read(path, "xl/worksheets/sheet1.xml").encode())
    return root.findtext("s:headerFooter/s:oddHeader", namespaces=NS)


@pytest.mark.parametrize(
    ("before", "after", "header", "expected"),
    [
        # 書式コードの外の文字だけを置換する
        ("架空商会", "〇〇", "&L架空商会&R&P / &N", "&L〇〇&R&P / &N"),
        # 書式コードの文字（&P の P、&D の D）には当てない
        ("P", "Q", "&LPマーク&R&P", "&LQマーク&R&P"),
        ("D", "X", "&LD案&C&D &T", "&LX案&C&D &T"),
        # * は書式コードをまたがない（ページ番号が消えない）
        ("社外秘*", "X", "&L社外秘 第1版&R&P", "&LX&R&P"),
        # フォント名・文字の大きさ・色の書式コードの中は置換しない
        ("太字", "細字", '&"ＭＳ ゴシック,太字"&14太字の見出し', '&"ＭＳ ゴシック,太字"&14細字の見出し'),
        ("0", "9", "&KFF0000赤0&K01+000黒", "&KFF0000赤9&K01+000黒"),
        ("14", "99", "&14第14期", "&14第99期"),
        # 置換後の & は && にする（&B などの書式として読まれない）
        ("架空商会", "A&B", "&C架空商会 御中", "&CA&&B 御中"),
        # 文字の &（&&）も置換前に当てられる
        ("A&B商会", "〇〇", "&CA&&B商会 御中", "&C〇〇 御中"),
    ],
)
def test_header_footer_codes_are_kept(tmp_path: Path, before: str, after: str, header: str, expected: str):
    src = tmp_path / "in"
    src.mkdir()
    _book(src / "a.xlsx", header)
    out = tmp_path / "out"
    rc, msg = run(str(src), "--from", before, "--to", after, "-o", str(out))
    assert _header(out / "a.xlsx") == expected, msg
    # 書式コードの中の文字（&P の P など）は「残り」として数えない（終了コード 0）
    assert rc == 0, msg


def test_leftover_in_header_text_is_still_reported(tmp_path: Path):
    # 書式コードを除いた文字に残っていれば、これまでどおり残りとして出す（&& は & として調べる）.
    # 完全一致（--whole）は書式コードを含む文字列の全体で見るので、ヘッダーには当たらず残る
    src = tmp_path / "in"
    src.mkdir()
    _book(src / "a.xlsx", "&LA&&B商会&R&P")
    out = tmp_path / "out"
    rc, msg = run(str(src), "--from", "A&B商会", "--to", "〇〇", "--whole", "-o", str(out))
    assert rc == 1, msg
    assert "残っています" in msg, msg


def test_hf_pieces():
    assert hf_pieces("&L架空&R&P / &N") == (["&L", "架空", "&R", "&P", " / ", "&N"], {0, 2, 3, 5})
    assert hf_pieces("&CA&&B") == (["&C", "A&B"], {0})
    assert hf_pieces('&"Arial,Bold"&12&KFF0000x&P+1') == (['&"Arial,Bold"', "&12", "&KFF0000", "x", "&P+1"], {0, 1, 2, 4})
    assert hf_pieces("文字だけ") == (["文字だけ"], set())
    assert hf_pieces("末尾&") == (["末尾", "&"], {1})  # 知らない書式・壊れた書式もそのまま残す
