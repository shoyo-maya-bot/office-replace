"""残りの確認: 出力ファイルの名前（フォルダ名を含む）に置換前の文字列が残っていれば「残り」として出す."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from test_office_replace import report_of, run


def _rows(md: str) -> list[str]:
    return [line for line in md.splitlines() if line.startswith("|") and "replace.leftover" in line]


def _input(tmp_path: Path) -> Path:
    src = tmp_path / "in"
    (src / "架空商会").mkdir(parents=True)
    wb = Workbook()
    wb.active["A1"] = "架空商会"
    wb.save(src / "架空商会_見積.xlsx")
    Workbook().save(src / "架空商会" / "b.xlsx")
    return src


def test_file_name_containing_before_is_leftover(tmp_path: Path):
    src = _input(tmp_path)
    out = tmp_path / "out"
    rc, msg = run(str(src), "--from", "架空商会", "--to", "〇〇", "-o", str(out))  # 対象は文字列だけ
    assert rc == 1, msg
    assert (out / "架空商会_見積.xlsx").exists()
    rows = _rows(report_of(out).read_text(encoding="utf-8"))
    assert any("（出力ファイル名） 架空商会_見積.xlsx" in r and "--target file" in r for r in rows), rows
    # 試行でも出る
    rc, dry = run(str(src), "--from", "架空商会", "--to", "〇〇", "--dry-run")
    assert rc == 1 and "（出力ファイル名）" in dry
    # --allow-leftover なら 0（残りの規則どおり）
    rc, msg = run(str(src), "--from", "架空商会", "--to", "〇〇", "-o", str(tmp_path / "o2"), "--allow-leftover")
    assert rc == 0, msg


def test_file_name_replaced_but_folder_name_remains(tmp_path: Path):
    src = _input(tmp_path)
    out = tmp_path / "out"
    rc, msg = run(str(src), "--from", "架空商会", "--to", "〇〇", "--target", "text,file", "-r", "--copy-all", "-o", str(out))
    assert rc == 1, msg
    rows = _rows(report_of(out).read_text(encoding="utf-8"))
    assert not any("〇〇_見積.xlsx" in r for r in rows)  # ファイル名は置換済み
    assert any("（出力ファイル名） 架空商会/b.xlsx" in r and "フォルダ名" in r for r in rows), rows


def test_mask_report_hides_leftover_file_name(tmp_path: Path):
    src = _input(tmp_path)
    out = tmp_path / "out"
    run(str(src), "--from", "架空商会", "--to", "〇〇", "-o", str(out), "--mask-report")
    md = report_of(out).read_text(encoding="utf-8")
    assert "（出力ファイル名）" in md and "架空商会" not in md
