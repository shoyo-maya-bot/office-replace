"""第 6 回: 出力フォルダの前回の残りの警告・YAML の置換表の形の誤り・単一ファイル入力の出力先."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import pytest
from devtools_common.cli import run_main
from docx import Document

from office_replace.cli import main
from office_replace.rules import load


def run(*args: str) -> tuple[int, str, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(main, list(args))
    return rc, out.getvalue(), err.getvalue()


def _doc(path: Path, text: str) -> None:
    d = Document()
    d.add_paragraph(text)
    d.save(path)


# --- 出力フォルダの前回の残り -------------------------------------------------------------


def test_stale_output_from_previous_run_is_warned(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    _doc(src / "a.docx", "架空商事の資料")
    _doc(src / "b.docx", "架空工業の資料")
    out = tmp_path / "out"
    # 1 回目: 両方を置換
    rc, _, err = run(str(src), "-o", str(out), "--rules", str(_rules(tmp_path, ["架空商事", "架空工業"])))
    assert rc == 0, err
    assert "今回書き出していない" not in err  # 空の出力フォルダなら警告しない
    # 2 回目: 置換表から 架空工業 を外す → b.docx は書き出さないが、前回の b.docx が残る
    rc, _, err = run(str(src), "-o", str(out), "--rules", str(_rules(tmp_path, ["架空商事"])))
    assert rc == 0  # 終了コードは変えない（互換）
    assert "今回書き出していないファイルが 1 件" in err
    assert "b.docx" in err
    assert "a.docx" not in err.split("確認するか")[-1]  # 今回書き出したものは出さない


def test_stale_warning_ignores_report_inside_out_and_lists_at_most_five(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    _doc(src / "a.docx", "架空商事")
    out = tmp_path / "out"
    out.mkdir()
    for i in range(7):
        (out / f"old{i}.txt").write_text("x", encoding="utf-8")
    rep = out / "r.md"
    rc, _, err = run(str(src), "-o", str(out), "--from", "架空商事", "--to", "サンプル", "--report", str(rep))
    assert rc == 0
    assert "今回書き出していないファイルが 7 件" in err
    assert "ほか 2 件" in err
    assert "r.md" not in err.split("今回書き出していない")[-1]
    assert "r.xlsx" not in err


def test_dry_run_does_not_warn_stale(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    _doc(src / "a.docx", "架空商事")
    out = tmp_path / "out"
    out.mkdir()
    (out / "old.docx").write_bytes(b"x")
    rc, _, err = run(str(src), "-o", str(out), "--from", "架空商事", "--to", "サンプル", "--dry-run")
    assert "今回書き出していない" not in err


def _rules(tmp_path: Path, befores: list[str]) -> Path:
    p = tmp_path / f"rules{len(befores)}.csv"
    p.write_text("置換前,置換後\n" + "".join(f"{b},サンプル\n" for b in befores), encoding="utf-8")
    return p


# --- YAML の置換表の形 --------------------------------------------------------------------


@pytest.mark.parametrize(
    "text",
    [
        "置換前: 架空商事\n置換後: サンプル\n",  # トップが 1 件のマッピング（- が無い）
        "rules:\n  置換前: 架空商事\n  置換後: サンプル\n",  # rules: の下が 1 件のマッピング
        "rules:\n  - 架空商事\n",  # 行がマッピングでない
        "rules: 架空商事\n",  # 文字列
    ],
)
def test_yaml_wrong_shape_is_usage_error(tmp_path: Path, text: str):
    p = tmp_path / "r.yaml"
    p.write_text(text, encoding="utf-8")
    with pytest.raises(ValueError, match="リスト|組になっていません|形が違います"):
        load(p)
    src = tmp_path / "in"
    src.mkdir()
    _doc(src / "a.docx", "架空商事")
    rc, _, err = run(str(src), "--rules", str(p), "--dry-run")
    assert rc == 2, err  # 以前は TypeError で 3
    assert "TypeError" not in err


def test_yaml_empty_still_reports_no_rules(tmp_path: Path):
    p = tmp_path / "r.yaml"
    p.write_text("rules:\n", encoding="utf-8")
    with pytest.raises(ValueError, match="ルールがありません"):
        load(p)


# --- 単一ファイル入力の出力先 ----------------------------------------------------------------


def test_single_file_output_inside_its_folder_is_refused_with_clear_message(tmp_path: Path):
    src = tmp_path / "docs"
    src.mkdir()
    _doc(src / "a.docx", "架空商事")
    rc, _, err = run(str(src / "a.docx"), "-o", str(src / "out"), "--from", "架空商事", "--to", "サンプル")
    assert rc == 2
    assert "入力ファイルのあるフォルダ（docs）" in err
    rc, _, err = run(str(src / "a.docx"), "-o", str(tmp_path / "out"), "--from", "架空商事", "--to", "サンプル")
    assert rc == 0, err
    assert (tmp_path / "out" / "a.docx").is_file()
