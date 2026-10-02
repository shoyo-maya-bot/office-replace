"""office-replace のテスト. サンプル生成スクリプトで入力を作り、置換後のファイルを各ライブラリで開いて確かめる."""

from __future__ import annotations

import contextlib
import importlib.util
import io
import zipfile
from pathlib import Path

import pytest
from devtools_common.cli import run_main
from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation

from office_replace.cli import main
from office_replace.rules import Rule, load, validate, wildcard_to_regex
from office_replace.sheetname import check_sheet_name
from office_replace.textruns import replace_runs

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("make_samples", ROOT / "samples/make_samples.py")
assert _spec and _spec.loader
make_samples = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(make_samples)


@pytest.fixture()
def samples(tmp_path: Path) -> Path:
    make_samples.generate(tmp_path / "s")
    return tmp_path / "s"


def run(*args: str) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(main, list(args))
    return rc, out.getvalue() + err.getvalue()


# --- ルール ----------------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("pattern", "text", "hit"),
    [
        ("会員*管理", "会員情報管理", True),
        ("担当: ?田", "担当: 山田", True),
        ("担当: ?田", "担当: 田中", False),
        ("100~*", "100*", True),
        ("100~*", "1000", False),
        ("a.b", "axb", False),  # ワイルドカード以外の記号はそのまま
    ],
)
def test_wildcard(pattern, text, hit):
    import re

    assert bool(re.search(wildcard_to_regex(pattern), text)) is hit


def test_rule_regex_groups_and_whole():
    assert Rule(r"第(\d+)版", r"Ver.\1", regex=True).apply("第12版") == ("Ver.12", 1)
    assert Rule("商事", "X", whole=True).apply("サンプル商事") == ("サンプル商事", 0)
    assert Rule("商事", "X", whole=True).apply("商事") == ("X", 1)
    assert Rule("abc", "X", ignore_case=True).apply("ABC") == ("X", 1)


def test_validate():
    assert validate([Rule("*", "x")])  # すべてに一致
    assert validate([Rule("(", "x", regex=True)])
    assert validate([Rule("a", "b", targets=("cell",))])
    assert validate([Rule("x*", "y")]) == []


def test_load_rules_csv_xlsx_cp932(samples: Path, tmp_path: Path):
    for f in ("rules.csv", "rules.xlsx"):
        rules = load(samples / f)
        assert [r.targets for r in rules][:2] == [("text", "sheet", "file"), ("text",)]
    p = tmp_path / "sjis.csv"
    p.write_text("置換前,置換後,対象,完全一致\n旧,新,シート名,○\n", encoding="cp932")
    r = load(p)[0]
    assert (r.before, r.after, r.targets, r.whole) == ("旧", "新", ("sheet",), True)
    (tmp_path / "empty.csv").write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(ValueError, match="置換前"):
        load(tmp_path / "empty.csv")


def test_replace_runs_across_formatting():
    assert replace_runs(["会員", "管理", "システム"], Rule("会員管理", "顧客管理")) == (["顧客管理", "", "システム"], 1)
    assert replace_runs(["ab", "cd", "ef"], Rule("b*e", "-")) == (["a-", "", "f"], 1)
    assert replace_runs(["xx", "yy"], Rule("x", "Z")) == (["ZZ", "yy"], 2)


@pytest.mark.parametrize(
    ("name", "problem"),
    [
        ("a" * 32, "31 文字"),
        ("集計[1]", "使えない文字"),
        ("a/b", "使えない文字"),
        ("History", "予約語"),
        ("集計", "同じ名前"),
        ("'x", "アポストロフィ"),
        ("", "空"),
    ],
)
def test_sheet_name_checks(name, problem):
    assert problem in check_sheet_name(name, ["集計"])


# --- 置換の実行 --------------------------------------------------------------------------------


def test_replace_folder(samples: Path, tmp_path: Path):
    out = tmp_path / "out"
    rc, msg = run(str(samples / "input"), "--rules", str(samples / "rules.xlsx"), "-r", "-o", str(out))
    assert rc == 0, msg
    # Excel: セル・図形・シート名・シートを参照する数式
    wb = load_workbook(out / "見積書_〇〇商事.xlsx")
    assert wb.sheetnames == ["〇〇商事向け", "集計"]
    ws = wb["〇〇商事向け"]
    assert [ws["A2"].value, ws["A3"].value, ws["A4"].value, ws["A7"].value] == [
        "〇〇商事 御中",
        "（支店名）",
        "担当: （担当者）",
        "設計（2027年度）",
    ]
    assert ws["B7"].value == 1200000  # 数値は変えない
    assert wb["集計"]["B1"].value == "='〇〇商事向け'!B9"
    with zipfile.ZipFile(out / "見積書_〇〇商事.xlsx") as z:
        drawing = z.read("xl/drawings/drawing1.xml").decode()
    assert "<a:t>〇〇商事</a:t>" in drawing and 'b="1"' in drawing  # 書式付きのランに入る
    # Word: ランをまたぐ文字列・書式・ヘッダー・表
    doc = Document(out / "提案書_〇〇商事.docx")
    para = doc.paragraphs[1]
    assert para.text == "本書は〇〇商事様の 2027年度 の業務改善についての提案です。"
    assert para.runs[1].text == "〇〇商事" and para.runs[1].bold
    assert doc.sections[0].header.paragraphs[0].text == "〇〇商事様 提案書（社外秘）"
    assert doc.tables[0].cell(1, 1).text == "〇〇商事の受注業務"
    # PowerPoint（サブフォルダ）: 図形・表・ノート
    prs = Presentation(out / "sub/説明資料_〇〇商事.pptx")
    s1, s2 = prs.slides
    assert s1.shapes.title.text == "〇〇商事様 ご説明資料"
    assert s1.notes_slide.notes_text_frame.text == "〇〇商事の部長様が出席予定"
    texts = [sh.text_frame.text for sh in s2.shapes if sh.has_text_frame]
    assert "窓口: 〇〇商事 情報システム部" in texts
    table = next(sh for sh in s2.shapes if sh.has_table).table
    assert table.cell(1, 1).text == "（支店名）"
    # 元のファイルは変わらない・古い形式は出力しない
    assert load_workbook(samples / "input/見積書_サンプル商事.xlsx").sheetnames[0] == "サンプル商事向け"
    assert not (out / "旧形式.xls").exists()
    # レポート
    md = (out / "office-replace-report.md").read_text(encoding="utf-8")
    assert "シート名" in md and "古い形式" in md and "1 か所を新しいシート名に更新" in md
    assert load_workbook(out / "office-replace-report.xlsx").sheetnames == ["ファイル", "置換箇所"]


def test_unchanged_parts_are_copied_byte_for_byte(samples: Path, tmp_path: Path):
    out = tmp_path / "out"
    run(str(samples / "input"), "--from", "2026年度", "--to", "2027年度", "-o", str(out))
    with zipfile.ZipFile(samples / "input/提案書_サンプル商事.docx") as a, zipfile.ZipFile(out / "提案書_サンプル商事.docx") as b:
        assert a.namelist() == b.namelist()
        for n in a.namelist():
            if n != "word/document.xml":
                assert a.read(n) == b.read(n), n


def test_dry_run_writes_nothing(samples: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.chdir(tmp_path)
    rc, out = run(str(samples / "input"), "--rules", str(samples / "rules.csv"), "-r", "--dry-run")
    assert rc == 0
    assert "試行" in out and "| 見積書_サンプル商事.xlsx | シート名 |" in out
    assert not (tmp_path / "_dry_run").exists()


def test_invalid_sheet_name_is_skipped_and_reported(samples: Path, tmp_path: Path):
    out = tmp_path / "out"
    rc, msg = run(
        str(samples / "input/見積書_サンプル商事.xlsx"),
        "--from",
        "サンプル商事向け",
        "--to",
        "見積/2027",
        "--target",
        "sheet",
        "-o",
        str(out),
    )
    assert rc == 1
    assert "使えない文字" in (out / "office-replace-report.md").read_text(encoding="utf-8")


def test_sheet_name_too_long(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    wb = Workbook()
    wb.active.title = "短い"
    wb.save(src / "a.xlsx")
    rc, _ = run(str(src), "--from", "短い", "--to", "長" * 32, "--target", "sheet", "-o", str(tmp_path / "out"))
    assert rc == 1
    assert "31 文字" in (tmp_path / "out/office-replace-report.md").read_text(encoding="utf-8")


def test_file_name_collision(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    for n in ("資料A.docx", "資料B.docx"):
        Document().save(src / n)
    rc, _ = run(str(src), "--from", "資料?", "--to", "資料", "--target", "file", "-o", str(tmp_path / "out"))
    assert rc == 1
    assert len(list((tmp_path / "out").glob("*.docx"))) == 1


def test_output_inside_input_is_refused(samples: Path):
    rc, out = run(str(samples / "input"), "--from", "a", "--to", "b", "-o", str(samples / "input/out"))
    assert rc == 2 and "入力フォルダの中" in out


def test_broken_file_is_error(tmp_path: Path):
    src = tmp_path / "in"
    src.mkdir()
    (src / "pw.xlsx").write_bytes(b"not a zip")
    rc, _ = run(str(src), "--from", "a", "--to", "b", "-o", str(tmp_path / "out"))
    assert rc == 1
    assert "開けません" in (tmp_path / "out/office-replace-report.md").read_text(encoding="utf-8")


def test_copy_all(samples: Path, tmp_path: Path):
    out = tmp_path / "out"
    run(str(samples / "input"), "--from", "存在しない語", "--to", "x", "-r", "--copy-all", "-o", str(out))
    assert (out / "提案書_サンプル商事.docx").exists() and (out / "sub/説明資料_サンプル商事.pptx").exists()
