"""office-replace — Excel / Word / PowerPoint の文字列・シート名・ファイル名を一括置換する.

  office-replace 入力フォルダ -o 出力フォルダ --from "旧社名" --to "新社名"
  office-replace 入力フォルダ -o 出力フォルダ --rules rules.xlsx -r          # 置換表で一括・サブフォルダも
  office-replace 入力フォルダ --rules rules.csv --dry-run                    # 書き換えずに、置換される箇所だけ確認

元のファイルは書き換えません（出力フォルダに書き出す）。置換前には * と ? が使えます（--regex で正規表現）。
レポート: 出力フォルダの office-replace-report.md / .xlsx（--dry-run のときは --report で指定した場所）
終了コード: 0 正常 / 1 エラーのあったファイルあり / 2 引数・置換表の誤り / 3 実行時エラー
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from devtools_common import cli as common_cli

from . import __version__, engine, report
from .rules import Rule, load, validate

TOOL = "office-replace"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog=TOOL, description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("input", type=Path, help="入力フォルダ（またはファイル）")
    p.add_argument("-o", "--out", type=Path, help="出力フォルダ（--dry-run 以外は必須。入力フォルダの中は不可）")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--rules", type=Path, help="置換表（.xlsx / .csv / .tsv / .yaml。列: 置換前・置換後・対象・完全一致）")
    g.add_argument("--from", dest="before", help="置換前（1 組だけのとき）")
    p.add_argument("--to", dest="after", default="", help="置換後（--from と組で使う）")
    p.add_argument("--target", default="text", help="--from の対象: text / sheet / file / all（カンマ区切り可）")
    p.add_argument("-r", "--recursive", action="store_true", help="サブフォルダも対象にする")
    p.add_argument("--regex", action="store_true", help="置換前を正規表現として扱う")
    p.add_argument("-i", "--ignore-case", action="store_true", help="英字の大文字・小文字を区別しない")
    p.add_argument("--whole", action="store_true", help="--from: セル・段落の文字列全体が一致するときだけ")
    p.add_argument("--dry-run", action="store_true", help="書き換えずに、置換される箇所をレポートする")
    p.add_argument("--copy-all", action="store_true", help="置換の無いファイルも出力フォルダにコピーする")
    p.add_argument("--include-masters", action="store_true", help="PowerPoint のスライドマスター・レイアウトも対象にする")
    p.add_argument("--report", type=Path, help="レポートの出力先（.md。同じ名前の .xlsx も作る）")
    p.add_argument("--log-level", default="WARNING", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return p


def _rules(args) -> list[Rule]:
    if args.rules:
        if not args.rules.is_file():
            raise common_cli.UsageError(f"置換表が見つかりません: {args.rules}")
        try:
            rules = load(args.rules, regex=args.regex, ignore_case=args.ignore_case)
        except ValueError as e:
            raise common_cli.UsageError(str(e)) from None
    else:
        from .rules import _targets

        rules = [Rule(args.before, args.after, _targets(args.target), args.regex, args.ignore_case, args.whole)]
    problems = validate(rules)
    if problems:
        raise common_cli.UsageError("置換表の誤り:\n  " + "\n  ".join(problems))
    return rules


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    common_cli.configure_from_args(args)
    if not args.input.exists():
        raise common_cli.UsageError(f"入力が見つかりません: {args.input}")
    rules = _rules(args)
    if not args.dry_run:
        if not args.out:
            raise common_cli.UsageError("出力フォルダ（-o）を指定してください（元のファイルは書き換えません）")
        src_dir = (args.input if args.input.is_dir() else args.input.parent).resolve()
        out = args.out.resolve()
        if out == src_dir or src_dir in out.parents:
            raise common_cli.UsageError("出力フォルダを入力フォルダの中にはできません（次回の実行で二重に置換されるため）")
    out_dir = args.out or Path("_dry_run")
    results = engine.run(args.input, out_dir, rules, engine.Options(args.recursive, args.dry_run, args.copy_all, args.include_masters))
    md = report.markdown(results, dry_run=args.dry_run, src=args.input, out=out_dir)
    report_md = args.report or (None if args.dry_run else out_dir / "office-replace-report.md")
    if report_md:
        report_md.parent.mkdir(parents=True, exist_ok=True)
        report_md.write_text(md, encoding="utf-8", newline="\n")
        report.excel(results, report_md.with_suffix(".xlsx"))
    s = report.summary(results)
    print(
        md
        if args.dry_run
        else f"置換あり {s['changed']} / 置換なし {s['unchanged']} / 対象外 {s['skipped']} / エラー {s['error']}（置換 {s['置換件数']} 件）"
    )
    if report_md:
        print(f"レポート: {report_md} / {report_md.with_suffix('.xlsx')}")
    has_error = any(r.status == "error" or (r.messages and r.status == "changed") for r in results)
    return common_cli.ExitCode.FINDINGS if has_error else common_cli.ExitCode.OK


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))
