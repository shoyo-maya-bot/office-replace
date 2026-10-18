"""office-replace — Excel / Word / PowerPoint の文字列・シート名・ファイル名を一括置換する.

  office-replace 入力フォルダ -o 出力フォルダ --from "旧社名" --to "新社名"
  office-replace 入力フォルダ -o 出力フォルダ --rules rules.xlsx -r          # 置換表で一括・サブフォルダも
  office-replace 入力フォルダ --rules rules.csv --dry-run                    # 書き換えずに、置換される箇所だけ確認

元のファイルは書き換えません（出力フォルダに書き出す）。出力フォルダは入力フォルダ（入力がファイルならそのフォルダ）の外に。
出力フォルダに今回書き出していないファイル（前回の出力の残りなど）があれば警告します（空の出力フォルダを推奨）。
置換前には * と ? が使えます（--regex で正規表現）。
レポート: 出力フォルダと同じ階層の <出力フォルダ名>_report/office-replace-report.md / .xlsx
  （出力フォルダを共有しても置換前の文字列が漏れないよう、出力フォルダの外に置く。--report で場所を指定、--mask-report で伏せる）
終了コード: 0 正常 / 1 エラーのあったファイルあり、または置換前の文字列が残っている（--allow-leftover で 0）
           / 2 引数・置換表の誤り / 3 実行時エラー
  1 ファイルの処理に失敗しても残りのファイルは続けて処理し、レポートは必ず書く（失敗したファイルはエラー、終了コード 1）。
  zip 爆弾対策: 展開後の合計・部品の数・圧縮率が上限（--max-unzip-mb / --max-entries / --max-ratio）を超えるファイルはエラーにする
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
    p.add_argument(
        "-o",
        "--out",
        type=Path,
        help="出力フォルダ（--dry-run 以外は必須。入力フォルダの中は不可。入力がファイルのときは、そのファイルのあるフォルダの中も不可）",
    )
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--rules", type=Path, help="置換表（.xlsx / .csv / .tsv / .yaml。列: 置換前・置換後・対象・完全一致）")
    g.add_argument("--from", dest="before", help="置換前（1 組だけのとき）")
    p.add_argument("--to", dest="after", default="", help="置換後（--from と組で使う）")
    p.add_argument(
        "--target",
        default="text",
        help="--from の対象: text / sheet / file / formula（Excel の数式の中の文字列）/ all（text・sheet・file）。カンマ区切り可",
    )
    p.add_argument("-r", "--recursive", action="store_true", help="サブフォルダも対象にする")
    p.add_argument("--regex", action="store_true", help="置換前を正規表現として扱う")
    p.add_argument("-i", "--ignore-case", action="store_true", help="英字の大文字・小文字を区別しない")
    p.add_argument("--whole", action="store_true", help="--from: セル・段落の文字列全体が一致するときだけ")
    p.add_argument("--dry-run", action="store_true", help="書き換えずに、置換される箇所をレポートする")
    p.add_argument("--copy-all", action="store_true", help="置換の無いファイルも出力フォルダにコピーする")
    p.add_argument("--include-masters", action="store_true", help="PowerPoint のスライドマスター・レイアウトも対象にする")
    p.add_argument(
        "--report",
        type=Path,
        help="レポートの出力先（.md。同じ名前の .xlsx も作る）。既定は <出力フォルダ>_report/office-replace-report.md",
    )
    p.add_argument("--mask-report", action="store_true", help="レポートの置換前の文字列を伏せる（ファイル名・場所は置換後の名前で出す）")
    p.add_argument(
        "--allow-leftover", action="store_true", help="置換前の文字列が残っていても終了コードを 0 にする（以前の動き。レポートには出す）"
    )
    lim = engine.ZipLimits()
    p.add_argument(
        "--max-unzip-mb",
        type=_positive(int),
        default=lim.max_total // 1024**2,
        help=f"zip 爆弾対策: 1 ファイルを展開した後の合計の上限（MB）。既定 {lim.max_total // 1024**2}",
    )
    p.add_argument(
        "--max-entries",
        type=_positive(int),
        default=lim.max_entries,
        help=f"zip 爆弾対策: 1 ファイルの部品の数の上限。既定 {lim.max_entries}",
    )
    p.add_argument(
        "--max-ratio",
        type=_positive(float),
        default=lim.max_ratio,
        help=f"zip 爆弾対策: 部品の圧縮率（展開後 / 圧縮後）の上限。既定 {lim.max_ratio:g}",
    )
    p.add_argument("--log-level", default="WARNING", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return p


def _positive(kind):
    def conv(value: str):
        try:
            v = kind(value)
        except ValueError:
            raise argparse.ArgumentTypeError(f"数で指定してください: {value}") from None
        if v <= 0:
            raise argparse.ArgumentTypeError(f"0 より大きい数で指定してください: {value}")
        return v

    return conv


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
            where = "入力フォルダ" if args.input.is_dir() else f"入力ファイルのあるフォルダ（{src_dir.name}）"
            raise common_cli.UsageError(f"出力フォルダを{where}の中にはできません（次回の実行で二重に置換されるため）")
    out_dir = args.out or Path("_dry_run")
    report_md = args.report or (None if args.dry_run else default_report(out_dir))
    if report_md and not args.dry_run and _inside(report_md, out_dir):
        print(
            "警告: レポートを出力フォルダの中に置きます。レポートには置換前の文字列が入るので、"
            "出力フォルダを共有するときは消すか --mask-report を付けてください",
            file=sys.stderr,
        )
    results = engine.run(
        args.input,
        out_dir,
        rules,
        engine.Options(
            args.recursive,
            args.dry_run,
            args.copy_all,
            args.include_masters,
            engine.ZipLimits(max_total=args.max_unzip_mb * 1024**2, max_entries=args.max_entries, max_ratio=args.max_ratio),
        ),
    )
    mask = report.masker(rules) if args.mask_report else None
    md = report.markdown(results, dry_run=args.dry_run, src=args.input, out=out_dir, mask=mask)
    if report_md:
        report_md.parent.mkdir(parents=True, exist_ok=True)
        report_md.write_text(md, encoding="utf-8", newline="\n")
        report.excel(results, report_md.with_suffix(".xlsx"), mask=mask)
    s = report.summary(results)
    left_note = ""
    if s["残り"]:
        left_note = f"\n注意: 置換前の文字列が {s['残り']} 件残っています（レポートの「残っている文字列」の理由と対処を確認してください）"
        left_note += "" if args.allow_leftover else "。終了コードは 1 です（--allow-leftover で 0）"
    if s["警告"]:
        left_note += f"\n警告: {s['警告']} 件（* の一致が長すぎるなど。レポートの「エラー・注意」を確認してください）"
    print(
        md
        if args.dry_run
        else f"置換あり {s['changed']} / 置換なし {s['unchanged']} / 対象外 {s['skipped']} / エラー {s['error']}（置換 {s['置換件数']} 件）"
        + left_note
    )
    if args.dry_run and left_note:
        print(left_note.strip())
    if report_md:
        print(f"レポート: {report_md} / {report_md.with_suffix('.xlsx')}")
    if not args.dry_run:
        _warn_stale(out_dir, results, report_md)
    has_error = any(r.status == "error" or (r.messages and r.status == "changed") for r in results)
    has_leftover = bool(s["残り"]) and not args.allow_leftover
    return common_cli.ExitCode.FINDINGS if has_error or has_leftover else common_cli.ExitCode.OK


def default_report(out_dir: Path) -> Path:
    """既定のレポートの場所: 出力フォルダと同じ階層の <出力フォルダ名>_report/（出力フォルダを共有しても漏れないように）."""
    out = out_dir.resolve()
    return out.parent / f"{out.name}_report" / "office-replace-report.md"


STALE_SHOW = 5


def stale_outputs(out_dir: Path, results: list[engine.FileResult], report_md: Path | None) -> list[Path]:
    """出力フォルダにあるが、今回書き出していないファイル（前回の出力の残りなど）. 出力フォルダからの相対パス.

    前回の置換表で書き出したファイルや、今回は置換が無かったファイルの前回の出力が残っていると、
    出力フォルダを共有したときに置換前の文字列が漏れるため、警告に使う.
    """
    if not out_dir.is_dir():
        return []
    written = {r.dest.resolve() for r in results if r.dest is not None}
    if report_md:
        written |= {report_md.resolve(), report_md.with_suffix(".xlsx").resolve()}
    base = out_dir.resolve()
    return sorted(p.relative_to(base) for p in base.rglob("*") if p.is_file() and p.resolve() not in written)


def _warn_stale(out_dir: Path, results: list[engine.FileResult], report_md: Path | None) -> None:
    stale = stale_outputs(out_dir, results, report_md)
    if not stale:
        return
    names = "\n  ".join(p.as_posix() for p in stale[:STALE_SHOW])
    more = f"\n  …ほか {len(stale) - STALE_SHOW} 件" if len(stale) > STALE_SHOW else ""
    print(
        f"警告: 出力フォルダに、今回書き出していないファイルが {len(stale)} 件あります（前回の出力の残りなど）。"
        "置換前の文字列が入っている場合があるので、共有する前に確認するか、空の出力フォルダで実行し直してください:\n  " + names + more,
        file=sys.stderr,
    )


def _inside(path: Path, folder: Path) -> bool:
    return path.resolve().is_relative_to(folder.resolve())


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))
