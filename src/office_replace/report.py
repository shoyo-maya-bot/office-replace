"""置換の結果のレポート（Markdown / Excel）."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .engine import FileResult

STATUS = {"changed": "置換あり", "unchanged": "置換なし", "skipped": "対象外", "error": "エラー"}
KIND = {"text": "文字列", "sheet": "シート名", "file": "ファイル名"}


def summary(results: list[FileResult]) -> dict[str, int]:
    out = {k: 0 for k in STATUS}
    for r in results:
        out[r.status] += 1
    out["置換件数"] = sum(c.count for r in results for c in r.changes)
    out["残り"] = sum(x.count for r in results for x in r.leftovers)
    return out


def _cell(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def markdown(results: list[FileResult], *, dry_run: bool, src: Path, out: Path) -> str:
    s = summary(results)
    lines = [
        "# 置換の結果" + ("（試行: ファイルは書き換えていません）" if dry_run else ""),
        "",
        f"- 入力: {src}",
        f"- 出力: {out if not dry_run else '（試行のため書き出しなし）'}",
        f"- ファイル: 置換あり {s['changed']} / 置換なし {s['unchanged']} / 対象外 {s['skipped']} / エラー {s['error']}",
        f"- 置換: {s['置換件数']} 件",
        f"- 置換しきれずに残った置換前の文字列: {s['残り']} 件" + ("（下の「残っている文字列」を確認してください）" if s["残り"] else ""),
        "",
    ]
    errs = [(r, m) for r in results for m in r.messages if r.status in ("error", "changed")]
    if errs:
        lines += ["## エラー・注意", "", "| ファイル | 内容 |", "|---|---|"]
        lines += [f"| {_cell(r.source.name)} | {_cell(m)} |" for r, m in errs]
        lines.append("")
    rows = [(r, c) for r in results for c in r.changes]
    if rows:
        lines += ["## 置換した箇所", "", "| ファイル | 種類 | 場所 | 置換前 | 置換後 | 件数 |", "|---|---|---|---|---|---|"]
        lines += [
            f"| {_cell(r.source.name)} | {KIND.get(c.kind, c.kind)} | {_cell(c.location)} | {_cell(c.before)} | {_cell(c.after)} | {c.count} |"
            for r, c in rows
        ]
        lines.append("")
    left = [(r, x) for r in results for x in r.leftovers]
    if left:
        lines += [
            "## 残っている文字列（注意）",
            "",
            "置換後のファイルの中に、置換前の文字列がまだあります（数式の中の文字列・変更履歴・グラフの元データなど、このツールが置換しない場所）。",
            "固有名を伏せる目的のときは、Office で開いて確認してください。",
            "",
            "| ファイル | 種類 | 部品 | 場所 | 置換前 | 件数 |",
            "|---|---|---|---|---|---|",
        ]
        lines += [
            f"| {_cell(r.source.name)} | {x.rule_id} | {_cell(x.part)} | {_cell(x.where)} | {_cell(x.before)} | {x.count} |"
            for r, x in left
        ]
        lines.append("")
    skipped = [r for r in results if r.status == "skipped" and r.messages and "古い形式" in r.messages[0]]
    if skipped:
        lines += ["## 対象外（古い形式）", ""] + [f"- {r.source.name}: {r.messages[0]}" for r in skipped] + [""]
    return "\n".join(lines)


def excel(results: list[FileResult], path: Path) -> None:
    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="DDEBF7")

    def sheet(ws, header: list[str], rows: list[list], widths: list[int]) -> None:
        ws.append(header)
        for i, w in enumerate(widths, start=1):
            ws.cell(1, i).font = Font(bold=True)
            ws.cell(1, i).fill = head_fill
            ws.column_dimensions[get_column_letter(i)].width = w
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    ws = wb.active
    ws.title = "ファイル"
    sheet(
        ws,
        ["ファイル", "出力先", "結果", "置換件数", "メッセージ"],
        [[str(r.source), str(r.dest or ""), STATUS[r.status], sum(c.count for c in r.changes), " / ".join(r.messages)] for r in results],
        [50, 50, 10, 10, 60],
    )
    sheet(
        wb.create_sheet("置換箇所"),
        ["ファイル", "種類", "場所", "置換前", "置換後", "件数"],
        [[r.source.name, KIND.get(c.kind, c.kind), c.location, c.before, c.after, c.count] for r in results for c in r.changes],
        [30, 10, 30, 50, 50, 8],
    )
    sheet(
        wb.create_sheet("残り"),
        ["ファイル", "種類", "部品", "場所", "置換前", "件数"],
        [[r.source.name, x.rule_id, x.part, x.where, x.before, x.count] for r in results for x in r.leftovers],
        [30, 18, 40, 30, 30, 8],
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
