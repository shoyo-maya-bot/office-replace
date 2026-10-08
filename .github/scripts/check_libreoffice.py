"""置換後のファイルを LibreOffice（headless）で開けるかを確かめる（CI 用）.

    python .github/scripts/check_libreoffice.py 出力フォルダ 作業フォルダ

- 出力フォルダの .xlsx / .docx / .pptx をすべて PDF に変換できること（開けない・壊れているファイルは変換に失敗する）
- サンプルの見積書は LibreOffice で .xlsx に保存し直し、テーブルの列名と構造化参照の数式が計算できること
"""

from __future__ import annotations

import shutil
import subprocess  # nosec B404 — CI で LibreOffice を動かすため
import sys
from pathlib import Path

from openpyxl import load_workbook


def convert(soffice: str, files: list[Path], fmt: str, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore", "--convert-to", fmt, "--outdir", str(outdir), *map(str, files)]
    subprocess.run(cmd, check=True, timeout=600)  # nosec B603 — 引数は固定のコマンドとファイル名だけ


def main(out: Path, work: Path) -> int:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice（soffice）が見つかりません", file=sys.stderr)
        return 2
    files = sorted(p for p in out.rglob("*") if p.suffix.lower() in (".xlsx", ".docx", ".pptx"))
    if not files:
        print(f"{out} に Office のファイルがありません", file=sys.stderr)
        return 1
    convert(soffice, files, "pdf", work / "pdf")
    problems = [
        f.name for f in files if not (work / "pdf" / (f.stem + ".pdf")).is_file() or (work / "pdf" / (f.stem + ".pdf")).stat().st_size == 0
    ]
    estimate = next((f for f in files if f.suffix == ".xlsx" and f.name.startswith("見積書")), None)
    if estimate:
        convert(soffice, [estimate], "xlsx", work / "roundtrip")
        rt = work / "roundtrip" / estimate.name
        wb = load_workbook(rt, data_only=True)
        ws = wb["集計"]
        if ws["B3"].value != 2 or ws["B1"].value != 4600000:
            problems.append(f"{estimate.name}: 計算結果が違います（集計!B3={ws['B3'].value!r}, 集計!B1={ws['B1'].value!r}）")
        cols = [c.name for c in load_workbook(rt)["集計"].tables["窓口一覧"].tableColumns]
        if cols[0] != "〇〇商事の窓口":
            problems.append(f"{estimate.name}: テーブルの列名が違います（{cols}）")
    for p in problems:
        print(f"NG: {p}", file=sys.stderr)
    print(f"LibreOffice で開けたファイル: {len(files) - len(problems)} / {len(files)}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1]), Path(sys.argv[2])))
