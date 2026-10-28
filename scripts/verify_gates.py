#!/usr/bin/env python3
"""CI（.github/workflows/ci.yml）と同じ品質ゲートを、同じ順番で手元で実行する（push 前の確認用）.

    python scripts/verify_gates.py            # 全ゲート
    python scripts/verify_gates.py --quick    # lint・フォーマット・テストだけ
    python scripts/verify_gates.py --fix      # ruff の自動修正とフォーマットを先に適用してから確認

ゲート（ci.yml のジョブ・手順と同じ）:
  1. lint        ruff check .
  2. format      ruff format --check .
  3. test        pytest -q
  4. sample      サンプル（samples/input）を --dry-run と実際の書き出し（一時フォルダ）で置換する
                 （CI は先に samples/make_samples.py で作り直す. ここではコミット済みのサンプルを使い、作業ツリーを変えない）
  5. sast        bandit -q -r src -ll（未導入なら SKIP）
  6. deps        pip-audit --skip-editable（未導入なら SKIP）
  - LibreOffice で開けるかの確認（ci.yml の libreoffice ジョブ）は soffice が要るので CI に任せる

終了コード: 0 すべて PASS（SKIP を含む）/ 1 FAIL あり。
**ゲートを通すためにテストやチェックを甘くしない**（アサーションの削除・除外の追加・無効化はしない）。
"""

from __future__ import annotations

import argparse
import importlib.util
import subprocess  # nosec B404 — 固定のコマンドだけを実行する
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def gates(quick: bool, tmp: Path) -> list[tuple[str, list[list[str]] | None, str]]:
    py = sys.executable
    tool = [py, "-m", "office_replace"]
    out: list[tuple[str, list[list[str]] | None, str]] = [
        ("lint", [[py, "-m", "ruff", "check", "."]], ""),
        ("format", [[py, "-m", "ruff", "format", "--check", "."]], ""),
        ("test", [[py, "-m", "pytest", "-q"]], ""),
    ]
    if quick:
        return out
    out.append(
        (
            "sample",
            [
                [*tool, "samples/input", "--rules", "samples/rules.xlsx", "-r", "--dry-run"],
                [*tool, "samples/input", "--rules", "samples/rules.csv", "-r", "-o", str(tmp / "out")],
            ],
            "",
        )
    )
    out.append(
        ("sast", [[py, "-m", "bandit", "-q", "-r", "src", "-ll"]], "")
        if _module("bandit")
        else ("sast", None, "bandit 未導入（pip install bandit==1.8.6）")
    )
    out.append(
        ("deps", [[py, "-m", "pip_audit", "--skip-editable", "--progress-spinner", "off"]], "")
        if _module("pip_audit")
        else ("deps", None, "pip-audit 未導入（pip install pip-audit==2.9.0）")
    )
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--quick", action="store_true", help="lint・フォーマット・テストだけ")
    ap.add_argument("--fix", action="store_true", help="ruff の自動修正とフォーマットを先に適用する")
    args = ap.parse_args(argv)
    if not _module("ruff") or not _module("pytest"):
        print('ruff と pytest が必要です: pip install -c constraints.txt -e ".[dev]"（または ./setup.sh / setup.cmd）', file=sys.stderr)
        return 1
    if args.fix:
        subprocess.run([sys.executable, "-m", "ruff", "check", "--fix", "."], cwd=ROOT)  # nosec B603
        subprocess.run([sys.executable, "-m", "ruff", "format", "."], cwd=ROOT)  # nosec B603

    results: list[tuple[str, str, str]] = []
    with tempfile.TemporaryDirectory(prefix="office-replace-gates-") as d:
        for name, cmds, note in gates(args.quick, Path(d)):
            if cmds is None:
                results.append((name, "SKIP", note))
                continue
            rc = 0
            for cmd in cmds:
                print(f"\n=== {name}: {' '.join(cmd[1:])}", flush=True)
                rc = subprocess.run(cmd, cwd=ROOT).returncode  # nosec B603
                if rc != 0:
                    break
            results.append((name, "PASS" if rc == 0 else "FAIL", "" if rc == 0 else f"終了コード {rc}"))

    print("\n| ゲート | 結果 | 補足 |\n|---|---|---|")
    for name, status, note in results:
        print(f"| {name} | {status} | {note} |")
    failed = [n for n, s, _ in results if s == "FAIL"]
    if failed:
        print(f"\nFAIL: {', '.join(failed)}。原因を直してください（チェックを甘くして通さない）。")
        return 1
    print("\nすべてのゲートを通過しました（SKIP と LibreOffice の確認は CI で実行されます）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
