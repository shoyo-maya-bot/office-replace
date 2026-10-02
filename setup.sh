#!/usr/bin/env bash
# セットアップ（macOS / Linux / Git Bash）: .venv を作って office-replace を入れる
set -euo pipefail
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/python -m pip install -q --upgrade pip || true
.venv/bin/python -m pip install -q -e ".[dev]"
echo "完了: source .venv/bin/activate のあと office-replace --help"
