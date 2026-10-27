#!/usr/bin/env bash
# セットアップ（macOS / Linux / Git Bash）: .venv を作って office-replace を入れる
set -euo pipefail
cd "$(dirname "$0")"

# Python 3.12 以上（pyproject.toml の requires-python と同じ）を探す. Git Bash では python3 が無く python か py のことがある
PY=""
for c in python3 python "py -3"; do
  if $c -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)' >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "Python 3.12 以上が見つかりません（python3 / python / py -3 を探しました）: https://www.python.org/downloads/" >&2
  exit 1
fi

[ -d .venv ] || $PY -m venv .venv
# 仮想環境の python: macOS・Linux は .venv/bin、Windows（Git Bash）は .venv/Scripts
if [ -x .venv/bin/python ]; then
  VPY=.venv/bin/python; ACT="source .venv/bin/activate"
else
  VPY=.venv/Scripts/python; ACT="source .venv/Scripts/activate"
fi
"$VPY" -m pip install -q --upgrade pip || true
# 依存の版は constraints.txt でそろえる（CI と同じ. deps-lock が作る）
C=""; [ -f constraints.txt ] && C="-c constraints.txt"
"$VPY" -m pip install -q $C -e ".[dev]"
echo "完了: $ACT のあと office-replace --help"
