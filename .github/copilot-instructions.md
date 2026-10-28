# Copilot への指示（office-replace）

Office ファイル（Excel / Word / PowerPoint）の一括置換ツールです。説明は日本語で行ってください。
詳しい約束は [AGENTS.md](../AGENTS.md)、使い方は [README.md](../README.md) にあります。

## 使うとき

- 置換を実行する前に、必ず `--dry-run` の結果をユーザーに見せて確認を取る（手順は `.github/prompts/office-replace.prompt.md`）
- 元のファイルは書き換えない（`-o` の出力フォルダに書き出す。入力フォルダの中は不可）。この方針を崩す提案をしない
- 終了コード: 0 正常 / 1 エラーのあったファイルあり、または置換前の文字列が残っている（`--allow-leftover` で 0）/ 2 引数・置換表の誤り / 3 実行時エラー
- 終了コード 1 のときは、レポートの「残っている文字列」の理由と対処（`--target formula` / `--target sheet` / `--target file` など）を説明する。
  フォルダ名に残ったものはこのツールでは直せない
- `--to` / `--target` / `--whole` は `--from` と組で使う。`--rules` のときは置換表の列（置換後・対象・完全一致）で指定する
- 置換表に社名・個人名が入ることがある。置換表やレポートの中身をコミット・他の場所へ書き写さない。共有するレポートは `--mask-report`

## 変更するとき

- XML は lxml で扱い、変更していない部品はバイト単位でそのまま写す（名前空間の接頭辞を保つ）
- ワイルドカード（`rules.py`）: `*` は最長一致だが、改行・タブと同じ置換前の 2 か所目をまたがない。README の「置換表」と食い違わせない
- オプション・終了コード・既定値・Python の版を変えたら README と `cli.py` の docstring も直す（`tests/test_docs_consistency.py` が確かめる）
- CHANGELOG の「未リリース」に書いたら、`pyproject.toml` の版を `.devN` にする
- `vba/` は未検証の参考実装。Python 版の仕様を変えたら vba/README.md の比較表も直す
- バグ修正には再現テストを添える。テストのデータは架空の名前にする
- コードを変更したら `python scripts/verify_gates.py`（CI と同じゲート）を実行して結果を報告する。
  急ぐときは `--quick`（`ruff check . && ruff format --check . && pytest` と同じ）
