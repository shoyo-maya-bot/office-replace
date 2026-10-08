"""段落の中で複数の「ラン」に分かれた文字列を置換する.

Word / PowerPoint / Excel のリッチテキストでは、見た目は 1 続きの文字列でも、書式の切れ目で別々の要素（ラン）に
分かれている。段落の文字列をつないで一致を探し、結果を元のランに書き戻す。
置換後の文字列は一致の始まるランに入り、その書式になる。一致の外の文字と書式は変わらない。
"""

from __future__ import annotations

from .rules import Rule


def replace_runs(pieces: list[str], rule: Rule) -> tuple[list[str], int]:
    """ランの文字列の並び → (新しい並び, 置換の件数). 並びの長さは変えない（空になるランはある）."""
    texts = list(pieces)
    full = "".join(texts)
    matches = [m for m in rule.compiled().finditer(full) if m.end() > m.start()]
    if not matches:
        return texts, 0
    for m in matches:
        rule.note(m)
    for m in reversed(matches):  # 後ろから直すと、前のランの位置がずれない
        s, e = m.start(), m.end()
        rep = m.expand(rule.after) if rule.regex else rule.after
        starts, acc = [], 0
        for t in texts:
            starts.append(acc)
            acc += len(t)
        i = next(k for k, st in enumerate(starts) if st <= s < st + len(texts[k]))
        j = next(k for k, st in enumerate(starts) if st < e <= st + len(texts[k]))
        if i == j:
            texts[i] = texts[i][: s - starts[i]] + rep + texts[i][e - starts[i] :]
            continue
        texts[i] = texts[i][: s - starts[i]] + rep
        for k in range(i + 1, j):
            texts[k] = ""
        texts[j] = texts[j][e - starts[j] :]
    return texts, len(matches)


def replace_all(pieces: list[str], rules: list[Rule]) -> tuple[list[str], int]:
    total = 0
    for r in rules:
        pieces, n = replace_runs(pieces, r)
        total += n
    return pieces, total
