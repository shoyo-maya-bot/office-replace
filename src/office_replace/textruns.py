"""段落の中で複数の「ラン」に分かれた文字列を置換する.

Word / PowerPoint / Excel のリッチテキストでは、見た目は 1 続きの文字列でも、書式の切れ目で別々の要素（ラン）に
分かれている。段落の文字列をつないで一致を探し、結果を元のランに書き戻す。
置換後の文字列は一致の始まるランに入り、その書式になる。一致の外の文字と書式は変わらない。

段落の中の改行・タブ（Word の w:br・w:tab、PowerPoint の a:br）は "\n"・"\t" の 1 文字のランとして並びに入れ、
fixed（書き換えないランの番号）で渡す。一致は改行・タブをまたがない（`*` も、置換前の普通の文字も）.
"""

from __future__ import annotations

import re

from .rules import Rule


def _matches(full: str, rule: Rule, cuts: list[tuple[int, int]]) -> list[re.Match[str]]:
    """一致の一覧. cuts（書き換えないランの範囲）をまたぐ一致は使わない."""
    rx = rule.compiled()
    if not cuts:
        return [m for m in rx.finditer(full) if m.end() > m.start()]
    if rule.whole:  # 全体一致は段落の全体で見る（改行・タブを含む段落の一部には一致させない）
        return [m for m in rx.finditer(full) if m.end() > m.start() and not any(m.start() < e and s < m.end() for s, e in cuts)]
    out: list[re.Match[str]] = []
    start = 0
    for s, e in [*cuts, (len(full), len(full))]:
        if s > start:  # 改行・タブの間の区間ごとに探す（* がまたがない）
            out += [m for m in rx.finditer(full, start, s) if m.end() > m.start()]
        start = e
    return out


def replace_runs(pieces: list[str], rule: Rule, fixed: frozenset[int] | set[int] = frozenset()) -> tuple[list[str], int]:
    """ランの文字列の並び → (新しい並び, 置換の件数). 並びの長さは変えない（空になるランはある）.

    fixed: 書き換えないラン（改行・タブ）の番号. 一致はこれをまたがない.
    """
    texts = list(pieces)
    full = "".join(texts)
    cuts, acc = [], 0
    for k, t in enumerate(texts):
        if k in fixed and t:
            cuts.append((acc, acc + len(t)))
        acc += len(t)
    matches = _matches(full, rule, cuts)
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


def replace_all(pieces: list[str], rules: list[Rule], fixed: frozenset[int] | set[int] = frozenset()) -> tuple[list[str], int]:
    total = 0
    for r in rules:
        pieces, n = replace_runs(pieces, r, fixed)
        total += n
    return pieces, total
