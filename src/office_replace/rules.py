"""置換のルール（置換前 → 置換後）. 置換前には * と ? のワイルドカードが使える.

- `*` は 0 文字以上、`?` は 1 文字（Excel の検索と同じ）。`~*` `~?` `~~` で記号そのもの
- `--regex` のときは置換前を正規表現として扱い、置換後で \\1 などのグループを使える
- 対象（targets）: text（セル・図形・本文など）/ sheet（シート名）/ file（ファイル名）。既定は text
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

TARGETS = ("text", "sheet", "file")
_TARGET_ALIASES = {
    "文字列": "text",
    "テキスト": "text",
    "シート名": "sheet",
    "シート": "sheet",
    "ファイル名": "file",
    "ファイル": "file",
    "すべて": "all",
    "all": "all",
}


def wildcard_to_regex(pattern: str) -> str:
    out, i = [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        out.append(".*?" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return "".join(out)


@dataclass
class Rule:
    before: str
    after: str
    targets: tuple[str, ...] = ("text",)
    regex: bool = False
    ignore_case: bool = False
    whole: bool = False  # セル・図形の文字列全体が一致するときだけ
    _re: re.Pattern[str] | None = field(default=None, repr=False)

    def compiled(self) -> re.Pattern[str]:
        if self._re is None:
            body = self.before if self.regex else wildcard_to_regex(self.before)
            if self.whole:
                body = rf"\A(?:{body})\Z"
            self._re = re.compile(body, re.I if self.ignore_case else 0)
        return self._re

    def apply(self, text: str) -> tuple[str, int]:
        rx = self.compiled()
        if self.regex:
            return rx.subn(self.after, text)
        return rx.subn(lambda _m: self.after, text)


def validate(rules: list[Rule]) -> list[str]:
    problems = []
    for i, r in enumerate(rules, start=1):
        if not r.before:
            problems.append(f"ルール {i}: 置換前が空です")
            continue
        if not r.regex and r.before.strip("*") == "":
            problems.append(f"ルール {i}: 置換前が * だけでは、すべての文字列が置換されます")
        try:
            rx = r.compiled()
        except re.error as e:
            problems.append(f"ルール {i}: 正規表現の誤り: {e}")
            continue
        if rx.search("") and not r.whole:
            problems.append(f"ルール {i}: 空文字に一致するパターンです（{r.before}）")
        bad = [t for t in r.targets if t not in TARGETS]
        if bad:
            problems.append(f"ルール {i}: 対象 {bad} は text / sheet / file のどれかです")
    return problems


def _targets(value: str) -> tuple[str, ...]:
    if not value:
        return ("text",)
    parts = [p.strip() for p in re.split(r"[,、/\s]+", value) if p.strip()]
    out: list[str] = []
    for p in parts:
        t = _TARGET_ALIASES.get(p, _TARGET_ALIASES.get(p.lower(), p.lower()))
        out += list(TARGETS) if t == "all" else [t]
    return tuple(dict.fromkeys(out))


def _row_rule(row: dict[str, str], *, regex: bool, ignore_case: bool) -> Rule | None:
    def get(*names: str) -> str:
        for n in names:
            if n in row and row[n] is not None:
                return str(row[n])
        return ""

    before = get("置換前", "before", "from", "検索")
    if not before.strip():
        return None
    return Rule(
        before=before,
        after=get("置換後", "after", "to", "置換"),
        targets=_targets(get("対象", "targets", "target")),
        regex=regex,
        ignore_case=ignore_case,
        whole=get("完全一致", "whole").strip().lower() in ("○", "yes", "true", "1", "y"),
    )


def load(path: Path, *, regex: bool = False, ignore_case: bool = False) -> list[Rule]:
    """ルールの一覧（.csv / .tsv / .xlsx / .yaml）. 列: 置換前・置換後・対象（任意）・完全一致（任意）."""
    suffix = path.suffix.lower()
    rows: list[dict[str, str]]
    if suffix in (".yaml", ".yml"):
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        rows = data.get("rules", data) if isinstance(data, dict) else data
    elif suffix in (".csv", ".tsv", ".txt"):
        text = None
        for enc in ("utf-8-sig", "cp932"):
            try:
                text = path.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError(f"{path.name}: 文字コードを判別できません（UTF-8 か Shift_JIS）")
        rows = list(csv.DictReader(text.splitlines(), delimiter="\t" if suffix != ".csv" else ","))
    elif suffix in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook

        ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
        it = ws.iter_rows(values_only=True)
        header = [str(h or "").strip() for h in next(it, [])]
        rows = [{h: ("" if v is None else str(v)) for h, v in zip(header, r, strict=False)} for r in it]
    else:
        raise ValueError(f"ルールの形式に対応していません（.csv / .tsv / .xlsx / .yaml）: {path.name}")
    rules = [r for r in (_row_rule(row, regex=regex, ignore_case=ignore_case) for row in rows) if r]
    if not rules:
        raise ValueError(f"{path.name}: ルールがありません（列名は 置換前・置換後）")
    return rules
