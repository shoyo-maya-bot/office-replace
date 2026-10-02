"""フォルダ（またはファイル）を走査して置換し、別のフォルダに書き出す."""

from __future__ import annotations

import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from . import processors
from .ooxml import Change, Package
from .rules import Rule

EXCEL = {".xlsx", ".xlsm"}
WORD = {".docx", ".docm"}
POWERPOINT = {".pptx", ".pptm"}
SUPPORTED = EXCEL | WORD | POWERPOINT
LEGACY = {".xls": ".xlsx", ".doc": ".docx", ".ppt": ".pptx"}


@dataclass
class FileResult:
    source: Path
    dest: Path | None
    status: str  # changed / unchanged / skipped / error
    changes: list[Change] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)


@dataclass
class Options:
    recursive: bool = False
    dry_run: bool = False
    copy_all: bool = False
    include_masters: bool = False


def _for(rules: list[Rule], target: str) -> list[Rule]:
    return [r for r in rules if target in r.targets]


def collect(src: Path, recursive: bool) -> list[Path]:
    if src.is_file():
        return [src]
    it = src.rglob("*") if recursive else src.glob("*")
    return sorted(p for p in it if p.is_file() and not p.name.startswith("~$"))


def new_name(path: Path, rules: list[Rule]) -> tuple[str, int]:
    stem, total = path.stem, 0
    for r in _for(rules, "file"):
        stem, n = r.apply(stem)
        total += n
    return stem + path.suffix, total


def process_file(path: Path, rel: Path, out_dir: Path, rules: list[Rule], opt: Options) -> FileResult:
    suffix = path.suffix.lower()
    name, renamed = new_name(path, rules)
    dest = out_dir / rel.parent / name
    if suffix in LEGACY:
        return FileResult(path, None, "skipped", messages=[f"古い形式（{suffix}）は対象外です。{LEGACY[suffix]} で保存し直してください"])
    if suffix not in SUPPORTED:
        if opt.copy_all and not opt.dry_run:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, dest)
        return FileResult(path, dest if opt.copy_all else None, "skipped", messages=["Office のファイルではありません"])
    result = FileResult(path, dest, "unchanged")
    if renamed:
        result.changes.append(Change(path.name, "file", "ファイル名", path.name, name, renamed))
    try:
        pkg = Package(path)
    except zipfile.BadZipFile:
        result.status = "error"
        result.messages.append("開けません（パスワード付き・壊れている・拡張子と中身が違う、のいずれか）")
        return result
    try:
        text_rules = _for(rules, "text")
        if suffix in EXCEL:
            ch, errs = processors.excel(pkg, text_rules, _for(rules, "sheet"))
            result.changes += ch
            result.messages += errs
        elif suffix in WORD:
            result.changes += processors.word(pkg, text_rules)
        else:
            result.changes += processors.powerpoint(pkg, text_rules, opt.include_masters)
        if result.changes:
            result.status = "changed"
        if result.messages:
            result.status = "error" if not result.changes else "changed"
        if not opt.dry_run and (result.changes or opt.copy_all):
            if pkg.dirty:
                pkg.save(dest)
            else:
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, dest)
        elif not result.changes:
            result.dest = None
    finally:
        pkg.close()
    return result


def run(src: Path, out_dir: Path, rules: list[Rule], opt: Options) -> list[FileResult]:
    base = src if src.is_dir() else src.parent
    results: list[FileResult] = []
    seen: dict[Path, Path] = {}
    for path in collect(src, opt.recursive):
        rel = path.relative_to(base)
        dest = out_dir / rel.parent / new_name(path, rules)[0]
        key = Path(str(dest).lower())  # Windows は大文字小文字を区別しない
        if key in seen:
            msg = f"出力先が {seen[key].name} と同じ名前になるため処理しません（ファイル名の置換を見直してください）"
            results.append(FileResult(path, None, "error", messages=[msg]))
            continue
        seen[key] = path
        results.append(process_file(path, rel, out_dir, rules, opt))
    return results
