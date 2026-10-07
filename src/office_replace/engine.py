"""フォルダ（またはファイル）を走査して置換し、別のフォルダに書き出す."""

from __future__ import annotations

import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from . import leftover, processors
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
    leftovers: list[leftover.Leftover] = field(default_factory=list)  # 置換しきれずに残った置換前の文字列（注意）


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


_FORBIDDEN = set('/\\:*?"<>|') | {chr(c) for c in range(32)}


def _bad_name(name: str, suffix: str) -> str:
    """ファイル名の置換の結果が、出力フォルダの外や別のフォルダを指したり、Windows で作れない名前にならないか."""
    stem = name[: len(name) - len(suffix)] if suffix else name
    if any(c in _FORBIDDEN for c in name):
        return 'フォルダの区切りや Windows で使えない文字（\\ / : * ? " < > |）が入っています'
    if stem.strip(" .") == "":
        return "空、または「.」だけの名前です"
    if stem.rstrip() != stem or stem.endswith("."):
        return "末尾が空白か「.」です（Windows で作れません）"
    return ""


def process_file(path: Path, rel: Path, out_dir: Path, rules: list[Rule], opt: Options) -> FileResult:
    suffix = path.suffix.lower()
    name, renamed = new_name(path, rules)
    bad = _bad_name(name, path.suffix) if renamed else ""
    if bad:
        return FileResult(
            path, None, "error", messages=[f"置換後のファイル名「{name}」は使えません（{bad}）. 置換表の「置換後」を見直してください"]
        )
    dest = out_dir / rel.parent / name
    if not dest.resolve().is_relative_to(out_dir.resolve()):
        return FileResult(path, None, "error", messages=[f"出力先が出力フォルダの外になります（{name}）"])
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
            ch, errs = processors.excel(pkg, text_rules, _for(rules, "sheet"), _for(rules, "formula"))
            result.changes += ch
            result.messages += errs
        elif suffix in WORD:
            result.changes += processors.word(pkg, text_rules)
        else:
            result.changes += processors.powerpoint(pkg, text_rules, opt.include_masters)
        if text_rules:
            result.changes += processors.chart_caches(pkg, text_rules)
            result.changes += processors.properties(pkg, text_rules, titles=suffix not in EXCEL)
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
        # 置換前の文字列が残っていないか（書き出したときは書き出したファイルを、それ以外は置換後の内容を調べる）
        check = leftover.scan_rules(rules)
        if check:
            if not opt.dry_run and result.dest is not None and result.dest.exists():
                with zipfile.ZipFile(result.dest) as zf:
                    result.leftovers = leftover.scan_zip(zf, check)
            else:
                result.leftovers = leftover.scan_zip(pkg.zip, check, {n: pkg.xml(n) for n in pkg.dirty})
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
