"""フォルダ（またはファイル）を走査して置換し、別のフォルダに書き出す."""

from __future__ import annotations

import logging
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from . import leftover, processors
from .ooxml import Change, Package, UnsafeZipError, ZipLimits

__all__ = ["FileResult", "Options", "ZipLimits", "process_file", "run"]
from .rules import LONG_MATCH, Rule

log = logging.getLogger(__name__)

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
    warnings: list[str] = field(default_factory=list)  # 終了コードには影響しない注意（* の一致が長すぎる など）


@dataclass
class Options:
    recursive: bool = False
    dry_run: bool = False
    copy_all: bool = False
    include_masters: bool = False
    zip_limits: ZipLimits = field(default_factory=ZipLimits)  # zip 爆弾対策の上限（展開後の合計・部品の数・圧縮率）


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
        pkg = Package(path, opt.zip_limits)
    except zipfile.BadZipFile:
        result.status = "error"
        result.messages.append("開けません（パスワード付き・壊れている・拡張子と中身が違う、のいずれか）")
        return result
    except UnsafeZipError as e:
        result.status = "error"
        result.messages.append(f"安全のため処理しません（{e}）")
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
                    result.leftovers = leftover.scan_zip(zf, check, limits=opt.zip_limits)
            else:
                result.leftovers = leftover.scan_zip(pkg.zip, check, {n: pkg.xml(n) for n in pkg.dirty}, limits=opt.zip_limits)
    finally:
        pkg.close()
    return result


def _long_match_warnings(rules: list[Rule], before: list[int]) -> list[str]:
    """このファイルの処理中に増えた、* の長すぎる一致の警告. 置換前の文字列は入れない（レポートを伏せても漏れないように）."""
    out = []
    for i, (r, n) in enumerate(zip(rules, before, strict=True), start=1):
        new = r.long_matches[n:]
        if new:
            out.append(
                f"ルール {i} の * が {LONG_MATCH} 文字より長く一致しました（{len(new)} か所、最長 {max(new)} 文字）。"
                "意図しない範囲まで置換していないか、置換後のファイルを確認してください"
            )
    return out


def process_safely(path: Path, rel: Path, out_dir: Path, rules: list[Rule], opt: Options) -> FileResult:
    """1 ファイルを処理する. 想定外の例外でも全体を止めず、そのファイルをエラーとして記録する."""
    marks = [len(r.long_matches) for r in rules]
    try:
        result = process_file(path, rel, out_dir, rules, opt)
    except Exception as e:  # 1 ファイルの失敗で残りのファイルとレポートを失わないため
        log.debug("処理に失敗: %s", path, exc_info=True)
        result = FileResult(path, None, "error", messages=[f"処理中にエラーが起きました（{type(e).__name__}: {e}）"])
    result.warnings += _long_match_warnings(rules, marks)
    return result


def run(src: Path, out_dir: Path, rules: list[Rule], opt: Options) -> list[FileResult]:
    base = src if src.is_dir() else src.parent
    results: list[FileResult] = []
    seen: dict[Path, Path] = {}
    for path in collect(src, opt.recursive):
        rel = path.relative_to(base)
        try:
            dest = out_dir / rel.parent / new_name(path, rules)[0]
        except Exception as e:  # 置換後の \\1 の誤りなど. このファイルだけエラーにして続ける
            results.append(FileResult(path, None, "error", messages=[f"ファイル名を置換できません（{type(e).__name__}: {e}）"]))
            continue
        key = Path(str(dest).lower())  # Windows は大文字小文字を区別しない
        if key in seen:
            msg = f"出力先が {seen[key].name} と同じ名前になるため処理しません（ファイル名の置換を見直してください）"
            results.append(FileResult(path, None, "error", messages=[msg]))
            continue
        seen[key] = path
        results.append(process_safely(path, rel, out_dir, rules, opt))
    return results
