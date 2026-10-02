"""Excel のシート名の制約."""

from __future__ import annotations

FORBIDDEN = set(":\\/?*[]")
MAX_LEN = 31
RESERVED = {"history"}  # Excel が予約している名前（大文字小文字を区別しない）


def check_sheet_name(name: str, others: list[str]) -> str:
    """問題があれば理由を返す（無ければ空文字）."""
    if not name.strip():
        return "空のシート名はつけられません"
    if len(name) > MAX_LEN:
        return f"{MAX_LEN} 文字を超えます（{len(name)} 文字）"
    bad = sorted({ch for ch in name if ch in FORBIDDEN})
    if bad:
        return f"使えない文字 {' '.join(bad)} を含みます"
    if name.startswith("'") or name.endswith("'"):
        return "先頭・末尾にアポストロフィ（'）は使えません"
    if name.lower() in RESERVED:
        return f"「{name}」は Excel の予約語です"
    if name.lower() in {o.lower() for o in others}:
        return "同じブックに同じ名前のシートがあります"
    return ""
