"""ワイルドカード * の一致が Excel と違うところ（README に書いた仕様. 変えない）."""

from __future__ import annotations

from office_replace.rules import Rule


def test_star_does_not_cross_the_next_head():
    # Excel は「AAB」の全体に当たるが、このツールは * が置換前の頭（A）をまたがないので「AB」に当たる
    assert Rule("A*B", "X").apply("AAB") == ("AX", 1)
    assert Rule("A*B", "X").apply("ACB") == ("X", 1)
