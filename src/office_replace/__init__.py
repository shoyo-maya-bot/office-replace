"""office-replace — Excel / Word / PowerPoint の文字列・図形・シート名・ファイル名を一括置換する."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("office-replace")
except PackageNotFoundError:
    __version__ = "0.0.0+local"
