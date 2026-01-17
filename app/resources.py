from pathlib import Path
import sys


# Resolve base path for resources in dev and PyInstaller bundles.
def _base_path() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

# Return a Path to a packaged resource.
def resource_path(*relative_parts: str) -> Path:
    return _base_path().joinpath(*relative_parts)

# Read a text resource from the packaged resources.
def read_text_resource(*relative_parts: str, encoding: str = "utf-8") -> str:
    return resource_path(*relative_parts).read_text(encoding=encoding)
