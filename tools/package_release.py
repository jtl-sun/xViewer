from __future__ import annotations

import hashlib
import shutil
import zipfile
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from mdir import __version__

DIST = ROOT / "dist"
ARCHIVE_NAME = f"xViewer-{__version__}.zip"
EXCLUDED_DIRS = {
    ".git", ".venv", "venv", "env", "build", "dist",
    "bootstrap", "bootstrap267", "payload",
    "__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache",
}
EXCLUDED_SUFFIXES = {".pyc", ".pyo", ".log", ".tmp", ".bak", ".swp"}


def should_include(path: Path) -> bool:
    relative = path.relative_to(ROOT)
    if any(part in EXCLUDED_DIRS or part.endswith(".egg-info") for part in relative.parts):
        return False
    if path.suffix.lower() in EXCLUDED_SUFFIXES:
        return False
    return path.is_file()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> int:
    DIST.mkdir(exist_ok=True)
    archive = DIST / ARCHIVE_NAME
    if archive.exists():
        archive.unlink()

    prefix = Path(f"xViewer-{__version__}")
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for path in sorted(ROOT.rglob("*")):
            if should_include(path):
                zf.write(path, prefix / path.relative_to(ROOT))

    sums = DIST / "SHA256SUMS.txt"
    release_files = [archive]
    release_files.extend(sorted(DIST.glob("*.whl")))
    release_files.extend(sorted(DIST.glob("*.tar.gz")))
    sums.write_text(
        "".join(f"{sha256(path)}  {path.name}\n" for path in release_files),
        encoding="utf-8",
    )
    print(f"Created {archive}")
    print(f"SHA256 {sha256(archive)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
