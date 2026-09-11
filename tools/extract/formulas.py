"""Recovers which cells hold literal values and which hold formulas.

The .xlsb stores formulas as parsed token streams, so reading them needs a
detour: LibreOffice converts the workbook to .xlsx once, and openpyxl then
reports each cell as either a constant or a formula string.

This is what makes ``lastActualYear`` exact instead of guessed.  In the workbook
a series' decided values are typed in as literals and the projection begins at
the first formula cell, so "last literal year" *is* "last actual year" -- no
heuristics, no reviewer judgement required for the common case.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl

CONVERT_TIMEOUT_SECONDS = 900


class FormulaMap:
    """Per-sheet lookup of whether a 1-indexed cell holds a formula."""

    def __init__(self, xlsx_path: Path):
        self._wb = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=True)
        self._sheets: dict[str, dict[tuple[int, int], str]] = {}

    def _sheet(self, name: str) -> dict[tuple[int, int], str]:
        if name not in self._sheets:
            ws = self._wb[name]
            cells: dict[tuple[int, int], str] = {}
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cells[(cell.row, cell.column)] = cell.value
            self._sheets[name] = cells
        return self._sheets[name]

    def is_formula(self, sheet: str, row: int, col: int) -> bool:
        return (row, col) in self._sheet(sheet)

    def formula(self, sheet: str, row: int, col: int) -> str | None:
        return self._sheet(sheet).get((row, col))

    def last_literal_row(self, sheet: str, col: int, first_row: int, last_row: int) -> int | None:
        """Highest row in the range whose cell in ``col`` is a literal value."""
        cells = self._sheet(sheet)
        for row in range(last_row, first_row - 1, -1):
            if (row, col) not in cells:
                return row
        return None


def convert_to_xlsx(xlsb_path: Path, cache_dir: Path | None = None) -> Path:
    """Convert the .xlsb to .xlsx with LibreOffice, reusing a cached result.

    The cache key is the source file's size and mtime, so re-running the
    extractor against the same workbook skips the (slow) conversion.
    """
    xlsb_path = Path(xlsb_path).resolve()
    stat = xlsb_path.stat()
    cache_dir = cache_dir or Path(tempfile.gettempdir()) / "typfallsmodellen-xlsx"
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached = cache_dir / f"{xlsb_path.stem}-{stat.st_size}-{int(stat.st_mtime)}.xlsx"
    if cached.exists():
        return cached

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise RuntimeError(
            "LibreOffice is required to read worksheet formulas. Install it with:\n"
            "  sudo apt-get install -y libreoffice-calc"
        )

    with tempfile.TemporaryDirectory() as tmp:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, str(xlsb_path)],
            capture_output=True,
            text=True,
            timeout=CONVERT_TIMEOUT_SECONDS,
        )
        produced = Path(tmp) / f"{xlsb_path.stem}.xlsx"
        if not produced.exists():
            raise RuntimeError(
                f"LibreOffice did not produce {produced.name}.\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )
        shutil.copy2(produced, cached)
    return cached
