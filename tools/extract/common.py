"""Shared helpers for reading the Typfallsmodellen .xlsb workbook.

The extractors deliberately address cells the same way the VBA does -- 1-indexed
row and column, e.g. ``sheet.cell(year - 1959 + 6, 83)`` mirrors
``wsNyckelTal.Cells(year_(age) - 1959 + 6, 83)`` in ``VBA_go.bas``.  Keeping the
addressing identical is what makes the yearly diff against a new workbook
readable.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from pyxlsb import open_workbook

# Excel error codes surface through pyxlsb as small ints; the shipped workbook
# has its result sheets cleared, so #N/A (0x2a) and friends are common there.
_ERROR_CODES = {0x00, 0x07, 0x0F, 0x17, 0x1D, 0x24, 0x2A, 0x2B}


class Sheet:
    """A worksheet read fully into memory, addressed with 1-indexed (row, col)."""

    def __init__(self, name: str, rows: list[list[Any]]):
        self.name = name
        self._rows = rows

    @property
    def nrows(self) -> int:
        return len(self._rows)

    def cell(self, row: int, col: int) -> Any:
        """Value at 1-indexed (row, col); None when outside the used range."""
        if row < 1 or row > len(self._rows):
            return None
        r = self._rows[row - 1]
        if col < 1 or col > len(r):
            return None
        return r[col - 1]

    def num(self, row: int, col: int) -> float | None:
        """Numeric value at (row, col), or None for blanks, text and errors."""
        v = self.cell(row, col)
        if v is None or isinstance(v, (str, bool)):
            return None
        if isinstance(v, int) and v in _ERROR_CODES and not isinstance(v, bool):
            # pyxlsb reports Excel errors as bare ints. Genuine small integers do
            # occur, so only treat these as errors on sheets we know are cleared;
            # callers that need 0..43 pass through num_raw instead.
            return float(v)
        return float(v)

    def text(self, row: int, col: int) -> str:
        v = self.cell(row, col)
        if v is None or isinstance(v, (int, float)) and not isinstance(v, bool):
            return "" if v is None else str(v)
        return str(v).strip()

    def col_values(self, col: int, first_row: int, last_row: int) -> list[Any]:
        return [self.cell(r, col) for r in range(first_row, last_row + 1)]


class Workbook:
    """Lazily-read view of the workbook, keyed by sheet name."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self._cache: dict[str, Sheet] = {}
        with open_workbook(str(self.path)) as wb:
            self.sheet_names: list[str] = list(wb.sheets)

    def sheet(self, name: str) -> Sheet:
        if name in self._cache:
            return self._cache[name]
        if name not in self.sheet_names:
            raise KeyError(f"no sheet named {name!r}; have {self.sheet_names}")
        with open_workbook(str(self.path)) as wb:
            index = self.sheet_names.index(name) + 1
            with wb.get_sheet(index) as sh:
                rows: list[list[Any]] = []
                for row in sh.rows():
                    rows.append([c.v for c in row])
        sheet = Sheet(name, rows)
        self._cache[name] = sheet
        return sheet


def clean(value: Any) -> Any:
    """Normalise a cell value for JSON: drop NaN/inf, keep ints as ints."""
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        if value == int(value) and abs(value) < 2**53:
            return int(value)
        return value
    return value


def exact(value: float | None) -> float | None:
    """Pass a cell value through at full double precision.

    Deliberately does NOT round to a "tidy" number of significant digits. Several
    of these series are factors just above 1 -- inheritance gains at 1.0003,
    management-cost factors at 0.9997 -- and the engine uses them as ``x - 1``.
    That subtraction cancels the leading digits, so trimming the stored value to
    twelve significant figures puts a ~1e-8 relative error into the result, which
    is enough to miss the workbook's own figures.

    Python writes floats at shortest round-trip precision, so the JSON holds the
    exact double and is still byte-stable from one run to the next.
    """
    if value is None:
        return None
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        if value == int(value) and abs(value) < 2**53:
            return int(value)
    return value


def write_json(path: Path, payload: Any, *, description: str = "") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1, sort_keys=False)
        f.write("\n")
    size = path.stat().st_size
    label = f" ({description})" if description else ""
    print(f"  wrote {path.relative_to(REPO_ROOT)}{label} - {size:,} bytes")
    return path


def write_bytes(path: Path, payload: bytes, *, description: str = "") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    label = f" ({description})" if description else ""
    print(f"  wrote {path.relative_to(REPO_ROOT)}{label} - {len(payload):,} bytes")
    return path


REPO_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = REPO_ROOT / "packages" / "data"


@dataclass
class Series:
    """One yearly economic series plus the last year whose value is an actual."""

    name: str
    first_year: int
    values: list[float | None]
    last_actual_year: int
    source: str
    note: str = ""

    def to_json(self) -> dict[str, Any]:
        return {
            "firstYear": self.first_year,
            "lastActualYear": self.last_actual_year,
            "source": self.source,
            "note": self.note,
            "values": [clean(v) for v in self.values],
        }

    def year_range(self) -> Iterator[int]:
        return iter(range(self.first_year, self.first_year + len(self.values)))
