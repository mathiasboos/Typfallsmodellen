#!/usr/bin/env python3
"""Checks the workbook addresses the golden-file export depends on.

`reference/golden/ExportGoldenCases.bas` fills the Mikrosim sheet and reads the
results back, and the comparison harness decodes those results by position. Both
rest on where things sit in the workbook:

    rngXTopleft   B7    inputs start here; data from the row below
    rngYtopleft   M7    results start here
    rngTopXYOutput D28  Table 1 column D, the adjusted column, row for IP

A moved column in next year's release would not fail loudly -- it would produce
a CSV whose columns mean something other than the harness thinks. So the
addresses are asserted here.

Reads the defined names out of the .xlsb directly (BrtName records in
xl/workbook.bin) with nothing but the standard library, so it runs on a clean
checkout without the extraction venv.
"""

from __future__ import annotations

import argparse
import struct
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

BRT_NAME = 39
PTG_REF3D = 0x1A

# What the export macro and the harness assume, and where each assumption is
# spent if it turns out to be wrong.
EXPECTED = {
    "rngXTopleft": ("B7", "the input block: ExportGoldenCases writes the ten input columns here"),
    "rngYtopleft": ("M7", "the result block: the twelve output columns the harness decodes"),
    "rngExecuteFromRow": ("P3", "the batch runner's first row"),
    "rngExecuteUntilRow": ("U3", "the batch runner's last row"),
    "rngTopXYOutput": ("D28", "Table 1 column D, which every exported value is read from"),
    "rng_Top_tabell1": ("C24", "Table 1's first column, so column D is C + 1"),
    "rng_Tabell1_IP": ("A28", "row offset 1 from rngTopXYOutput"),
    "rng_Tabell1_Tot_Brutto": ("A36", "row offset 9: the exported 'Total pension'"),
    "rng_Tabell1_Efterskatt": ("A42", "row offset 15"),
    "rng_Tabell1_Bidrag": ("A43", "row offset 16"),
    "rng_Tabell1_Disp_efterskatt": ("A45", "row offset 18"),
}


def records(buf: bytes):
    """Walks BIFF12's variable-width record stream: id, size, payload."""
    i, n = 0, len(buf)
    while i < n:
        b = buf[i]
        i += 1
        rid = b & 0x7F
        if b & 0x80:
            rid |= (buf[i] & 0x7F) << 7
            i += 1
        size = 0
        for shift in range(4):
            b = buf[i]
            i += 1
            size |= (b & 0x7F) << (7 * shift)
            if not b & 0x80:
                break
        yield rid, buf[i : i + size]
        i += size


def column_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def defined_names(xlsb: Path) -> dict[str, str]:
    """Every defined name that is a plain single-cell reference, as A1 text."""
    with zipfile.ZipFile(xlsb) as z:
        workbook = z.read("xl/workbook.bin")

    found: dict[str, str] = {}
    for rid, payload in records(workbook):
        if rid != BRT_NAME:
            continue
        try:
            # flags(4) chKey(1) itab(4), then the name, then the formula.
            pos = 9
            cch = struct.unpack_from("<I", payload, pos)[0]
            pos += 4
            name = payload[pos : pos + 2 * cch].decode("utf-16-le")
            pos += 2 * cch
            cce = struct.unpack_from("<I", payload, pos)[0]
            pos += 4
            rgce = payload[pos : pos + cce]
        except (struct.error, UnicodeDecodeError):
            continue

        if not rgce or (rgce[0] & 0x1F) != PTG_REF3D or len(rgce) < 9:
            continue
        row, colflags = struct.unpack_from("<IH", rgce, 3)
        found[name] = f"{column_letter((colflags & 0x3FFF) + 1)}{row + 1}"
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook", nargs="?", default=str(REPO_ROOT / "source" / "Typfallsmodellen.xlsb"))
    parser.add_argument("--check", action="store_true", help="accepted for symmetry; this script only checks")
    parser.add_argument("--list", action="store_true", help="print every single-cell defined name")
    args = parser.parse_args()

    xlsb = Path(args.workbook).resolve()
    if not xlsb.exists():
        print(f"error: no workbook at {xlsb}", file=sys.stderr)
        return 1

    names = defined_names(xlsb)

    if args.list:
        for name in sorted(names):
            print(f"{name:34} {names[name]}")
        return 0

    problems = []
    for name, (want, why) in EXPECTED.items():
        got = names.get(name)
        if got is None:
            problems.append(f"{name} is not defined (expected {want}) -- {why}")
        elif got != want:
            problems.append(f"{name} is {got}, expected {want} -- {why}")

    if problems:
        print("the workbook's layout has moved:", file=sys.stderr)
        for problem in problems:
            print(f"  {problem}", file=sys.stderr)
        print(
            "\nreference/golden/ExportGoldenCases.bas resolves the sheet geometry at runtime, "
            "\nbut packages/engine/test/golden/compare.ts decodes the columns by position and "
            "\nwould need updating. See docs/VBA-MAPPING.md.",
            file=sys.stderr,
        )
        return 1

    print(f"workbook layout: {len(EXPECTED)} defined names are where the golden-file export expects")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
