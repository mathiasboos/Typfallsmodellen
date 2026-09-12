#!/usr/bin/env python3
"""Extract Socialstyrelsen's riksnorm tables out of Bidrag.bas.

``bistOld`` and ``bist25`` hold the social assistance norm as 113 ``Array(...)``
literals: three table families -- ``xn`` (cost per child by age band), ``vuxna``
(cost per adult, single and cohabiting) and ``Gn`` (shared household costs by
family size) -- with one row per income year from 1985 to the present.

That is roughly nine hundred numbers whose only structure is the year they
belong to, and next year's workbook adds one row to each. Typing them by hand
invites exactly the transposed digit no test would catch, so they are parsed out
of the VBA instead, the same way Skatteregler.bas is transpiled rather than
retyped.

    python tools/extract/extract_riksnorm.py            # regenerate the JSON
    python tools/extract/extract_riksnorm.py --check    # verify it still matches

The rules *around* the tables -- the deductions for medical and dental care in
1994-96, the indexation past 2025, the comparison against disposable income --
stay in the engine, hand-ported like the rest of Bidrag.bas. Only the numbers
come from here.
"""

from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path
from typing import Any

# Stdlib only, like tools/transpile/skatteregler.py: `npm run check:riksnorm`
# has to work from a plain checkout, without the .xlsb reading toolchain.
REPO_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = REPO_ROOT / "packages" / "data"

VBA_FILE = REPO_ROOT / "reference" / "vba" / "Bidrag.bas"
OUT_FILE = DATA_DIR / "riksnorm.json"

#: The functions whose tables we take, and the table variables in each.
FUNCTIONS = ("bistOld", "bist25")
TABLES = ("xn", "vuxna", "Gn")

#: `If year = 1985 Then`, `ElseIf year <= 1995 Then`, and so on.
_CONDITION = re.compile(r"^\s*(?:Else)?If\s+year\s*(<=|>=|<>|<|>|=)\s*(\d+)\s+Then\s*$", re.I)
_ELSE = re.compile(r"^\s*Else\s*(?:'.*)?$", re.I)
_ASSIGN = re.compile(r"^\s*([A-Za-z_]\w*)\s*=\s*Array\s*\(", re.I)
_FUNC_START = re.compile(r"^\s*(?:Public\s+|Private\s+)?Function\s+([A-Za-z_]\w*)\s*\(", re.I)
_FUNC_END = re.compile(r"^\s*End\s+Function\s*$", re.I)

_OPS = {"=": "eq", "<=": "le", ">=": "ge", "<": "lt", ">": "gt", "<>": "ne"}


class ParseError(RuntimeError):
    pass


def write_json(path: Path, payload: Any, *, description: str = "") -> Path:
    """Same shape and formatting as common.write_json, without its dependencies."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1, sort_keys=False)
        f.write("\n")
    label = f" ({description})" if description else ""
    print(f"  wrote {path.relative_to(REPO_ROOT)}{label} - {path.stat().st_size:,} bytes")
    return path


def _join_continuations(lines: list[str]) -> list[tuple[int, str]]:
    """Fold VBA's ``_`` line continuations into single logical lines.

    Returns (source line number of the first physical line, text) pairs, so an
    error can still point at the file.
    """
    out: list[tuple[int, str]] = []
    buffer = ""
    start = 0
    for number, raw in enumerate(lines, start=1):
        text = raw.rstrip("\r\n")
        if not buffer:
            start = number
        stripped = text.rstrip()
        if stripped.endswith("_") and (len(stripped) == 1 or stripped[-2].isspace()):
            buffer += stripped[:-1]
            continue
        out.append((start, buffer + text))
        buffer = ""
    if buffer:
        out.append((start, buffer))
    return out


def _strip_comment(text: str) -> str:
    """Drop a trailing VBA comment, respecting string literals."""
    in_string = False
    for index, char in enumerate(text):
        if char == '"':
            in_string = not in_string
        elif char == "'" and not in_string:
            return text[:index]
    return text


def _split_arguments(inner: str) -> list[str]:
    """Split an argument list on commas that are not inside parentheses."""
    parts: list[str] = []
    depth = 0
    current = ""
    for char in inner:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += char
    parts.append(current)
    return [p.strip() for p in parts if p.strip()]


_ALLOWED_NODES = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.USub,
    ast.UAdd,
    ast.Constant,
)


def _evaluate(expression: str) -> float:
    """Evaluate one array element.

    The entries are not all plain integers: several blend two levels over the
    months they applied (``1666 * 9 / 12 + 1320 * 3 / 12``), a few subtract a
    care charge inline (``3451 - 57``), and one is wrapped in ``CDbl``. So the
    expression is evaluated -- through a whitelist of arithmetic nodes, never
    ``eval`` on arbitrary text.
    """
    cleaned = re.sub(r"\bCDbl\s*\(", "(", expression, flags=re.I)
    cleaned = cleaned.replace("#", "")  # VBA's Double type suffix, as in `1#`
    try:
        tree = ast.parse(cleaned, mode="eval")
    except SyntaxError as error:
        raise ParseError(f"cannot parse {expression!r}") from error
    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED_NODES):
            raise ParseError(f"unsupported syntax in {expression!r}: {type(node).__name__}")
        if isinstance(node, ast.Constant) and not isinstance(node.value, (int, float)):
            raise ParseError(f"non-numeric constant in {expression!r}")
    return float(eval(compile(tree, "<riksnorm>", "eval")))  # noqa: S307 - whitelisted above


def _function_bodies(text: str) -> dict[str, list[tuple[int, str]]]:
    """The logical lines of each function we care about, keyed by name."""
    bodies: dict[str, list[tuple[int, str]]] = {}
    current: str | None = None
    for number, line in _join_continuations(text.splitlines(keepends=True)):
        start = _FUNC_START.match(line)
        if start:
            current = start.group(1) if start.group(1) in FUNCTIONS else None
            if current:
                bodies[current] = []
            continue
        if _FUNC_END.match(line):
            current = None
            continue
        if current:
            bodies[current].append((number, line))
    missing = [name for name in FUNCTIONS if name not in bodies]
    if missing:
        raise ParseError(f"functions not found in {VBA_FILE.name}: {', '.join(missing)}")
    return bodies


def _branches(body: list[tuple[int, str]]) -> dict[str, list[dict]]:
    """Every ``<table> = Array(...)`` in one function, with its year condition.

    Each branch is recorded in source order with the comparison that guards it,
    so the engine can walk the same If/ElseIf chain the VBA does rather than
    relying on a year being a dictionary key.
    """
    tables: dict[str, list[dict]] = {}
    pending: dict | None = None

    for number, raw in body:
        line = _strip_comment(raw)
        if not line.strip():
            continue

        condition = _CONDITION.match(line)
        if condition:
            pending = {"op": _OPS[condition.group(1)], "year": int(condition.group(2))}
            continue
        if _ELSE.match(line):
            pending = {"op": "else"}
            continue

        assign = _ASSIGN.match(line)
        if not assign:
            # Any other statement ends the run of "this branch opens with a
            # table assignment", so a later Array( cannot borrow a stale
            # condition from far above.
            if not line.strip().lower().startswith(("if ", "elseif ", "end if")):
                pending = None
            continue

        name = assign.group(1)
        if name not in TABLES:
            pending = None
            continue
        if pending is None:
            raise ParseError(f"{VBA_FILE.name}:{number}: `{name} = Array(` with no year condition")

        inner = line[line.index("(", assign.end(1)) + 1 : line.rindex(")")]
        values = [_evaluate(part) for part in _split_arguments(inner)]
        tables.setdefault(name, []).append({**pending, "values": values})
        pending = None

    missing = [name for name in TABLES if name not in tables]
    if missing:
        raise ParseError(f"no branches found for: {', '.join(missing)}")
    return tables


def build() -> dict:
    text = VBA_FILE.read_text(encoding="utf-8")
    bodies = _function_bodies(text)
    tables = {name: _branches(bodies[name]) for name in FUNCTIONS}

    for function, families in tables.items():
        for family, branches in families.items():
            widths = {len(branch["values"]) for branch in branches}
            if len(widths) != 1:
                raise ParseError(f"{function}.{family}: rows of differing width {sorted(widths)}")
            if branches[-1]["op"] != "else":
                raise ParseError(f"{function}.{family}: chain does not end in Else")

    return {
        "source": str(VBA_FILE.relative_to(REPO_ROOT)),
        "note": (
            "Socialstyrelsen's riksnorm, parsed out of the VBA rather than retyped. "
            "Branches are in source order and carry the comparison that guards them, so "
            "the engine walks the same If/ElseIf chain. Regenerate with "
            "tools/extract/extract_riksnorm.py; verify with `npm run check:riksnorm`."
        ),
        "families": {
            "xn": "cost per child, by the eight age bands 0, 1-2, 3, 4-6, 7-10, 11-14, 15-18, 19-20",
            "vuxna": "cost per adult: single, then cohabiting",
            "Gn": "shared household costs, for 1 to 7 family members",
        },
        "tables": tables,
    }


def extract() -> dict:
    """Entry point for run.py."""
    payload = build()
    write_json(OUT_FILE, payload, description="riksnorm tables")
    return payload


def check() -> int:
    """Re-parse the VBA and verify the committed JSON still matches.

    Catches a hand-edit of the generated file, and catches the committed data
    drifting from a newly downloaded workbook.
    """
    if not OUT_FILE.exists():
        print(f"error: {OUT_FILE.relative_to(REPO_ROOT)} is missing; run without --check", file=sys.stderr)
        return 1
    expected = build()
    actual = json.loads(OUT_FILE.read_text(encoding="utf-8"))
    if actual == expected:
        rows = sum(len(b) for f in expected["tables"].values() for b in f.values())
        print(f"riksnorm: {rows} rows match {VBA_FILE.relative_to(REPO_ROOT)}")
        return 0

    print(f"error: {OUT_FILE.relative_to(REPO_ROOT)} does not match the VBA", file=sys.stderr)
    for function in FUNCTIONS:
        for family in TABLES:
            want = expected["tables"][function][family]
            got = actual.get("tables", {}).get(function, {}).get(family)
            if want != got:
                print(f"  {function}.{family} differs", file=sys.stderr)
    print("  regenerate with: python tools/extract/extract_riksnorm.py", file=sys.stderr)
    return 1


def main() -> int:
    if "--check" in sys.argv:
        return check()
    extract()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
