#!/usr/bin/env python3
"""Mechanically translates the grundavdrag and jobbskatteavdrag functions to TypeScript.

`Skatteregler.bas` holds 32 near-identical rule functions -- one per tax year --
that are nothing but thresholds and arithmetic: roughly 1 300 lines of numbers.
Typing those out by hand invites exactly the kind of transposed digit that no
test would catch, so they are translated mechanically instead and reviewed after.

The translated subset is deliberately tiny: If/ElseIf/Else, assignments, and
arithmetic. Anything outside it raises rather than guessing, so an unhandled
construct fails loudly instead of producing plausible wrong code.

Run once and commit the result; this is not part of the yearly pipeline. When a
new rule year appears, re-run it for the new function and paste that one in.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SOURCE = REPO_ROOT / "reference" / "vba" / "Skatteregler.bas"

FUNCTION_RE = re.compile(r"^\s*Function\s+(avdrag\d+|Jobb\d+)\s*\(", re.IGNORECASE)
END_RE = re.compile(r"^\s*End Function", re.IGNORECASE)


class Unsupported(Exception):
    """Raised for any VBA the translator does not explicitly handle."""


def strip_comment(line: str) -> str:
    """Removes a trailing VBA comment, respecting quoted strings."""
    out, in_string = [], False
    for ch in line:
        if ch == '"':
            in_string = not in_string
        if ch == "'" and not in_string:
            break
        out.append(ch)
    return "".join(out).rstrip()


def translate_expression(expr: str) -> str:
    """Translates a VBA arithmetic or boolean expression."""
    e = expr.strip()
    # VBA's type-suffixed literals: 5# is the Double 5.
    e = re.sub(r"(\d)#", r"\1", e)
    e = re.sub(r"\bInt\s*\(", "vbaInt(", e)
    e = re.sub(r"\bMax\s*\(", "Math.max(", e, flags=re.IGNORECASE)
    e = re.sub(r"\bAnd\b", "&&", e, flags=re.IGNORECASE)
    e = re.sub(r"\bOr\b", "||", e, flags=re.IGNORECASE)
    e = e.replace("<>", "!==")
    # Equality: only where it is a comparison, i.e. not the outer assignment.
    e = re.sub(r"(?<![<>!=])=(?!=)", "===", e)
    e = re.sub(r"\bavdragxx\s*\(", "avdragxx(", e)
    if "^" in e:
        raise Unsupported(f"exponentiation: {expr}")
    return e


def translate_condition(cond: str, result_name: str) -> str:
    """Translates a condition, rewriting the function's own name as `result`.

    VBA assigns to and reads the function name as if it were a variable, so
    `If Jobb14 < 0 Then Jobb14 = 0` reads it on both sides.
    """
    out = translate_expression(cond)
    return re.sub(rf"\b{re.escape(result_name)}\b", "result", out)


def translate_statement(stmt: str, result_name: str, integer_vars: set[str]) -> str:
    """Translates one simple VBA statement.

    `integer_vars` names variables declared As Long/Integer/Byte. VBA rounds on
    assignment to those -- banker's rounding, not truncation -- so every
    assignment to one is wrapped. Jobb07 and Jobb08 declare `jobb As Long`, which
    quietly rounds the jobbskatteavdrag to whole kronor at every step.
    """
    s = stmt.strip()

    m = re.match(r"^Dim\s+(\w+)\s+As\s+(\w+)$", s, re.IGNORECASE)
    if m:
        if m.group(2).lower() in {"long", "integer", "byte"}:
            integer_vars.add(m.group(1))
        return f"let {m.group(1)} = 0;"

    # `If <cond> Then <statement>` on one line.
    m = re.match(r"^If\s+(.+?)\s+Then\s+(.+)$", s, re.IGNORECASE)
    if m:
        inner = translate_statement(m.group(2), result_name, integer_vars)
        return f"if ({translate_condition(m.group(1), result_name)}) {inner}"

    m = re.match(r"^(\w+)\s*=\s*(.+)$", s)
    if m:
        target, value = m.group(1), translate_expression(m.group(2))
        value = re.sub(rf"\b{re.escape(result_name)}\b", "result", value)
        if target in integer_vars:
            value = f"vbaCLng({value})"
        if target == result_name:
            target = "result"
        return f"{target} = {value};"

    raise Unsupported(s)


def translate_signature(lines: list[str], name: str) -> tuple[str, list[str]]:
    """Turns the VBA signature into a TypeScript parameter list."""
    text = " ".join(line.rstrip(" _").strip() for line in lines)
    inside = text[text.index("(") + 1 : text.rindex(")")]

    params, names = [], []
    for raw in inside.split(","):
        part = raw.strip()
        optional = bool(re.match(r"^Optional\b", part, re.IGNORECASE))
        part = re.sub(r"^Optional\s+", "", part, flags=re.IGNORECASE)
        part = re.sub(r"^ByVal\s+", "", part, flags=re.IGNORECASE)
        part = re.sub(r"\s+As\s+\w+$", "", part)
        if "=" in part:
            pname, default = (p.strip() for p in part.split("=", 1))
            default = re.sub(r"(\d)#", r"\1", default)
            params.append(f"{pname} = {default}")
        else:
            pname = part.strip()
            params.append(f"{pname}{'?: number' if optional else ': number'}")
        names.append(pname)
    return ", ".join(params), names


def translate_function(name: str, signature: list[str], body: list[str]) -> str:
    params, param_names = translate_signature(signature, name)
    integer_vars: set[str] = set()
    out = [f"export function {name}({params}): number {{", "  let result = 0;"]
    indent = 1

    for raw in body:
        line = strip_comment(raw).strip()
        if not line:
            continue
        low = line.lower()

        if low.startswith("if ") and low.endswith(" then"):
            cond = translate_condition(line[3:-5], name)
            out.append("  " * indent + f"if ({cond}) {{")
            indent += 1
        elif low.startswith("elseif ") and low.endswith(" then"):
            indent -= 1
            cond = translate_condition(line[7:-5], name)
            out.append("  " * indent + f"}} else if ({cond}) {{")
            indent += 1
        elif low == "else":
            indent -= 1
            out.append("  " * indent + "} else {")
            indent += 1
        elif low == "end if":
            indent -= 1
            out.append("  " * indent + "}")
        else:
            out.append("  " * indent + translate_statement(line, name, integer_vars))

    out.append("  return result;")
    out.append("}")
    return "\n".join(out)


GENERATED_FILES = {
    "avdrag": REPO_ROOT / "packages" / "engine" / "src" / "skatt" / "grundavdrag.ts",
    "Jobb": REPO_ROOT / "packages" / "engine" / "src" / "skatt" / "jobbskatteavdrag.ts",
}


def check_against_committed(translated: dict[str, str]) -> int:
    """Verifies the committed TypeScript still matches the VBA it came from.

    Run by `npm run check:transpile`. It catches a hand-edit of the generated
    functions, and it catches the committed code drifting from a newly downloaded
    workbook -- either of which would quietly break parity.
    """
    problems = []
    for prefix, path in GENERATED_FILES.items():
        if not path.exists():
            problems.append(f"{path.name}: missing")
            continue
        committed = path.read_text(encoding="utf-8")
        for name, body in translated.items():
            if not name.startswith(prefix):
                continue
            if body not in committed:
                problems.append(f"{path.name}: {name} does not match the VBA")

    if problems:
        print("Transpiled tax functions have drifted from Skatteregler.bas:", file=sys.stderr)
        for problem in problems:
            print(f"  {problem}", file=sys.stderr)
        print("\nRe-run tools/transpile/skatteregler.py and review the diff.", file=sys.stderr)
        return 1

    print(f"{len(translated)} transpiled tax functions match Skatteregler.bas")
    return 0


def main() -> int:
    lines = SOURCE.read_text(encoding="utf-8").split("\n")
    translated: dict[str, str] = {}
    pieces: list[str] = []
    i = 0
    while i < len(lines):
        m = FUNCTION_RE.match(lines[i])
        if not m:
            i += 1
            continue
        name = m.group(1)
        signature = [lines[i]]
        while signature[-1].rstrip().endswith("_"):
            i += 1
            signature.append(lines[i])
        body, i = [], i + 1
        while i < len(lines) and not END_RE.match(lines[i]):
            body.append(lines[i])
            i += 1
        try:
            code = translate_function(name, signature, body)
            translated[name] = code
            pieces.append(code)
        except Unsupported as exc:
            print(f"error: {name}: unsupported construct: {exc}", file=sys.stderr)
            return 1
        i += 1

    if "--check" in sys.argv:
        return check_against_committed(translated)

    print(f"// {len(pieces)} functions translated from Skatteregler.bas\n")
    print("\n\n".join(pieces))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
