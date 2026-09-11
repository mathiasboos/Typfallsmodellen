"""Prose content: the glossary, the concept explanations, and the changelog.

Ordlista is a term followed by one or more paragraphs, with single-letter rows
acting as alphabet separators.  Förklaring is a flat term/definition table.
Versionsinformation is a version heading followed by its bullet points.
"""

from __future__ import annotations

import re

ALPHABET_ROW = re.compile(r"^\s*[A-ZÅÄÖ]\s*$")
VERSION_HEADING = re.compile(r"^Version\s", re.IGNORECASE)


def _glossary(sheet) -> list[dict]:
    entries: list[dict] = []
    current: dict | None = None
    for r in range(4, sheet.nrows + 1):
        text = sheet.text(r, 1)
        if not text or ALPHABET_ROW.match(text):
            continue
        # A term is a short heading line; the paragraphs that follow describe it.
        is_heading = len(text) < 60 and not text.endswith(".")
        if is_heading and (current is None or current["body"]):
            current = {"term": text, "body": []}
            entries.append(current)
        elif current is not None:
            current["body"].append(text)
    return [{"term": e["term"], "body": e["body"]} for e in entries if e["body"]]


def _explanations(sheet) -> list[dict]:
    out = []
    for r in range(2, sheet.nrows + 1):
        term, body = sheet.text(r, 1), sheet.text(r, 2)
        if term and body:
            out.append({"term": term, "body": body})
    return out


def _changelog(sheet) -> list[dict]:
    versions: list[dict] = []
    current: dict | None = None
    for r in range(1, sheet.nrows + 1):
        heading, bullet = sheet.text(r, 1), sheet.text(r, 2)
        if heading and VERSION_HEADING.match(heading):
            current = {"version": heading, "changes": []}
            versions.append(current)
        elif bullet and current is not None:
            current["changes"].append(bullet)
    return versions


def extract(wb) -> dict:
    changelog = _changelog(wb.sheet("Versionsinformation"))
    return {
        "source": "Ordlista, Förklaring and Versionsinformation",
        "modelVersion": changelog[0]["version"] if changelog else None,
        "glossary": _glossary(wb.sheet("Ordlista")),
        "explanations": _explanations(wb.sheet("Förklaring")),
        "changelog": changelog,
    }
