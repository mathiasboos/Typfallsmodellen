"""Swedish and English UI strings, from the workbook's own translation tables.

The model already ships every label in both languages, so the website's i18n is
extracted rather than re-translated -- which also keeps the wording identical to
the Excel version users may be migrating from.

  SysLang        numbered UI strings, keyed by the NR column
  SysMsgboxLang  message-box texts, keyed by Nr
  Sysdata        button and shape captions, keyed by named range
"""

from __future__ import annotations

LANGUAGES = {"sv": 2, "en": 3}  # column holding each language


def _numbered_table(sheet, first_row: int, key_col: int) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for r in range(first_row, sheet.nrows + 1):
        key = sheet.num(r, key_col)
        if key is None:
            continue
        entry = {lang: sheet.text(r, col) for lang, col in LANGUAGES.items()}
        if any(entry.values()):
            out[str(int(key))] = entry
    return out


def _keyed_table(sheet, first_row: int, key_col: int) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for r in range(first_row, sheet.nrows + 1):
        key = sheet.text(r, key_col)
        if not key:
            continue
        entry = {lang: sheet.text(r, col + 1) for lang, col in LANGUAGES.items()}
        if any(entry.values()):
            out[key] = entry
    return out


def extract(wb) -> dict:
    labels = _numbered_table(wb.sheet("SysLang"), 2, 1)
    messages = _numbered_table(wb.sheet("SysMsgboxLang"), 2, 1)
    captions = _keyed_table(wb.sheet("Sysdata"), 4, 1)
    return {
        "source": "SysLang, SysMsgboxLang and Sysdata",
        "languages": list(LANGUAGES),
        "labels": labels,
        "messages": messages,
        "captions": captions,
        "counts": {"labels": len(labels), "messages": len(messages), "captions": len(captions)},
    }
