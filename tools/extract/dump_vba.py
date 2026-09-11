"""Dumps the VBA project to one file per module.

The dump is committed so that next year's workbook can be diffed against it:
a new release's rule changes show up as new procedures (avdrag27, Jobb27, ...)
in a readable diff, which is what turns the yearly update from an audit into a
review.  See docs/YEARLY-UPDATE.md.
"""

from __future__ import annotations

import re
from pathlib import Path

from oletools.olevba import VBA_Parser

# olevba emits a decompiled p-code listing alongside the source; it is large and
# adds nothing a diff can use.
PCODE_MODULE = re.compile(r"p-code", re.IGNORECASE)
SAFE_NAME = re.compile(r"[^\w.\-]+", re.UNICODE)


def extract(xlsb_path: Path, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("*.bas"):
        stale.unlink()
    for stale in out_dir.glob("*.cls"):
        stale.unlink()
    for stale in out_dir.glob("*.frm"):
        stale.unlink()

    parser = VBA_Parser(str(xlsb_path))
    modules = []
    try:
        for _, _, vba_filename, vba_code in parser.extract_macros():
            name = Path(vba_filename).name
            if not name or PCODE_MODULE.search(name):
                continue
            safe = SAFE_NAME.sub("_", name)
            path = out_dir / safe
            text = vba_code if isinstance(vba_code, str) else vba_code.decode("latin-1")
            path.write_text(text, encoding="utf-8")
            modules.append({"module": name, "lines": text.count("\n") + 1})
    finally:
        parser.close()

    modules.sort(key=lambda m: m["module"].lower())
    return {
        "moduleCount": len(modules),
        "totalLines": sum(m["lines"] for m in modules),
        "modules": modules,
    }
