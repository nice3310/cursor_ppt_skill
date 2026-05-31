"""Produce templates/default.pptx — the fallback template used when no
company template is supplied in outline.meta.template.

This is intentionally minimal: 16:9 widescreen, the standard python-pptx
layouts (Title, Title and Content, Section Header, Two Content, Comparison,
Title Only, Blank, Content with Caption, Picture with Caption). The renderers
in lib/renderers.py draw most content themselves anyway; this template's
job is to provide a sane master + slide size + the layout name set that
pick_layout() expects.

Run once after install:
    python scripts/build_default_template.py
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches


SCRIPT_DIR = Path(__file__).parent
OUT = SCRIPT_DIR.parent / "templates" / "default.pptx"


def main() -> int:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"wrote {OUT}")
    print(f"layouts available ({len(prs.slide_layouts)}):")
    for i, layout in enumerate(prs.slide_layouts):
        print(f"  [{i}] {layout.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
