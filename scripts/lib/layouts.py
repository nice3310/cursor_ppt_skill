"""Map a slide `type` to a concrete python-pptx slide_layout.

When a template is in use, we prefer layouts whose name matches a heuristic
hint (e.g. "Title and Content" for `bullets`). When the heuristic fails,
or when no template is given, we fall back to a "blank" layout and let
renderers.py draw shapes manually.

This file is intentionally dumb: no global state, no caches. The caller
passes the presentation in.
"""

from __future__ import annotations

from pptx.presentation import Presentation as PresentationT
from pptx.slide import SlideLayout


# Slide-type -> ordered list of layout-name substrings to look for.
# Lowercase comparison; first hit wins.
LAYOUT_HINTS: dict[str, list[str]] = {
    "title": ["title slide", "title"],
    "section": ["section header", "section"],
    "statement": ["title only", "centered", "big idea", "blank"],
    "bullets": ["title and content", "content"],
    "comparison": ["comparison", "two content"],
    "code": ["title and content", "content"],
    "diagram": ["title and content", "picture", "content"],
    "image": ["picture with caption", "picture", "title and content"],
    "table": ["title and content", "content"],
    "quote": ["title only", "blank"],
    "summary": ["title and content", "content"],
    "qa": ["title only", "section header", "blank"],
}

BLANK_FALLBACK_HINTS = ["blank", "title only"]


def pick_layout(prs: PresentationT, slide_type: str) -> SlideLayout:
    """Return the best-matching layout for `slide_type`.

    Falls back to a blank-style layout, and finally to the first layout
    in the presentation.
    """
    layouts = list(prs.slide_layouts)
    names_lower = [layout.name.lower() for layout in layouts]

    for hint in LAYOUT_HINTS.get(slide_type, []):
        for i, name in enumerate(names_lower):
            if hint in name:
                return layouts[i]

    for hint in BLANK_FALLBACK_HINTS:
        for i, name in enumerate(names_lower):
            if hint in name:
                return layouts[i]

    return layouts[0]


def pick_blank_layout(prs: PresentationT) -> SlideLayout:
    """Return whatever the cleanest empty layout is — used when we draw shapes manually."""
    layouts = list(prs.slide_layouts)
    names_lower = [layout.name.lower() for layout in layouts]
    for hint in BLANK_FALLBACK_HINTS:
        for i, name in enumerate(names_lower):
            if hint in name:
                return layouts[i]
    return layouts[-1]
