"""Real syntax highlighting for code slides.

Tokenizes source with Pygments and yields styled spans so the renderer can
emit one colored run per token. This is what makes code slides look like an
editor instead of flat monospace text.

Falls back gracefully: unknown language -> plain lexer; unknown style ->
'monokai' -> pygments default.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List

from pptx.dml.color import RGBColor
from pygments import lex
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.styles import get_style_by_name
from pygments.util import ClassNotFound


@dataclass
class Span:
    text: str
    color: RGBColor
    bold: bool
    italic: bool


def _resolve_lexer(language: str, code: str):
    try:
        return get_lexer_by_name(language, stripnl=False)
    except ClassNotFound:
        try:
            return guess_lexer(code)
        except ClassNotFound:
            from pygments.lexers.special import TextLexer

            return TextLexer(stripnl=False)


def _resolve_style(name: str):
    for candidate in (name, "monokai", "default"):
        try:
            return get_style_by_name(candidate)
        except ClassNotFound:
            continue
    return get_style_by_name("default")


def highlight_lines(code: str, language: str, style_name: str, default_hex: str) -> List[List[Span]]:
    """Return a list of lines; each line is a list of Spans.

    `default_hex` is used when a token type has no explicit color in the style.
    """
    lexer = _resolve_lexer(language, code)
    style = _resolve_style(style_name)

    token_styles = dict(style)  # tokentype -> style dict

    lines: List[List[Span]] = [[]]
    for tok_type, value in lex(code, lexer):
        # Walk up the token hierarchy to find a defined style.
        s = None
        t = tok_type
        while t is not None:
            if t in token_styles and token_styles[t].get("color"):
                s = token_styles[t]
                break
            t = t.parent
        color_hex = (s.get("color") if s else None) or default_hex
        bold = bool(s.get("bold")) if s else False
        italic = bool(s.get("italic")) if s else False
        color = RGBColor.from_string(color_hex.upper())

        # Split on newlines so each line is its own list of spans.
        parts = value.split("\n")
        for i, part in enumerate(parts):
            if i > 0:
                lines.append([])
            if part:
                lines[-1].append(Span(part, color, bold, italic))

    # Drop a trailing empty line produced by a final newline.
    if len(lines) > 1 and not lines[-1]:
        lines.pop()
    return lines


def style_background_hex(style_name: str) -> str | None:
    """The Pygments style's own background color (so the card matches the theme)."""
    style = _resolve_style(style_name)
    bg = getattr(style, "background_color", None)
    if bg and isinstance(bg, str) and bg.startswith("#") and len(bg) == 7:
        return bg.lstrip("#")
    return None
