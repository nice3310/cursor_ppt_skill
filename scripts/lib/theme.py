"""Design system for cursor-ppt-skill.

A Theme is a bundle of color, type, and spacing tokens that the renderers read.
Two presets ship: "light" (default, editorial) and "dark". Select via
`meta.theme` in the outline, or fall back to the default.

The goal is a deck that looks *designed*, not just non-broken: strong type
hierarchy, an accent system, surface cards, full-bleed impact slides, a
consistent footer + kicker, and a dark syntax-highlighted code card.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pptx.dml.color import RGBColor


def rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.lstrip("#").upper())


@dataclass(frozen=True)
class Theme:
    name: str = "light"

    # --- Core palette ---
    ink: RGBColor = field(default_factory=lambda: rgb("14161F"))          # primary text
    ink_soft: RGBColor = field(default_factory=lambda: rgb("5B6172"))     # secondary text
    background: RGBColor = field(default_factory=lambda: rgb("FFFFFF"))   # slide background
    surface: RGBColor = field(default_factory=lambda: rgb("F4F6FA"))      # cards / panels
    surface_edge: RGBColor = field(default_factory=lambda: rgb("E5E8F0")) # card hairline
    rule: RGBColor = field(default_factory=lambda: rgb("E5E8F0"))         # divider lines

    accent: RGBColor = field(default_factory=lambda: rgb("4F46E5"))       # primary accent (indigo)
    accent_soft: RGBColor = field(default_factory=lambda: rgb("EEF0FE"))  # accent tint fill
    accent2: RGBColor = field(default_factory=lambda: rgb("06B6D4"))      # secondary accent (cyan)

    # Text color to use *on top of* an accent / dark full-bleed background
    on_accent: RGBColor = field(default_factory=lambda: rgb("FFFFFF"))
    on_accent_soft: RGBColor = field(default_factory=lambda: rgb("C9CBF5"))

    # Code card (always dark for contrast, independent of light/dark preset)
    code_bg: RGBColor = field(default_factory=lambda: rgb("1E1E2E"))
    code_default: RGBColor = field(default_factory=lambda: rgb("E6E6F0"))
    code_line_highlight: RGBColor = field(default_factory=lambda: rgb("2D2D44"))
    pygments_style: str = "monokai"

    # Table
    table_header_bg: RGBColor = field(default_factory=lambda: rgb("4F46E5"))
    table_header_fg: RGBColor = field(default_factory=lambda: rgb("FFFFFF"))
    table_zebra: RGBColor = field(default_factory=lambda: rgb("F4F6FA"))

    # --- Type scale (Pt) ---
    size_kicker: int = 13          # eyebrow above title (uppercase, tracked)
    size_title: int = 32           # content-slide title
    size_section: int = 46         # section divider
    size_section_num: int = 120    # ghosted section number
    size_statement: int = 40       # statement slide
    size_headline: int = 20        # assertion sentence under title
    size_body: int = 18            # bullets, table cells
    size_body_lg: int = 20         # summary / emphasised body
    size_caption: int = 13
    size_code: int = 14
    size_quote: int = 32
    size_quote_mark: int = 130
    size_attribution: int = 16
    size_footer: int = 10
    size_qa: int = 54

    # --- Fonts ---
    font_display: str = "Calibri"  # titles / headlines
    font_body: str = "Calibri"     # body text
    font_mono: str = "Consolas"

    # --- Geometry (inches) ---
    margin_left: float = 0.7
    margin_right: float = 0.7
    margin_top: float = 0.55
    margin_bottom: float = 0.5
    footer_height: float = 0.35
    accent_bar_w: float = 0.6      # short accent bar under title
    accent_bar_h: float = 0.07
    corner_radius: float = 0.12

    # Content top: where the body area begins on a titled slide
    content_top: float = 2.0


def _light() -> Theme:
    return Theme(name="light")


def _dark() -> Theme:
    return Theme(
        name="dark",
        ink=rgb("F2F3F8"),
        ink_soft=rgb("A4A9BD"),
        background=rgb("12131A"),
        surface=rgb("1C1E2A"),
        surface_edge=rgb("2C2F3F"),
        rule=rgb("2C2F3F"),
        accent=rgb("7C7BFF"),
        accent_soft=rgb("23243A"),
        accent2=rgb("22D3EE"),
        on_accent=rgb("FFFFFF"),
        on_accent_soft=rgb("D7D8FA"),
        code_bg=rgb("0E0E16"),
        table_header_bg=rgb("7C7BFF"),
        table_zebra=rgb("1C1E2A"),
        pygments_style="monokai",
    )


_PRESETS = {
    "light": _light,
    "dark": _dark,
}


def get_theme(name: str | None = None) -> Theme:
    if not name:
        return _light()
    return _PRESETS.get(name.lower(), _light)()


def default_theme() -> Theme:
    return _light()
