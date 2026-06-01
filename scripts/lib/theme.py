"""Design system for cursor_ppt_skill.

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


def _ocean() -> Theme:
    """Deep blue palette — pairs naturally with the 'modern' style."""
    return Theme(
        name="ocean",
        ink=rgb("E8EAF6"),
        ink_soft=rgb("90A4AE"),
        background=rgb("0D1B2A"),
        surface=rgb("1B2838"),
        surface_edge=rgb("2A3A4E"),
        rule=rgb("2A3A4E"),
        accent=rgb("0EA5E9"),
        accent_soft=rgb("0C2D48"),
        accent2=rgb("38BDF8"),
        on_accent=rgb("FFFFFF"),
        on_accent_soft=rgb("BAE6FD"),
        code_bg=rgb("081420"),
        code_default=rgb("D0E8F8"),
        table_header_bg=rgb("0EA5E9"),
        table_header_fg=rgb("FFFFFF"),
        table_zebra=rgb("1B2838"),
        pygments_style="monokai",
    )


def _warm() -> Theme:
    """Amber and earth tones — pairs naturally with the 'vibrant' style."""
    return Theme(
        name="warm",
        ink=rgb("1C1410"),
        ink_soft=rgb("6B5B4F"),
        background=rgb("FFFBF5"),
        surface=rgb("FFF3E0"),
        surface_edge=rgb("F0D9B5"),
        rule=rgb("F0D9B5"),
        accent=rgb("E65100"),
        accent_soft=rgb("FFF3E0"),
        accent2=rgb("F59E0B"),
        on_accent=rgb("FFFFFF"),
        on_accent_soft=rgb("FFCC80"),
        code_bg=rgb("1A1210"),
        code_default=rgb("F0DCC8"),
        table_header_bg=rgb("E65100"),
        table_header_fg=rgb("FFFFFF"),
        table_zebra=rgb("FFF3E0"),
        pygments_style="monokai",
    )


def _forest() -> Theme:
    """Deep green and gold — sophisticated and grounded."""
    return Theme(
        name="forest",
        ink=rgb("1A2E1A"),
        ink_soft=rgb("5C7A5C"),
        background=rgb("F5F9F0"),
        surface=rgb("E8F0E0"),
        surface_edge=rgb("C8D8B8"),
        rule=rgb("C8D8B8"),
        accent=rgb("2E7D32"),
        accent_soft=rgb("E8F5E9"),
        accent2=rgb("C9A837"),
        on_accent=rgb("FFFFFF"),
        on_accent_soft=rgb("A5D6A7"),
        code_bg=rgb("101810"),
        code_default=rgb("D0E8D0"),
        table_header_bg=rgb("2E7D32"),
        table_header_fg=rgb("FFFFFF"),
        table_zebra=rgb("E8F0E0"),
        pygments_style="monokai",
    )


_PRESETS = {
    "light": _light,
    "dark": _dark,
    "ocean": _ocean,
    "warm": _warm,
    "forest": _forest,
}


def get_theme(name: str | None = None) -> Theme:
    if not name:
        return _light()
    return _PRESETS.get(name.lower(), _light)()


def default_theme() -> Theme:
    return _light()

