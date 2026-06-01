"""Visual style system for cursor_ppt_skill.

A Style is a bundle of *layout geometry and decorative language* tokens.
It is **independent of colour** — colour is handled by Theme. A Style
controls *where* things go and *how* they are decorated:

  - Title slide composition (rail, centered, split, …)
  - Accent bar shape (short bar, underline, left thick, …)
  - Bullet marker shape (rounded rect, circle, dash, pill, diamond)
  - Section divider treatment
  - Card / surface treatment (comparison, code cards)
  - Statement & quote visual variants

Five presets ship: editorial (default / backward-compat), corporate,
minimal, modern, vibrant.  Select via ``meta.style`` in the outline.

The ``recommend_style()`` helper picks a style based on free-text
content description — used when the user selects "Surprise me".
"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass


# ---------------------------------------------------------------------------
# Style dataclass
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Style:
    name: str

    # --- Title slide ---
    # "rail"       — left accent rail + text offset (editorial)
    # "centered"   — centered text, minimal decoration (minimal)
    # "full_bleed" — full-width accent band at top (corporate)
    # "split"      — left half accent, right half content (modern)
    # "diagonal"   — diagonal accent element (vibrant)
    title_variant: str = "rail"

    # --- Accent bar on content slides ---
    # "short_bar"  — short rounded rect under title (editorial)
    # "underline"  — full-width thin line under title (corporate)
    # "left_thick" — thick left-edge bar beside title (minimal)
    # "dot"        — small circle accent (modern)
    # "none"       — no accent bar (vibrant — relies on color)
    accent_bar: str = "short_bar"

    # --- Bullet marker ---
    # "rounded_rect" — small rounded squares (editorial)
    # "circle"       — filled circles (corporate)
    # "dash"         — short horizontal dashes (minimal)
    # "pill"         — pill / capsule shape (modern)
    # "diamond"      — diamond / rotated square (vibrant)
    bullet_marker: str = "rounded_rect"

    # --- Section divider ---
    # "dark_full"    — full dark background + ghosted number (editorial)
    # "accent_band"  — accent band across middle (corporate)
    # "clean"        — white bg, large light number (minimal)
    # "gradient"     — gradient from accent to bg (modern)
    # "color_block"  — full accent2 background (vibrant)
    section_variant: str = "dark_full"

    # --- Card treatment (comparison columns, code background) ---
    # "bordered_round" — rounded corners + border + shadow (editorial)
    # "flat_fill"      — flat fill, no border, no shadow (corporate)
    # "border_only"    — thin border only, no fill, no shadow (minimal)
    # "heavy_shadow"   — no border, prominent shadow (modern)
    # "left_accent"    — left accent bar on cards (vibrant)
    card_style: str = "bordered_round"

    # --- Statement slide ---
    # "full_accent"    — full accent background (editorial)
    # "centered_large" — white bg, extra-large text (corporate)
    # "with_line"      — centered text + thin lines above/below (minimal)
    # "gradient_bg"    — gradient background (modern)
    # "dual_color"     — two-tone background (vibrant)
    statement_variant: str = "full_accent"

    # --- Quote slide ---
    # "big_mark"      — large quotation mark + surface bg (editorial)
    # "left_bar"      — thick left accent bar (corporate)
    # "italic_center" — centered italic, no decorations (minimal)
    # "card"          — quote inside a card with shadow (modern)
    # "accent_bg"     — accent background + white text (vibrant)
    quote_variant: str = "big_mark"

    # --- Geometry modifiers ---
    # Content top offset (inches delta from default CONTENT_TOP)
    content_top_offset: float = 0.0
    # Margin scale (1.0 = default, >1 = more whitespace)
    margin_scale: float = 1.0

    # --- Visual rhythm ---
    # Whether to show a thin rule under the kicker
    use_kicker_line: bool = False
    # Whether to italicise headlines for a softer feel
    headline_italic: bool = False
    # Footer style: "rule" (thin line + text) or "minimal" (text only)
    footer_style: str = "rule"

    # --- Macro-Layouts ---
    layout_bullets: str = "classic"
    layout_image: str = "classic"
    layout_comparison: str = "classic"
    layout_code: str = "classic"
    layout_table: str = "classic"


# ---------------------------------------------------------------------------
# Presets
# ---------------------------------------------------------------------------


def editorial() -> Style:
    """Clean, professional, accent bars and surface cards.  Default."""
    return Style(name="editorial")


def corporate() -> Style:
    """Bold headings, strong colour blocks, authoritative."""
    return Style(
        name="corporate",
        title_variant="full_bleed",
        accent_bar="underline",
        bullet_marker="circle",
        section_variant="accent_band",
        card_style="flat_fill",
        statement_variant="centered_large",
        quote_variant="left_bar",
        content_top_offset=0.05,
        margin_scale=0.95,
        footer_style="minimal",
        layout_image="split_bleed",
        layout_comparison="columns",
        layout_table="clean",
    )


def minimal() -> Style:
    """Lots of whitespace, understated elegance."""
    return Style(
        name="minimal",
        title_variant="centered",
        accent_bar="left_thick",
        bullet_marker="dash",
        section_variant="clean",
        card_style="border_only",
        statement_variant="with_line",
        quote_variant="italic_center",
        content_top_offset=0.15,
        margin_scale=1.2,
        headline_italic=True,
        footer_style="minimal",
        layout_bullets="magazine",
        layout_image="split_bleed",
        layout_comparison="columns",
        layout_table="clean",
    )


def modern() -> Style:
    """Gradients, rounded cards, contemporary feel."""
    return Style(
        name="modern",
        title_variant="split",
        accent_bar="dot",
        bullet_marker="pill",
        section_variant="gradient",
        card_style="heavy_shadow",
        statement_variant="gradient_bg",
        quote_variant="card",
        content_top_offset=0.0,
        margin_scale=1.05,
        use_kicker_line=True,
        layout_bullets="grid",
        layout_image="immersive",
        layout_comparison="split_vs",
        layout_table="floating_cards",
    )


def vibrant() -> Style:
    """Colorful, energetic, eye-catching."""
    return Style(
        name="vibrant",
        title_variant="diagonal",
        accent_bar="none",
        bullet_marker="diamond",
        section_variant="color_block",
        card_style="left_accent",
        statement_variant="dual_color",
        quote_variant="accent_bg",
        content_top_offset=-0.05,
        margin_scale=0.95,
        layout_bullets="magazine",
        layout_image="split_bleed",
        layout_code="side_by_side",
        layout_comparison="rows",
        layout_table="classic",
    )


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

_PRESETS = {
    "editorial": editorial,
    "corporate": corporate,
    "minimal": minimal,
    "modern": modern,
    "vibrant": vibrant,
}


def get_style(name: str | None = None) -> Style:
    """Return the named style preset, or ``editorial`` if *name* is None."""
    if not name:
        return editorial()
    return _PRESETS.get(name.lower(), editorial)()


def list_styles() -> list[str]:
    return list(_PRESETS.keys())


def random_style() -> Style:
    """Pick a random style preset."""
    return _PRESETS[random.choice(list(_PRESETS.keys()))]()


# ---------------------------------------------------------------------------
# Smart recommendation
# ---------------------------------------------------------------------------

# Keyword groups used by recommend_style() to infer content type.
_TECHNICAL_KW = re.compile(
    r"\b(code|api|sdk|architecture|infra|backend|frontend|migration|deploy|"
    r"debug|refactor|diff|pr|pull.?request|commit|pipeline|cicd|devops|"
    r"database|schema|microservice|kubernetes|docker|aws|gcp|azure)\b",
    re.IGNORECASE,
)
_BUSINESS_KW = re.compile(
    r"\b(revenue|okr|kpi|roadmap|strategy|exec|board|investor|pitch|"
    r"budget|quarterly|annual|stakeholder|roi|market|growth|sales|"
    r"leadership|ceo|cfo|cto|vp|director|c.suite)\b",
    re.IGNORECASE,
)
_TEACHING_KW = re.compile(
    r"\b(explain|primer|intro|tutorial|concept|learn|teach|workshop|"
    r"training|onboard|101|beginner|fundamental|overview|what.is)\b",
    re.IGNORECASE,
)
_CREATIVE_KW = re.compile(
    r"\b(design|brand|creative|marketing|campaign|launch|product|"
    r"ux|ui|user.experience|demo|showcase|portfolio|pitch.deck)\b",
    re.IGNORECASE,
)


def recommend_style(description: str) -> Style:
    """Pick a style based on a free-text content description.

    Heuristic precedence: business → creative → teaching → technical.
    Falls back to a random pick if nothing matches.
    """
    scores: dict[str, int] = {
        "technical": len(_TECHNICAL_KW.findall(description)),
        "business": len(_BUSINESS_KW.findall(description)),
        "teaching": len(_TEACHING_KW.findall(description)),
        "creative": len(_CREATIVE_KW.findall(description)),
    }

    top = max(scores, key=scores.get)  # type: ignore[arg-type]
    if scores[top] == 0:
        # No keywords matched — truly random.
        return random_style()

    mapping: dict[str, list[str]] = {
        "technical": ["editorial", "minimal"],
        "business": ["corporate", "modern"],
        "teaching": ["modern", "minimal"],
        "creative": ["vibrant", "modern"],
    }
    candidates = mapping[top]
    return get_style(random.choice(candidates))


# ---------------------------------------------------------------------------
# Theme harmony — suggested default theme for each style
# ---------------------------------------------------------------------------

STYLE_DEFAULT_THEME: dict[str, str] = {
    "editorial": "light",
    "corporate": "dark",
    "minimal": "light",
    "modern": "ocean",
    "vibrant": "warm",
}


def suggested_theme(style_name: str) -> str:
    """Return the harmonious default theme name for a style."""
    return STYLE_DEFAULT_THEME.get(style_name, "light")
