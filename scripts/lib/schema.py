"""Pydantic models for outline.yaml.

These models are the single source of truth for the structured prompt format.
Both validate_outline.py and build_deck.py load YAML through `Outline.model_validate`.

Hard structural rules live in the validators here. "Soft" density warnings
(e.g. >5 bullets is suspicious but legal) live in validate_outline.py.

Type annotations use `typing.Optional` / `typing.List` etc. (not PEP 604/585
shorthand) so this file evaluates cleanly on Python 3.9.
"""

from __future__ import annotations

from datetime import date as _date
from pathlib import Path
from typing import Annotated, List, Literal, Optional, Set, Union

from pydantic import BaseModel, ConfigDict, Field, model_validator


SlideType = Literal[
    "title",
    "section",
    "statement",
    "bullets",
    "comparison",
    "code",
    "diagram",
    "image",
    "table",
    "quote",
    "summary",
    "qa",
]


class _Base(BaseModel):
    model_config = ConfigDict(extra="forbid")


# ---------------------------------------------------------------------------
# meta / narrative
# ---------------------------------------------------------------------------


class Meta(_Base):
    title: str = Field(..., min_length=1, max_length=120)
    subtitle: Optional[str] = None
    presenter: Optional[str] = None
    date: Optional[Union[_date, str]] = None
    template: Optional[str] = None
    output: str = Field(..., min_length=1)
    theme: Optional[Literal["light", "dark", "ocean", "warm", "forest"]] = None
    style: Optional[Literal["editorial", "corporate", "minimal", "modern", "vibrant"]] = None
    assets_dir: Optional[str] = None
    footer: Optional[str] = None  # footer text on content slides; defaults to meta.title


class Narrative(_Base):
    audience: str = Field(..., min_length=1)
    duration_minutes: int = Field(..., ge=1, le=180)
    goal: str = Field(..., min_length=1)
    key_takeaways: List[str] = Field(..., min_length=1, max_length=5)


# ---------------------------------------------------------------------------
# slide types
# ---------------------------------------------------------------------------


class TitleSlide(_Base):
    type: Literal["title"]
    title: str = Field(..., min_length=1, max_length=120)
    subtitle: Optional[str] = None
    presenter: Optional[str] = None
    date: Optional[Union[_date, str]] = None
    notes: Optional[str] = None


class SectionSlide(_Base):
    type: Literal["section"]
    title: str = Field(..., min_length=1, max_length=80)
    number: Optional[str] = None  # ghosted numeral; auto-numbered if omitted
    notes: Optional[str] = None


class StatementSlide(_Base):
    type: Literal["statement"]
    headline: str = Field(..., min_length=1, max_length=200)
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None


class BulletsSlide(_Base):
    type: Literal["bullets"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    bullets: List[str] = Field(..., min_length=1)
    kicker: Optional[str] = Field(default=None, max_length=40)
    numbered: bool = False  # render as numbered chips instead of dots
    notes: Optional[str] = None


class ComparisonColumn(_Base):
    label: str = Field(..., min_length=1, max_length=40)
    bullets: List[str] = Field(..., min_length=1)


class ComparisonSlide(_Base):
    type: Literal["comparison"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    columns: List[ComparisonColumn] = Field(..., min_length=2, max_length=3)
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None


class CodeSlide(_Base):
    type: Literal["code"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    language: str = Field(default="text")
    code: str = Field(..., min_length=1)
    highlight_lines: Optional[List[int]] = None
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None


class DiagramSlide(_Base):
    type: Literal["diagram"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    mermaid: str = Field(..., min_length=1)
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None


class ImageSlide(_Base):
    type: Literal["image"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    path: str = Field(..., min_length=1)
    caption: Optional[str] = None
    source: Optional[str] = None  # URL or citation; rendered into speaker notes
    bullets: Optional[List[str]] = None  # if set, text sits beside the image
    layout: Literal["auto", "full", "right", "left"] = "auto"
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None


class TableSlide(_Base):
    type: Literal["table"]
    title: str = Field(..., min_length=1, max_length=60)
    headline: str = Field(..., min_length=1, max_length=200)
    columns: List[str] = Field(..., min_length=1)
    rows: List[List[str]] = Field(..., min_length=1)
    kicker: Optional[str] = Field(default=None, max_length=40)
    notes: Optional[str] = None

    @model_validator(mode="after")
    def _check_row_widths(self) -> "TableSlide":
        n = len(self.columns)
        for i, row in enumerate(self.rows):
            if len(row) != n:
                raise ValueError(
                    f"table row {i} has {len(row)} cells but columns has {n}"
                )
        return self


class QuoteSlide(_Base):
    type: Literal["quote"]
    quote: str = Field(..., min_length=1, max_length=400)
    attribution: Optional[str] = None
    notes: Optional[str] = None


class SummarySlide(_Base):
    type: Literal["summary"]
    title: str = Field(default="Takeaways", max_length=60)
    points: List[str] = Field(..., min_length=1, max_length=5)
    notes: Optional[str] = None


class QASlide(_Base):
    type: Literal["qa"]
    title: str = Field(default="Questions?", max_length=60)
    contact: Optional[str] = None
    notes: Optional[str] = None


Slide = Annotated[
    Union[
        TitleSlide,
        SectionSlide,
        StatementSlide,
        BulletsSlide,
        ComparisonSlide,
        CodeSlide,
        DiagramSlide,
        ImageSlide,
        TableSlide,
        QuoteSlide,
        SummarySlide,
        QASlide,
    ],
    Field(discriminator="type"),
]


# Content slide types that have a non-trivial body; used by the validator's
# density rules and by build_deck.py to decide pacing math.
CONTENT_SLIDE_TYPES: Set[str] = {
    "statement",
    "bullets",
    "comparison",
    "code",
    "diagram",
    "image",
    "table",
}


class Outline(_Base):
    meta: Meta
    narrative: Narrative
    slides: List[Slide] = Field(..., min_length=2)

    def resolve_path(self, base: Path, value: str) -> Path:
        """Resolve a path relative to the outline file's directory."""
        p = Path(value)
        return p if p.is_absolute() else (base / p).resolve()
