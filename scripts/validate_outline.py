"""Validate an outline.yaml file.

Two layers:
  1. Pydantic schema validation (hard errors).
  2. Density + pacing rules (soft warnings).

Exit code:
  0 - clean, or only warnings present and --strict not passed
  1 - schema error, or warnings with --strict
  2 - bad invocation
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import yaml
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).parent))
from lib.schema import (  # noqa: E402
    CONTENT_SLIDE_TYPES,
    BulletsSlide,
    CodeSlide,
    ComparisonSlide,
    DiagramSlide,
    ImageSlide,
    Outline,
    TableSlide,
)


# Density thresholds. Tuned to match methodology.md billboard test.
MAX_BULLETS_PER_LIST = 5
MAX_WORDS_PER_BULLET = 15
MAX_CODE_LINES = 20
MAX_TITLE_CHARS = 60
MAX_HEADLINE_CHARS = 100
MAX_TABLE_COLS = 5
MAX_TABLE_ROWS = 7
MIN_MINUTES_PER_CONTENT_SLIDE = 1.0
MAX_MINUTES_PER_CONTENT_SLIDE = 3.0


def _word_count(s: str) -> int:
    return len(s.split())


def density_warnings(outline: Outline) -> list[str]:
    warns: list[str] = []

    for i, slide in enumerate(outline.slides):
        loc = f"slides[{i}] ({slide.type})"

        title = getattr(slide, "title", None)
        if isinstance(title, str) and len(title) > MAX_TITLE_CHARS:
            warns.append(f"{loc}: title is {len(title)} chars (>{MAX_TITLE_CHARS})")

        headline = getattr(slide, "headline", None)
        if isinstance(headline, str) and len(headline) > MAX_HEADLINE_CHARS:
            warns.append(
                f"{loc}: headline is {len(headline)} chars (>{MAX_HEADLINE_CHARS})"
            )

        if isinstance(slide, BulletsSlide):
            _bullet_warnings(slide.bullets, loc, warns)

        if isinstance(slide, ComparisonSlide):
            for j, col in enumerate(slide.columns):
                _bullet_warnings(col.bullets, f"{loc}.columns[{j}]", warns)

        if isinstance(slide, CodeSlide):
            n = slide.code.count("\n") + (0 if slide.code.endswith("\n") else 1)
            if n > MAX_CODE_LINES:
                warns.append(
                    f"{loc}: code is {n} lines (>{MAX_CODE_LINES}); split or move to appendix"
                )

        if isinstance(slide, TableSlide):
            if len(slide.columns) > MAX_TABLE_COLS:
                warns.append(
                    f"{loc}: table has {len(slide.columns)} columns (>{MAX_TABLE_COLS})"
                )
            if len(slide.rows) > MAX_TABLE_ROWS:
                warns.append(
                    f"{loc}: table has {len(slide.rows)} rows (>{MAX_TABLE_ROWS})"
                )

        if isinstance(slide, DiagramSlide):
            if not slide.mermaid.strip():
                warns.append(f"{loc}: mermaid block is empty")

        if isinstance(slide, ImageSlide):
            if slide.bullets:
                _bullet_warnings(slide.bullets, loc, warns)
            if not slide.source:
                warns.append(
                    f"{loc}: image has no `source`; add a citation/URL "
                    "(goes into speaker notes)"
                )

    warns.extend(_pacing_warnings(outline))
    return warns


def _bullet_warnings(bullets: list[str], loc: str, warns: list[str]) -> None:
    if len(bullets) > MAX_BULLETS_PER_LIST:
        warns.append(
            f"{loc}: {len(bullets)} bullets (>{MAX_BULLETS_PER_LIST}); split the slide"
        )
    for k, b in enumerate(bullets):
        wc = _word_count(b)
        if wc > MAX_WORDS_PER_BULLET:
            warns.append(
                f"{loc}.bullets[{k}]: {wc} words (>{MAX_WORDS_PER_BULLET}); tighten"
            )


def _pacing_warnings(outline: Outline) -> list[str]:
    content = [s for s in outline.slides if s.type in CONTENT_SLIDE_TYPES]
    if not content:
        return ["narrative: no content slides found"]

    mins = outline.narrative.duration_minutes
    rate = mins / len(content)
    out: list[str] = []
    if rate < MIN_MINUTES_PER_CONTENT_SLIDE:
        out.append(
            f"pacing: {len(content)} content slides for {mins} min "
            f"= {rate:.1f} min/slide (<{MIN_MINUTES_PER_CONTENT_SLIDE}); too rushed"
        )
    elif rate > MAX_MINUTES_PER_CONTENT_SLIDE:
        out.append(
            f"pacing: {len(content)} content slides for {mins} min "
            f"= {rate:.1f} min/slide (>{MAX_MINUTES_PER_CONTENT_SLIDE}); add detail or trim time"
        )
    return out


def load_outline(path: Path) -> Outline:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    return Outline.model_validate(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate an outline.yaml file.")
    parser.add_argument("outline", type=Path, help="Path to outline.yaml")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero on density warnings too",
    )
    args = parser.parse_args(argv)

    if not args.outline.exists():
        print(f"error: {args.outline} not found", file=sys.stderr)
        return 2

    try:
        outline = load_outline(args.outline)
    except ValidationError as e:
        print("schema errors:", file=sys.stderr)
        for err in e.errors():
            loc = ".".join(str(p) for p in err["loc"])
            print(f"  {loc}: {err['msg']}", file=sys.stderr)
        return 1
    except yaml.YAMLError as e:
        print(f"yaml parse error: {e}", file=sys.stderr)
        return 1

    warns = density_warnings(outline)
    if warns:
        print(f"{len(warns)} warning(s):")
        for w in warns:
            print(f"  - {w}")
        if args.strict:
            return 1
        print("\n(run with --strict to make warnings fail)")
    else:
        print(f"OK: {len(outline.slides)} slides validated.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
