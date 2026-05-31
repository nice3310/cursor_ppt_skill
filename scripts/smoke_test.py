"""Smoke test: validate -> build every outline in ../examples/.

Run from the skill root:
    python scripts/smoke_test.py

Exit code:
  0 - every example validated and built (warnings allowed)
  1 - one or more failures

Diagram rendering is disabled by default because mmdc is optional. Pass
--with-mermaid to exercise the full pipeline.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from build_deck import build  # noqa: E402
from validate_outline import density_warnings, load_outline  # noqa: E402

SCRIPT_DIR = Path(__file__).parent
EXAMPLES = SCRIPT_DIR.parent / "examples"


def main(argv=None):
    parser = argparse.ArgumentParser(description="Round-trip every example outline")
    parser.add_argument(
        "--with-mermaid",
        action="store_true",
        help="Actually render diagrams via mmdc (slow, requires Node)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat density warnings as failures",
    )
    args = parser.parse_args(argv)

    outlines = sorted(EXAMPLES.glob("*.outline.yaml"))
    if not outlines:
        print(f"no examples found in {EXAMPLES}", file=sys.stderr)
        return 1

    failures = 0
    for outline_path in outlines:
        print(f"\n=== {outline_path.name} ===")
        try:
            outline = load_outline(outline_path)
        except Exception as e:  # noqa: BLE001
            print(f"  schema/yaml error: {e}")
            failures += 1
            continue

        warns = density_warnings(outline)
        if warns:
            print(f"  {len(warns)} warning(s):")
            for w in warns:
                print(f"    - {w}")
            if args.strict:
                failures += 1
                continue
        else:
            print("  validate: OK")

        try:
            out = build(outline_path, None, mermaid=args.with_mermaid)
            print(f"  build: wrote {out}")
        except Exception as e:  # noqa: BLE001
            print(f"  build failed: {e}")
            failures += 1

    print(f"\n{len(outlines) - failures}/{len(outlines)} examples passed")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
