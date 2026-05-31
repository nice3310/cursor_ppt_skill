"""Scan an image folder and emit a manifest the agent can reason over.

The image workflow: the user (or the agent, via web download) drops images
into an assets folder. This script reports each image's dimensions, aspect
ratio, and a layout hint (wide / tall / square) so the agent can choose how to
place it: full-bleed, side-by-side with text, etc.

Usage:
    python scripts/scan_assets.py path/to/assets
    python scripts/scan_assets.py path/to/assets -o assets-manifest.yaml
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import yaml

try:
    from PIL import Image
except Exception:  # noqa: BLE001
    Image = None

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}


def _kind(ar: float) -> str:
    if ar >= 1.4:
        return "wide"      # good for full-width or top
    if ar <= 0.75:
        return "tall"      # good for side-by-side with text
    return "square"        # flexible


def _suggested_layout(kind: str) -> str:
    return {"wide": "full", "tall": "right", "square": "right"}[kind]


def scan(folder: Path) -> dict[str, Any]:
    if Image is None:
        raise RuntimeError("Pillow is required: pip install pillow")

    images = []
    for p in sorted(folder.iterdir()):
        if p.suffix.lower() not in IMAGE_EXTS:
            continue
        entry: dict[str, Any] = {"file": p.name, "path": str(p)}
        try:
            with Image.open(p) as im:
                w, h = im.size
                ar = round(w / h, 3) if h else None
                entry.update(
                    width=w,
                    height=h,
                    aspect_ratio=ar,
                    kind=_kind(ar) if ar else "unknown",
                    suggested_layout=_suggested_layout(_kind(ar)) if ar else "full",
                )
        except Exception as e:  # noqa: BLE001
            entry["error"] = str(e)
        images.append(entry)

    return {"folder": str(folder), "count": len(images), "images": images}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Scan an image folder")
    parser.add_argument("folder", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    args = parser.parse_args(argv)

    if not args.folder.exists() or not args.folder.is_dir():
        print(f"error: {args.folder} is not a folder", file=sys.stderr)
        return 2

    try:
        manifest = scan(args.folder)
    except RuntimeError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    text = yaml.safe_dump(manifest, sort_keys=False, allow_unicode=True)
    if args.output:
        args.output.write_text(text, encoding="utf-8")
        print(f"wrote {args.output} ({manifest['count']} images)")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
