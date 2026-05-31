"""Render a Mermaid diagram block to PNG via @mermaid-js/mermaid-cli (mmdc).

Quality matters here: the default mermaid themes look generic, so we drive a
`theme: base` + `themeVariables` config whose colors match the deck. This is
the only theme mermaid lets you recolor (built-in themes hardcode their
palette). Output is high-DPI (scale 3) on a transparent background so it sits
cleanly on the slide surface.

Tries `mmdc` on PATH first, then `npx -y -p @mermaid-js/mermaid-cli mmdc`.

Usage:
    python scripts/render_mermaid.py input.mmd -o out.png
    cat diagram.mmd | python scripts/render_mermaid.py - -o out.png
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


class MermaidRenderError(RuntimeError):
    """mmdc invocation failed or is unavailable."""


def _resolve_mmdc() -> list[str] | None:
    for candidate in ("mmdc", "mmdc.cmd"):
        path = shutil.which(candidate)
        if path:
            return [path]
    npx = shutil.which("npx") or shutil.which("npx.cmd")
    if npx:
        return [npx, "-y", "-p", "@mermaid-js/mermaid-cli", "mmdc"]
    return None


def default_theme_vars(
    *,
    accent: str = "4F46E5",
    accent2: str = "06B6D4",
    ink: str = "14161F",
    line: str = "8A90A6",
    surface: str = "F4F6FA",
    font: str = "Calibri, Segoe UI, Helvetica, Arial, sans-serif",
) -> dict:
    """Build a mermaid `base` theme variable set from deck colors.

    All inputs are hex without '#'. Mermaid ignores color names, so we always
    pass hex.
    """
    def h(x: str) -> str:
        return "#" + x.lstrip("#")

    return {
        "theme": "base",
        "themeVariables": {
            "darkMode": False,
            "fontFamily": font,
            "fontSize": "16px",
            "primaryColor": h(surface),
            "primaryTextColor": h(ink),
            "primaryBorderColor": h(accent),
            "secondaryColor": "#FFFFFF",
            "secondaryTextColor": h(ink),
            "secondaryBorderColor": h(accent2),
            "tertiaryColor": "#FFFFFF",
            "tertiaryTextColor": h(ink),
            "tertiaryBorderColor": h(line),
            "lineColor": h(line),
            "textColor": h(ink),
            "mainBkg": h(surface),
            "nodeBorder": h(accent),
            "clusterBkg": "#FFFFFF",
            "clusterBorder": h(line),
            "edgeLabelBackground": "#FFFFFF",
            # sequence diagram
            "actorBkg": h(surface),
            "actorBorder": h(accent),
            "actorTextColor": h(ink),
            "signalColor": h(ink),
            "signalTextColor": h(ink),
            "labelBoxBkgColor": h(surface),
            "labelBoxBorderColor": h(accent),
            "noteBkgColor": "#FFF6D6",
            "noteTextColor": h(ink),
        },
        "flowchart": {"curve": "basis", "htmlLabels": True, "padding": 14},
        "sequence": {"useMaxWidth": True},
    }


def render(
    mermaid_source: str,
    output_path: Path,
    *,
    scale: int = 3,
    background: str = "transparent",
    theme_vars: dict | None = None,
) -> Path:
    """Render Mermaid source to PNG. Raises MermaidRenderError on failure."""
    cmd = _resolve_mmdc()
    if cmd is None:
        raise MermaidRenderError(
            "mmdc not found. Install via `npm install` in the skill directory, "
            "or `npm install -g @mermaid-js/mermaid-cli`."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    config = theme_vars if theme_vars is not None else default_theme_vars()

    tmp_in = tmp_cfg = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".mmd", delete=False, encoding="utf-8"
        ) as f:
            f.write(mermaid_source)
            tmp_in = Path(f.name)
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        ) as f:
            json.dump(config, f)
            tmp_cfg = Path(f.name)

        full_cmd = [
            *cmd,
            "-i", str(tmp_in),
            "-o", str(output_path),
            "-c", str(tmp_cfg),
            "-s", str(scale),
            "-b", background,
        ]
        proc = subprocess.run(full_cmd, capture_output=True, text=True, shell=False)
        if proc.returncode != 0:
            raise MermaidRenderError(
                f"mmdc failed ({proc.returncode}):\n"
                f"stdout: {proc.stdout}\nstderr: {proc.stderr}"
            )
    finally:
        for p in (tmp_in, tmp_cfg):
            if p is not None:
                try:
                    os.unlink(p)
                except OSError:
                    pass

    if not output_path.exists():
        raise MermaidRenderError(f"mmdc returned 0 but {output_path} was not created")
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Render a Mermaid diagram to PNG")
    parser.add_argument("input", type=str, help="Path to .mmd file, or - for stdin")
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("-s", "--scale", type=int, default=3)
    parser.add_argument("-b", "--background", default="transparent")
    args = parser.parse_args(argv)

    if args.input == "-":
        source = sys.stdin.read()
    else:
        p = Path(args.input)
        if not p.exists():
            print(f"error: {p} not found", file=sys.stderr)
            return 2
        source = p.read_text(encoding="utf-8")

    try:
        render(source, args.output, scale=args.scale, background=args.background)
    except MermaidRenderError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
