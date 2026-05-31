"""Per-slide-type renderers with a real design system.

Each renderer draws its own shapes on a chosen layout so we inherit the
template's master while controlling the visual treatment: kicker eyebrows,
accent bars, surface cards, full-bleed impact slides, a dark syntax-
highlighted code card, zebra tables, and a consistent footer.

Positioning assumes a 16:9 slide (13.333 x 7.5 in) but reads the real size
from the presentation, so it degrades reasonably on 4:3.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationT
from pptx.slide import Slide as PptxSlide
from pptx.util import Emu, Inches, Pt

from .code_highlight import highlight_lines, style_background_hex
from .layouts import pick_blank_layout
from .schema import (
    BulletsSlide,
    CodeSlide,
    ComparisonSlide,
    DiagramSlide,
    ImageSlide,
    QASlide,
    QuoteSlide,
    SectionSlide,
    Slide,
    StatementSlide,
    SummarySlide,
    TableSlide,
    TitleSlide,
)
from .theme import Theme

try:
    from PIL import Image
except Exception:  # noqa: BLE001
    Image = None


# Vertical layout constants (inches) for content slides.
KICKER_TOP = 0.55
TITLE_TOP = 0.92
BAR_TOP = 1.64
HEADLINE_TOP = 1.78
CONTENT_TOP = 2.45


# ---------------------------------------------------------------------------
# low-level helpers
# ---------------------------------------------------------------------------


def _slide_size_in(prs: PresentationT) -> tuple[float, float]:
    return Emu(prs.slide_width).inches, Emu(prs.slide_height).inches


def _new_slide(prs: PresentationT) -> PptxSlide:
    return prs.slides.add_slide(pick_blank_layout(prs))


def _bg(slide: PptxSlide, color: RGBColor) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _rect(slide, shape_type, left, top, width, height, fill=None, line=None):
    shp = slide.shapes.add_shape(
        shape_type, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    return shp


def _soft_shadow(shape, blur_in=0.06, dist_in=0.035, alpha=24, dir_deg=90):
    """Inject a subtle outer drop shadow so cards/images read as layered."""
    try:
        spPr = shape._element.spPr
    except AttributeError:
        return
    existing = spPr.find(qn("a:effectLst"))
    if existing is not None:
        spPr.remove(existing)
    eff = spPr.makeelement(qn("a:effectLst"), {})
    shdw = eff.makeelement(
        qn("a:outerShdw"),
        {
            "blurRad": str(int(Inches(blur_in))),
            "dist": str(int(Inches(dist_in))),
            "dir": str(int(dir_deg * 60000)),
            "rotWithShape": "0",
        },
    )
    clr = shdw.makeelement(qn("a:srgbClr"), {"val": "000000"})
    a = clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 1000))})
    clr.append(a)
    shdw.append(clr)
    eff.append(shdw)
    spPr.append(eff)


def _textbox(
    slide,
    *,
    left,
    top,
    width,
    height,
    anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    return box, tf


def _run(p, text, *, size, color, bold=False, italic=False, font=None, tracking=None):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    if font:
        r.font.name = font
    if tracking is not None:
        r._r.get_or_add_rPr().set("spc", str(int(tracking * 100)))
    return r


def _simple_text(
    slide,
    text,
    *,
    left,
    top,
    width,
    height,
    size,
    color,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font=None,
    tracking=None,
    line_spacing=None,
):
    box, tf = _textbox(slide, left=left, top=top, width=width, height=height, anchor=anchor)
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    _run(p, text, size=size, color=color, bold=bold, italic=italic, font=font, tracking=tracking)
    return box


# ---------------------------------------------------------------------------
# shared chrome (kicker / title / headline / footer)
# ---------------------------------------------------------------------------


def _kicker(slide, text, theme: Theme, slide_w):
    _simple_text(
        slide,
        text.upper(),
        left=theme.margin_left,
        top=KICKER_TOP,
        width=slide_w - theme.margin_left - theme.margin_right,
        height=0.3,
        size=theme.size_kicker,
        color=theme.accent,
        bold=True,
        font=theme.font_display,
        tracking=2.2,
    )


def _title(slide, text, theme: Theme, slide_w):
    _simple_text(
        slide,
        text,
        left=theme.margin_left,
        top=TITLE_TOP,
        width=slide_w - theme.margin_left - theme.margin_right,
        height=0.7,
        size=theme.size_title,
        color=theme.ink,
        bold=True,
        font=theme.font_display,
    )
    _rect(
        slide,
        MSO_SHAPE.ROUNDED_RECTANGLE,
        theme.margin_left,
        BAR_TOP,
        theme.accent_bar_w,
        theme.accent_bar_h,
        fill=theme.accent,
    )


def _headline(slide, text, theme: Theme, slide_w):
    _simple_text(
        slide,
        text,
        left=theme.margin_left,
        top=HEADLINE_TOP,
        width=slide_w - theme.margin_left - theme.margin_right,
        height=0.6,
        size=theme.size_headline,
        color=theme.ink_soft,
        font=theme.font_body,
        line_spacing=1.05,
    )


def _footer(slide, theme: Theme, slide_w, slide_h, footer_text, page, total):
    if not footer_text and not page:
        return
    y = slide_h - theme.margin_bottom - 0.05
    _rect(slide, MSO_SHAPE.RECTANGLE, theme.margin_left, y, slide_w - theme.margin_left - theme.margin_right, 0.012, fill=theme.rule)
    if footer_text:
        _simple_text(
            slide,
            footer_text,
            left=theme.margin_left,
            top=y + 0.04,
            width=(slide_w - theme.margin_left - theme.margin_right) * 0.7,
            height=0.3,
            size=theme.size_footer,
            color=theme.ink_soft,
            font=theme.font_body,
        )
    if page:
        _simple_text(
            slide,
            f"{page} / {total}" if total else str(page),
            left=slide_w - theme.margin_right - 1.5,
            top=y + 0.04,
            width=1.5,
            height=0.3,
            size=theme.size_footer,
            color=theme.ink_soft,
            align=PP_ALIGN.RIGHT,
            font=theme.font_body,
        )


def _set_notes(slide, text: Optional[str], source: Optional[str] = None) -> None:
    parts = []
    if text:
        parts.append(text)
    if source:
        parts.append(f"Image source: {source}")
    if not parts:
        return
    slide.notes_slide.notes_text_frame.text = "\n\n".join(parts)


def _content_chrome(slide, model, theme, slide_w, slide_h, footer, page, total):
    """Draw background + kicker + title + accent bar + headline + footer."""
    _bg(slide, theme.background)
    kicker = getattr(model, "kicker", None)
    if kicker:
        _kicker(slide, kicker, theme, slide_w)
    _title(slide, model.title, theme, slide_w)
    if getattr(model, "headline", None):
        _headline(slide, model.headline, theme, slide_w)
    _footer(slide, theme, slide_w, slide_h, footer, page, total)


# ---------------------------------------------------------------------------
# bullet rows (designed markers, not raw dots)
# ---------------------------------------------------------------------------


def _bullet_rows(
    slide,
    items,
    *,
    left,
    top,
    width,
    height,
    theme: Theme,
    numbered=False,
    text_size=None,
    marker_color=None,
):
    text_size = text_size or theme.size_body
    marker_color = marker_color or theme.accent
    n = len(items)
    row_h = min(0.95, height / n)
    marker = 0.26 if numbered else 0.16
    text_left = left + (0.5 if numbered else 0.38)
    for i, item in enumerate(items):
        ry = top + i * row_h
        if numbered:
            chip = _rect(
                slide,
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left,
                ry + 0.04,
                0.34,
                0.34,
                fill=marker_color,
            )
            _, ctf = chip.text_frame, chip.text_frame
            ctf.word_wrap = False
            ctf.margin_left = Emu(0)
            ctf.margin_right = Emu(0)
            ctf.margin_top = Emu(0)
            ctf.margin_bottom = Emu(0)
            ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
            cp = ctf.paragraphs[0]
            cp.alignment = PP_ALIGN.CENTER
            _run(cp, str(i + 1), size=theme.size_caption, color=theme.on_accent, bold=True, font=theme.font_display)
        else:
            _rect(
                slide,
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left,
                ry + (text_size / 72.0) / 2.0,
                marker,
                marker,
                fill=marker_color,
            )
        _simple_text(
            slide,
            item,
            left=text_left,
            top=ry,
            width=width - (text_left - left),
            height=row_h,
            size=text_size,
            color=theme.ink,
            font=theme.font_body,
            anchor=MSO_ANCHOR.TOP,
            line_spacing=1.02,
        )


# ---------------------------------------------------------------------------
# image fitting
# ---------------------------------------------------------------------------


def _image_aspect(path: Path) -> Optional[float]:
    if Image is None:
        return None
    try:
        with Image.open(path) as im:
            w, h = im.size
            return w / h if h else None
    except Exception:  # noqa: BLE001
        return None


def _place_image_contain(slide, path: Path, box, theme: Theme, shadow=True):
    """Place an image fully inside box (inches), preserving aspect ratio, centered."""
    bx, by, bw, bh = box
    ar = _image_aspect(path)
    if ar is None:
        # Unknown size: fill the box and let it be.
        pic = slide.shapes.add_picture(str(path), Inches(bx), Inches(by), Inches(bw), Inches(bh))
    else:
        box_ar = bw / bh
        if ar > box_ar:
            w = bw
            h = bw / ar
        else:
            h = bh
            w = bh * ar
        x = bx + (bw - w) / 2
        y = by + (bh - h) / 2
        pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    pic.line.color.rgb = theme.surface_edge
    pic.line.width = Pt(0.75)
    if shadow:
        _soft_shadow(pic)
    return pic


def _placeholder(slide, box, theme, text, color=None):
    bx, by, bw, bh = box
    _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh, fill=theme.surface, line=theme.surface_edge)
    _simple_text(
        slide,
        text,
        left=bx,
        top=by,
        width=bw,
        height=bh,
        size=theme.size_body,
        color=color or theme.ink_soft,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        font=theme.font_body,
    )


# ---------------------------------------------------------------------------
# renderers
# ---------------------------------------------------------------------------


def render_title(prs, s: TitleSlide, theme: Theme, *, base_dir, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.background)
    # left accent rail
    _rect(slide, MSO_SHAPE.RECTANGLE, 0, 0, 0.28, slide_h, fill=theme.accent)

    _simple_text(
        slide, s.title,
        left=theme.margin_left + 0.3, top=slide_h * 0.30,
        width=slide_w - theme.margin_left - theme.margin_right - 0.3, height=1.8,
        size=theme.size_section, color=theme.ink, bold=True, font=theme.font_display,
        line_spacing=1.0,
    )
    bits = [b for b in (s.subtitle,) if b]
    if bits:
        _simple_text(
            slide, "  ".join(bits),
            left=theme.margin_left + 0.32, top=slide_h * 0.30 + 1.85,
            width=slide_w - theme.margin_left - theme.margin_right - 0.3, height=0.5,
            size=theme.size_headline, color=theme.ink_soft, font=theme.font_body,
        )
    meta_bits = [b for b in (s.presenter, str(s.date) if s.date else None) if b]
    if meta_bits:
        ry = slide_h - theme.margin_bottom - 0.5
        _rect(slide, MSO_SHAPE.RECTANGLE, theme.margin_left + 0.32, ry, 0.5, 0.03, fill=theme.accent)
        _simple_text(
            slide, "   \u00b7   ".join(meta_bits),
            left=theme.margin_left + 0.32, top=ry + 0.08,
            width=slide_w - theme.margin_left - theme.margin_right, height=0.4,
            size=theme.size_attribution, color=theme.ink_soft, font=theme.font_body,
        )
    _set_notes(slide, s.notes)
    return slide


def render_section(prs, s: SectionSlide, theme: Theme, *, section_index=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.ink)
    num = s.number or (f"{section_index:02d}" if section_index is not None else None)
    if num:
        _simple_text(
            slide, num,
            left=theme.margin_left, top=slide_h * 0.18,
            width=6.0, height=2.6,
            size=theme.size_section_num, color=theme.ink_soft, bold=True, font=theme.font_display,
        )
    _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, theme.margin_left + 0.04, slide_h * 0.60, 0.8, 0.09, fill=theme.accent)
    _simple_text(
        slide, s.title,
        left=theme.margin_left, top=slide_h * 0.62,
        width=slide_w - theme.margin_left - theme.margin_right, height=1.6,
        size=theme.size_section, color=theme.on_accent, bold=True, font=theme.font_display,
        line_spacing=1.0,
    )
    _set_notes(slide, s.notes)
    return slide


def render_statement(prs, s: StatementSlide, theme: Theme, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.accent)
    if s.kicker:
        _simple_text(
            slide, s.kicker.upper(),
            left=theme.margin_left + 0.3, top=slide_h * 0.26,
            width=slide_w - 2 * theme.margin_left - 0.6, height=0.4,
            size=theme.size_kicker, color=theme.on_accent_soft, bold=True,
            font=theme.font_display, tracking=2.4, align=PP_ALIGN.CENTER,
        )
    _simple_text(
        slide, s.headline,
        left=theme.margin_left + 0.3, top=slide_h * 0.30,
        width=slide_w - 2 * theme.margin_left - 0.6, height=slide_h * 0.42,
        size=theme.size_statement, color=theme.on_accent, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=theme.font_display,
        line_spacing=1.05,
    )
    _set_notes(slide, s.notes)
    return slide


def render_bullets(prs, s: BulletsSlide, theme: Theme, *, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)
    _bullet_rows(
        slide, s.bullets,
        left=theme.margin_left, top=CONTENT_TOP,
        width=slide_w - theme.margin_left - theme.margin_right,
        height=slide_h - CONTENT_TOP - theme.margin_bottom - 0.3,
        theme=theme, numbered=s.numbered,
    )
    _set_notes(slide, s.notes)
    return slide


def render_comparison(prs, s: ComparisonSlide, theme: Theme, *, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)

    n = len(s.columns)
    content_w = slide_w - theme.margin_left - theme.margin_right
    gap = 0.35
    col_w = (content_w - gap * (n - 1)) / n
    top = CONTENT_TOP
    height = slide_h - top - theme.margin_bottom - 0.3

    for i, col in enumerate(s.columns):
        left = theme.margin_left + i * (col_w + gap)
        card = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, col_w, height, fill=theme.surface, line=theme.surface_edge)
        _soft_shadow(card, alpha=14)
        # header band
        _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, col_w, 0.62, fill=theme.accent if i == 0 else theme.accent2)
        _simple_text(
            slide, col.label,
            left=left + 0.25, top=top, width=col_w - 0.5, height=0.62,
            size=theme.size_body + 2, color=theme.on_accent, bold=True,
            anchor=MSO_ANCHOR.MIDDLE, font=theme.font_display,
        )
        _bullet_rows(
            slide, col.bullets,
            left=left + 0.28, top=top + 0.85,
            width=col_w - 0.56, height=height - 1.1,
            theme=theme, text_size=theme.size_body - 1,
            marker_color=theme.accent if i == 0 else theme.accent2,
        )
    _set_notes(slide, s.notes)
    return slide


def render_code(prs, s: CodeSlide, theme: Theme, *, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)

    top = CONTENT_TOP
    left = theme.margin_left
    width = slide_w - theme.margin_left - theme.margin_right
    height = slide_h - top - theme.margin_bottom - 0.3

    bg_hex = style_background_hex(theme.pygments_style)
    card_bg = RGBColor.from_string(bg_hex) if bg_hex else theme.code_bg
    card = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, fill=card_bg)
    _soft_shadow(card, alpha=30)

    # language pill
    if s.language and s.language != "text":
        pill = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left + width - 1.4, top + 0.18, 1.2, 0.34, fill=theme.accent)
        ptf = pill.text_frame
        ptf.margin_left = Emu(0); ptf.margin_right = Emu(0); ptf.margin_top = Emu(0); ptf.margin_bottom = Emu(0)
        ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        pp = ptf.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
        _run(pp, s.language.lower(), size=theme.size_caption, color=theme.on_accent, bold=True, font=theme.font_display)

    spans_per_line = highlight_lines(s.code, s.language, theme.pygments_style, str(theme.code_default))
    n = max(1, len(spans_per_line))
    pad = 0.28
    avail_h = height - 2 * pad
    line_h = min(0.34, avail_h / n)
    font_pt = min(theme.size_code, max(9, int(line_h * 72 * 0.62)))
    highlight = set(s.highlight_lines or [])

    text_left = left + pad + 0.15
    for idx, spans in enumerate(spans_per_line, start=1):
        ly = top + pad + (idx - 1) * line_h
        if idx in highlight:
            _rect(slide, MSO_SHAPE.RECTANGLE, left + 0.08, ly - 0.01, width - 0.16, line_h, fill=theme.code_line_highlight)
            _rect(slide, MSO_SHAPE.RECTANGLE, left + 0.08, ly - 0.01, 0.05, line_h, fill=theme.accent)
        box, tf = _textbox(slide, left=text_left, top=ly, width=width - pad - 0.3, height=line_h, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        if not spans:
            _run(p, " ", size=font_pt, color=card_bg, font=theme.font_mono)
        for sp in spans:
            _run(p, sp.text, size=font_pt, color=sp.color, bold=sp.bold, italic=sp.italic, font=theme.font_mono)
    _set_notes(slide, s.notes)
    return slide


def render_diagram(prs, s: DiagramSlide, theme: Theme, *, base_dir, mermaid_renderer=None, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)

    top = CONTENT_TOP
    box_w = slide_w - theme.margin_left - theme.margin_right
    box_h = slide_h - top - theme.margin_bottom - 0.3
    box = (theme.margin_left, top, box_w, box_h)

    if mermaid_renderer is None:
        _placeholder(slide, box, theme, "[diagram rendering disabled \u2014 install mermaid]")
        _set_notes(slide, s.notes)
        return slide

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "diagram.png"
        try:
            mermaid_renderer(s.mermaid, out)
            _place_image_contain(slide, out, box, theme, shadow=False)
        except Exception as e:  # noqa: BLE001
            _placeholder(slide, box, theme, f"[diagram failed: {e}]", color=RGBColor.from_string("C0392B"))
    _set_notes(slide, s.notes)
    return slide


def render_image(prs, s: ImageSlide, theme: Theme, *, base_dir, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)

    img_path = Path(s.path)
    if not img_path.is_absolute():
        img_path = (base_dir / s.path).resolve()

    top = CONTENT_TOP
    full_w = slide_w - theme.margin_left - theme.margin_right
    full_h = slide_h - top - theme.margin_bottom - 0.3
    cap_h = 0.4 if s.caption else 0.0

    layout = s.layout
    if layout == "auto":
        layout = "right" if s.bullets else "full"

    if layout in ("right", "left") and s.bullets:
        gap = 0.4
        col_w = (full_w - gap) / 2
        if layout == "right":
            text_x, img_x = theme.margin_left, theme.margin_left + col_w + gap
        else:
            img_x, text_x = theme.margin_left, theme.margin_left + col_w + gap
        _bullet_rows(
            slide, s.bullets,
            left=text_x, top=top + 0.1, width=col_w, height=full_h - 0.2,
            theme=theme,
        )
        img_box = (img_x, top, col_w, full_h - cap_h)
        if img_path.exists():
            _place_image_contain(slide, img_path, img_box, theme)
        else:
            _placeholder(slide, img_box, theme, f"[missing image: {s.path}]", color=RGBColor.from_string("C0392B"))
        if s.caption:
            _simple_text(slide, s.caption, left=img_x, top=top + full_h - cap_h + 0.05, width=col_w, height=cap_h, size=theme.size_caption, color=theme.ink_soft, align=PP_ALIGN.CENTER, font=theme.font_body)
    else:
        img_box = (theme.margin_left, top, full_w, full_h - cap_h)
        if img_path.exists():
            _place_image_contain(slide, img_path, img_box, theme)
        else:
            _placeholder(slide, img_box, theme, f"[missing image: {s.path}]", color=RGBColor.from_string("C0392B"))
        if s.caption:
            _simple_text(slide, s.caption, left=theme.margin_left, top=top + full_h - cap_h + 0.05, width=full_w, height=cap_h, size=theme.size_caption, color=theme.ink_soft, align=PP_ALIGN.CENTER, font=theme.font_body)

    _set_notes(slide, s.notes, source=s.source)
    return slide


def render_table(prs, s: TableSlide, theme: Theme, *, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _content_chrome(slide, s, theme, slide_w, slide_h, footer, page, total)

    top = CONTENT_TOP
    width = slide_w - theme.margin_left - theme.margin_right
    n_rows = len(s.rows) + 1
    row_h = min(0.55, (slide_h - top - theme.margin_bottom - 0.3) / n_rows)
    height = row_h * n_rows

    shape = slide.shapes.add_table(
        n_rows, len(s.columns),
        Inches(theme.margin_left), Inches(top), Inches(width), Inches(height),
    )
    table = shape.table
    table.first_row = False
    table.horz_banding = False

    for j, name in enumerate(s.columns):
        c = table.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = theme.table_header_bg
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Inches(0.12); c.margin_right = Inches(0.12)
        para = c.text_frame.paragraphs[0]
        _run(para, str(name), size=theme.size_body - 1, color=theme.table_header_fg, bold=True, font=theme.font_display)

    for i, row in enumerate(s.rows, start=1):
        zebra = theme.table_zebra if i % 2 == 1 else theme.background
        for j, val in enumerate(row):
            c = table.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = zebra
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = Inches(0.12); c.margin_right = Inches(0.12)
            para = c.text_frame.paragraphs[0]
            _run(para, str(val), size=theme.size_body - 2, color=theme.ink, font=theme.font_body)
    _set_notes(slide, s.notes)
    return slide


def render_quote(prs, s: QuoteSlide, theme: Theme, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.surface)
    _simple_text(
        slide, "\u201C",
        left=theme.margin_left, top=slide_h * 0.10,
        width=2.0, height=2.0,
        size=theme.size_quote_mark, color=theme.accent, bold=True, font=theme.font_display,
    )
    _simple_text(
        slide, s.quote,
        left=theme.margin_left + 0.5, top=slide_h * 0.34,
        width=slide_w - 2 * theme.margin_left - 1.0, height=slide_h * 0.34,
        size=theme.size_quote, color=theme.ink, bold=False, font=theme.font_display,
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1,
    )
    if s.attribution:
        ry = slide_h * 0.74
        _rect(slide, MSO_SHAPE.RECTANGLE, theme.margin_left + 0.55, ry, 0.5, 0.03, fill=theme.accent)
        _simple_text(
            slide, s.attribution,
            left=theme.margin_left + 0.55, top=ry + 0.1,
            width=slide_w - 2 * theme.margin_left, height=0.5,
            size=theme.size_attribution, color=theme.ink_soft, font=theme.font_body,
        )
    _set_notes(slide, s.notes)
    return slide


def render_summary(prs, s: SummarySlide, theme: Theme, *, footer=None, page=None, total=None, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.background)
    _kicker(slide, "Key takeaways", theme, slide_w)
    _title(slide, s.title, theme, slide_w)
    _footer(slide, theme, slide_w, slide_h, footer, page, total)

    top = CONTENT_TOP
    width = slide_w - theme.margin_left - theme.margin_right
    avail = slide_h - top - theme.margin_bottom - 0.3
    n = len(s.points)
    row_h = min(1.2, avail / n)
    for i, point in enumerate(s.points):
        ry = top + i * row_h
        _simple_text(
            slide, f"{i+1:02d}",
            left=theme.margin_left, top=ry,
            width=1.0, height=row_h,
            size=theme.size_section, color=theme.accent, bold=True, font=theme.font_display,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        _simple_text(
            slide, point,
            left=theme.margin_left + 1.15, top=ry,
            width=width - 1.15, height=row_h,
            size=theme.size_body_lg, color=theme.ink, font=theme.font_body,
            anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.02,
        )
        if i < n - 1:
            _rect(slide, MSO_SHAPE.RECTANGLE, theme.margin_left + 1.15, ry + row_h - 0.02, width - 1.15, 0.01, fill=theme.rule)
    _set_notes(slide, s.notes)
    return slide


def render_qa(prs, s: QASlide, theme: Theme, **_):
    slide = _new_slide(prs)
    slide_w, slide_h = _slide_size_in(prs)
    _bg(slide, theme.accent)
    _simple_text(
        slide, s.title,
        left=theme.margin_left, top=slide_h * 0.34,
        width=slide_w - 2 * theme.margin_left, height=1.6,
        size=theme.size_qa, color=theme.on_accent, bold=True,
        align=PP_ALIGN.CENTER, font=theme.font_display, anchor=MSO_ANCHOR.MIDDLE,
    )
    if s.contact:
        pill = _rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, slide_w / 2 - 2.0, slide_h * 0.62, 4.0, 0.5, fill=theme.on_accent)
        ptf = pill.text_frame
        ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        pp = ptf.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
        _run(pp, s.contact, size=theme.size_attribution, color=theme.accent, bold=True, font=theme.font_body)
    _set_notes(slide, s.notes)
    return slide


# ---------------------------------------------------------------------------
# dispatch
# ---------------------------------------------------------------------------


RENDERERS = {
    "title": render_title,
    "section": render_section,
    "statement": render_statement,
    "bullets": render_bullets,
    "comparison": render_comparison,
    "code": render_code,
    "diagram": render_diagram,
    "image": render_image,
    "table": render_table,
    "quote": render_quote,
    "summary": render_summary,
    "qa": render_qa,
}


def render_slide(
    prs,
    slide_model: Slide,
    theme: Theme,
    *,
    base_dir: Path,
    mermaid_renderer=None,
    footer=None,
    page=None,
    total=None,
    section_index=None,
):
    fn = RENDERERS[slide_model.type]
    return fn(
        prs,
        slide_model,
        theme,
        base_dir=base_dir,
        mermaid_renderer=mermaid_renderer,
        footer=footer,
        page=page,
        total=total,
        section_index=section_index,
    )
