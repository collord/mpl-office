"""PowerPoint integration — insert matplotlib figures and arbitrary SVG into
`python-pptx` slides as native, editable DrawingML vector shapes.
"""
from __future__ import annotations

from io import StringIO
from pathlib import Path
from typing import Any, Optional, Union

from . import ConvertOptions, convert_svg_to_drawingml_with_images
from ._inject import append_to_sptree_with_images

# Public re-exports
__all__ = [
    "fig_to_slide",
    "fig_to_placeholder",
    "svg_to_slide",
    "Inches",
    "Emu",
]

EMU_PER_INCH = 914_400
EMU_PER_POINT = 12_700


def Inches(v: float) -> int:
    """Inches → EMU, matching python-pptx's convention."""
    return int(round(v * EMU_PER_INCH))


def Emu(v: int) -> int:
    return int(v)


def _render_fig_svg(fig, *, dpi: Optional[float] = None) -> tuple[str, float, float]:
    """Render a matplotlib Figure to an in-memory SVG string.

    Forces ``svg.fonttype='none'`` so text comes through as ``<text>``
    elements rather than glyph paths — this is what makes the resulting
    PowerPoint output *editable*.

    Returns ``(svg, width_px, height_px)`` where pixel dimensions are at the
    SVG backend's assumed 72 DPI.
    """
    import matplotlib as mpl

    buf = StringIO()
    # matplotlib's SVG backend always writes at 72 DPI.
    with mpl.rc_context({"svg.fonttype": "none"}):
        fig.savefig(buf, format="svg", bbox_inches=None)
    svg = buf.getvalue()
    w_in, h_in = fig.get_size_inches()
    return svg, w_in * 72.0, h_in * 72.0


def svg_to_slide(
    svg: str,
    slide,
    *,
    left: int = 0,
    top: int = 0,
    width: Optional[int] = None,
    height: Optional[int] = None,
    source_dpi: float = 96.0,
):
    """Inject an SVG string into a `python-pptx` slide.

    `left`, `top`, `width`, `height` are in EMU (use :func:`Inches`).
    If `width` or `height` is None the shape uses its natural size at
    `source_dpi`.
    """
    opts = ConvertOptions(
        source_dpi=source_dpi,
        target_width_emu=width,
        target_height_emu=height,
        offset_x_emu=left,
        offset_y_emu=top,
    )
    xml, images = convert_svg_to_drawingml_with_images(svg, opts)
    return append_to_sptree_with_images(slide, xml, images)


def fig_to_slide(
    fig,
    slide,
    *,
    left: int = 0,
    top: int = 0,
    width: Optional[int] = None,
    height: Optional[int] = None,
):
    """Render a matplotlib figure onto a slide as native DrawingML shapes.

    When `width`/`height` is omitted the figure is placed at its natural
    inch dimensions derived from ``fig.get_size_inches()``.
    """
    svg, w_px, h_px = _render_fig_svg(fig)

    if width is None:
        width = Inches(w_px / 72.0)
    if height is None:
        height = Inches(h_px / 72.0)

    return svg_to_slide(
        svg,
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        source_dpi=72.0,
    )


def fig_to_placeholder(fig, slide, placeholder, *, remove_placeholder: bool = True):
    """Insert a figure into a placeholder, matching its position and size.

    `slide` must be the slide that owns the placeholder. By default the
    placeholder element is removed from the slide after its bounds are
    captured, so the figure replaces it cleanly.
    """
    left = placeholder.left
    top = placeholder.top
    width = placeholder.width
    height = placeholder.height

    if remove_placeholder:
        sp = placeholder._element
        sp.getparent().remove(sp)

    return fig_to_slide(fig, slide, left=left, top=top, width=width, height=height)
