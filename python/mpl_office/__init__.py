"""mpl_office — native vector graphics from Python into Microsoft Office documents.

Low-level API::

    from mpl_office import convert_svg_to_drawingml, ConvertOptions

    opts = ConvertOptions(
        source_dpi=72,
        target_width_emu=6858000,
        target_height_emu=4572000,
    )
    xml = convert_svg_to_drawingml(svg_string, opts)

High-level PowerPoint / Word integration lives in :mod:`mpl_office.pptx`
and :mod:`mpl_office.docx`. A matplotlib backend is registered from
:mod:`mpl_office.backend`.
"""

from ._native import (
    ConvertOptions,
    convert_svg_to_drawingml,
    convert_svg_to_drawingml_with_images,
)

__all__ = [
    "ConvertOptions",
    "convert_svg_to_drawingml",
    "convert_svg_to_drawingml_with_images",
    "__version__",
]

__version__ = "0.1.0"
