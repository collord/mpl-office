"""Smoke tests — verify the native extension loads and the basic pipeline
turns a trivial SVG into DrawingML XML.
"""
from __future__ import annotations

import mpl_office


def test_module_exports():
    assert callable(mpl_office.convert_svg_to_drawingml)
    assert hasattr(mpl_office, "ConvertOptions")


def test_default_options():
    opts = mpl_office.ConvertOptions()
    assert opts.source_dpi == 96.0
    assert opts.target_width_emu is None
    assert opts.offset_x_emu == 0


def test_convert_trivial_rect():
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
        '<rect x="0" y="0" width="50" height="30" fill="#336699"/>'
        '</svg>'
    )
    out = mpl_office.convert_svg_to_drawingml(svg)
    assert "<p:sp>" in out
    assert 'prst="rect"' in out
    assert "336699" in out


def test_convert_with_options():
    svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100"/></svg>'
    opts = mpl_office.ConvertOptions(
        target_width_emu=6_858_000,
        target_height_emu=4_572_000,
    )
    out = mpl_office.convert_svg_to_drawingml(svg, opts)
    assert 'cx="6858000"' in out
    assert 'cy="4572000"' in out
