"""Verify `fig.savefig("out.pptx"/"out.docx")` works via the mpl_office backend."""
from __future__ import annotations

from pathlib import Path

import pytest

matplotlib = pytest.importorskip("matplotlib")
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

pptx = pytest.importorskip("pptx")
from pptx import Presentation  # noqa: E402

docx = pytest.importorskip("docx")
from docx import Document  # noqa: E402


def test_savefig_pptx(artifacts: Path):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot([1, 2, 3], [1, 4, 9])
    ax.set_title("Backend test")
    out = artifacts / "backend_savefig.pptx"
    fig.savefig(out, backend="module://mpl_office.backend")
    plt.close(fig)

    assert out.exists() and out.stat().st_size > 1000
    prs = Presentation(out)
    assert len(prs.slides) == 1
    # Should contain at least one (grouped) shape
    shapes = list(prs.slides[0].shapes)
    assert len(shapes) >= 1


def _document_xml(path: Path):
    """Read `word/document.xml` out of a saved .docx as an element tree."""
    import zipfile

    from lxml import etree

    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def test_savefig_docx(artifacts: Path):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot([1, 2, 3], [1, 4, 9])
    ax.set_title("Backend test")
    out = artifacts / "backend_savefig_test.docx"
    fig.savefig(out, backend="module://mpl_office.backend")
    plt.close(fig)

    assert out.exists() and out.stat().st_size > 1000
    # python-docx must be able to reopen it.
    assert len(Document(out).paragraphs) >= 1

    root = _document_xml(out)
    drawings = root.findall(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"
    )
    assert len(drawings) == 1


def test_savefig_docx_width_inches_preserves_aspect(artifacts: Path):
    fig, ax = plt.subplots(figsize=(8, 4))  # 2:1
    ax.plot([1, 2, 3], [1, 4, 9])
    out = artifacts / "backend_savefig_width.docx"
    fig.savefig(out, backend="module://mpl_office.backend", width_inches=6)
    plt.close(fig)

    root = _document_xml(out)
    extent = root.find(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/"
        "wordprocessingDrawing}extent"
    )
    cx, cy = int(extent.get("cx")), int(extent.get("cy"))
    assert cx == 6 * 914_400
    # The 2:1 figure aspect must survive the rescale.
    assert abs(cx / cy - 2.0) < 0.01


@pytest.mark.parametrize(
    ("suffix", "part", "needle"),
    [
        (".pptx", "ppt/slides/slide1.xml", b"<a:t>"),
        (".docx", "word/document.xml", b"<w:t"),
    ],
)
def test_savefig_keeps_text_editable(artifacts: Path, suffix, part, needle):
    """Text must land as real text, not glyph outlines.

    matplotlib's SVG backend defaults to ``svg.fonttype='path'``, which
    converts every glyph to a filled outline. The Office backend has to
    override that -- otherwise titles and tick labels arrive as
    unselectable vector blobs and the whole premise of the library fails.
    """
    import zipfile

    fig, ax = plt.subplots(figsize=(5, 3))
    ax.plot([1, 2, 3], [1, 4, 9])
    ax.set_title("Unique Title Text")
    out = artifacts / f"editable_text{suffix}"
    fig.savefig(out, backend="module://mpl_office.backend")
    plt.close(fig)

    with zipfile.ZipFile(out) as z:
        blob = z.read(part)
    assert needle in blob
    assert b"Unique Title Text" in blob
