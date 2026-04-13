"""Helpers for injecting DrawingML XML into `python-pptx` / `python-docx`
document object models.

Both libraries use `lxml` internally. Our Rust core emits raw XML fragments
with `p:` and `a:` prefixes but no namespace declarations, so we have to
parse them inside a wrapper element that declares the namespaces.
"""
from __future__ import annotations

from io import BytesIO
from typing import Iterable, List, Sequence, Tuple

from lxml import etree

# OOXML namespace constants
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

NS_MAP = {
    "a": NS_A,
    "p": NS_P,
    "r": NS_R,
    "w": NS_W,
    "wp": NS_WP,
    "wps": NS_WPS,
    "wpg": NS_WPG,
    "pic": NS_PIC,
}


def parse_drawingml_fragment(xml: str) -> List[etree._Element]:
    """Parse a DrawingML XML fragment (as emitted by the Rust core) into a
    list of top-level elements. The fragment is wrapped in a disposable root
    that declares the OOXML namespaces so lxml can resolve the prefixes.
    """
    nsdecl = " ".join(f'xmlns:{k}="{v}"' for k, v in NS_MAP.items())
    wrapped = f"<root {nsdecl}>{xml}</root>"
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    root = etree.fromstring(wrapped.encode("utf-8"), parser)
    return list(root)


def append_to_sptree(slide, xml: str) -> List[etree._Element]:
    """Append DrawingML shapes to a `python-pptx` slide's `<p:spTree>`.

    Returns the appended lxml elements.
    """
    sp_tree = slide.shapes._spTree
    elements = parse_drawingml_fragment(xml)
    for el in elements:
        sp_tree.append(el)
    return elements


ImageTuple = Tuple[str, bytes, str]
"""One entry returned by ``convert_svg_to_drawingml_with_images``:
``(sentinel, image_bytes, format)``."""


def register_images_on_slide(
    slide, images: Sequence[ImageTuple]
) -> dict[str, str]:
    """Register each image as an OOXML image part on the given slide.

    Returns a mapping from the Rust core's sentinel id to the real
    relationship id (``rId``) that ``python-pptx`` allocated. Duplicate
    image bytes are deduplicated by ``python-pptx``'s image part cache
    automatically — two sentinels may therefore map to the same ``rId``.
    """
    sentinel_to_rid: dict[str, str] = {}
    for sentinel, data, _format in images:
        _image_part, rid = slide.part.get_or_add_image_part(BytesIO(data))
        sentinel_to_rid[sentinel] = rid
    return sentinel_to_rid


def rewrite_sentinels(xml: str, mapping: dict[str, str]) -> str:
    """Replace each ``r:embed="{sentinel}"`` with the real relationship id.

    The Rust emitter writes sentinels like ``__mpl_office_img_0__`` in
    `r:embed` attributes. We rewrite them by literal string replacement
    — the sentinels are designed to be unique and never clash with any
    other XML content.
    """
    for sentinel, rid in mapping.items():
        needle = f'r:embed="{sentinel}"'
        replacement = f'r:embed="{rid}"'
        xml = xml.replace(needle, replacement)
    return xml


def append_to_sptree_with_images(
    slide, xml: str, images: Sequence[ImageTuple]
) -> List[etree._Element]:
    """Inject DrawingML shapes *and* embedded images into a slide.

    This is the image-aware counterpart to :func:`append_to_sptree`:

    1. Each image blob is registered as an image part on ``slide`` via
       ``python-pptx``, producing a real relationship id.
    2. The sentinel rIds in the XML are rewritten to the real ones.
    3. The resulting XML is parsed and appended to ``slide.shapes._spTree``.
    """
    if images:
        mapping = register_images_on_slide(slide, images)
        xml = rewrite_sentinels(xml, mapping)
    return append_to_sptree(slide, xml)
