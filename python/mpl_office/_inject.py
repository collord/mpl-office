"""Helpers for injecting DrawingML XML into `python-pptx` / `python-docx`
document object models.

Both libraries use `lxml` internally. Our Rust core emits raw XML fragments
with `p:` and `a:` prefixes but no namespace declarations, so we have to
parse them inside a wrapper element that declares the namespaces.
"""
from __future__ import annotations

from typing import List

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
