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


# ---------------------------------------------------------------------------
# Word (.docx) support
#
# The Rust core emits PresentationML-flavoured shapes (`p:sp`, `p:grpSp`,
# `p:pic`). Word expresses the same DrawingML geometry with a different
# vocabulary: shapes are `wps:wsp`, groups are `wpg:wgp`, pictures are
# `pic:pic`, and the tree hangs off a `w:drawing` / `wp:inline` anchor rather
# than a `p:spTree`. Crucially the `a:` geometry, fill, stroke and text-run
# elements are identical in both worlds — that is what makes a mechanical
# rewrite viable instead of a second emitter in the Rust core.
# ---------------------------------------------------------------------------

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _qa(tag: str) -> str:
    return "{%s}%s" % (NS_A, tag)


def _qp(tag: str) -> str:
    return "{%s}%s" % (NS_P, tag)


def _qw(tag: str) -> str:
    return "{%s}%s" % (NS_W, tag)


def _qwps(tag: str) -> str:
    return "{%s}%s" % (NS_WPS, tag)


def _qwpg(tag: str) -> str:
    return "{%s}%s" % (NS_WPG, tag)


def _qpic(tag: str) -> str:
    return "{%s}%s" % (NS_PIC, tag)


#: Straight one-for-one element renames, keyed by PresentationML tag.
_RENAME = {
    _qp("spPr"): _qwps("spPr"),
    _qp("cNvSpPr"): _qwps("cNvSpPr"),
    _qp("grpSpPr"): _qwpg("grpSpPr"),
    _qp("cNvGrpSpPr"): _qwpg("cNvGrpSpPr"),
    _qp("blipFill"): _qpic("blipFill"),
    _qp("cNvPicPr"): _qpic("cNvPicPr"),
    _qp("nvPicPr"): _qpic("nvPicPr"),
}

#: Wrapper elements Word has no use for; their children are promoted.
_UNWRAP = {_qp("nvSpPr"), _qp("nvGrpSpPr")}

#: Elements with no Word counterpart, dropped outright.
_DROP = {_qp("nvPr")}

#: DrawingML alignment -> WordprocessingML justification.
_ALGN_TO_JC = {"l": "left", "ctr": "center", "r": "right", "just": "both"}


def _convert_shape_element(el: etree._Element) -> List[etree._Element]:
    """Rewrite one PresentationML shape element into its Word equivalent.

    Returns the elements that replace `el` in its parent — normally one, but
    `_UNWRAP` tags return their converted children instead and `_DROP` tags
    return nothing. Children are converted depth-first, before the parent is
    restructured.
    """
    tag = el.tag

    if tag in _DROP:
        return []

    new_children: List[etree._Element] = []
    for child in list(el):
        el.remove(child)
        new_children.extend(_convert_shape_element(child))

    if tag in _UNWRAP:
        return new_children

    if tag == _qp("txBody"):
        return _convert_txbody(new_children)

    if tag == _qp("sp"):
        el.tag = _qwps("wsp")
    elif tag == _qp("grpSp"):
        el.tag = _qwpg("wgp")
    elif tag == _qp("pic"):
        el.tag = _qpic("pic")
    elif tag in _RENAME:
        el.tag = _RENAME[tag]
    # `p:cNvPr` is left alone here: the right namespace depends on the parent,
    # which `_reparent_cnvpr` resolves once the tree is reassembled.

    for child in new_children:
        el.append(child)

    # Every `wps:wsp` must end with a `wps:bodyPr` — the schema requires it
    # even for shapes that carry no text, and Word refuses to open the file
    # when it is missing. Shapes converted from a `p:txBody` already got one
    # from `_convert_txbody`.
    if el.tag == _qwps("wsp") and el.find(_qwps("bodyPr")) is None:
        etree.SubElement(el, _qwps("bodyPr"))

    return [el]


def _convert_txbody(children: List[etree._Element]) -> List[etree._Element]:
    """Turn a `<p:txBody>`'s children into Word's `wps:txbx` + `wps:bodyPr`.

    PowerPoint stores shape text as DrawingML paragraphs (`a:p`) inside
    `p:txBody`. Word nests *WordprocessingML* paragraphs inside
    `wps:txbx/w:txbxContent`, and carries the body properties in a sibling
    `wps:bodyPr` that must follow the txbx. Returning both elements keeps the
    caller's splice-into-parent logic uniform.
    """
    txbx = etree.Element(_qwps("txbx"))
    content = etree.SubElement(txbx, _qw("txbxContent"))

    source_body_pr = None
    for child in children:
        if child.tag == _qa("bodyPr"):
            source_body_pr = child
        elif child.tag == _qa("p"):
            content.append(_a_p_to_w_p(child))
        # `a:lstStyle` has no Word counterpart and is dropped.

    if len(content) == 0:
        # Word rejects an empty txbxContent; give it one empty paragraph.
        etree.SubElement(content, _qw("p"))

    body_pr = etree.Element(_qwps("bodyPr"))
    if source_body_pr is not None:
        for attr in ("wrap", "lIns", "tIns", "rIns", "bIns", "anchor", "anchorCtr"):
            value = source_body_pr.get(attr)
            if value is not None:
                body_pr.set(attr, value)
        if source_body_pr.find(_qa("spAutoFit")) is not None:
            etree.SubElement(body_pr, _qa("spAutoFit"))

    return [txbx, body_pr]


def _a_p_to_w_p(a_p: etree._Element) -> etree._Element:
    """Convert one DrawingML paragraph into a WordprocessingML paragraph."""
    w_p = etree.Element(_qw("p"))

    a_ppr = a_p.find(_qa("pPr"))
    algn = a_ppr.get("algn") if a_ppr is not None else None
    if algn:
        w_ppr = etree.SubElement(w_p, _qw("pPr"))
        jc = etree.SubElement(w_ppr, _qw("jc"))
        jc.set(_qw("val"), _ALGN_TO_JC.get(algn, "left"))

    for a_r in a_p.findall(_qa("r")):
        w_p.append(_a_r_to_w_r(a_r))
    return w_p


def _a_r_to_w_r(a_r: etree._Element) -> etree._Element:
    """Convert one DrawingML text run into a WordprocessingML run.

    Maps the properties the emitter actually writes: typeface, size, bold,
    italic and solid fill colour.
    """
    w_r = etree.Element(_qw("r"))
    a_rpr = a_r.find(_qa("rPr"))

    if a_rpr is not None:
        w_rpr = etree.SubElement(w_r, _qw("rPr"))

        latin = a_rpr.find(_qa("latin"))
        typeface = latin.get("typeface") if latin is not None else None
        if typeface:
            rfonts = etree.SubElement(w_rpr, _qw("rFonts"))
            for attr in ("ascii", "hAnsi", "cs"):
                rfonts.set(_qw(attr), typeface)

        # `w:rPr`'s children are an ordered sequence in the schema (CT_RPr):
        # rFonts, b, i, color, sz, szCs. Word is strict about this, so the
        # elements are appended in exactly that order rather than the order
        # they happen to appear in the DrawingML source.
        if a_rpr.get("b") == "1":
            etree.SubElement(w_rpr, _qw("b"))
        if a_rpr.get("i") == "1":
            etree.SubElement(w_rpr, _qw("i"))

        srgb = a_rpr.find("%s/%s" % (_qa("solidFill"), _qa("srgbClr")))
        if srgb is not None and srgb.get("val"):
            etree.SubElement(w_rpr, _qw("color")).set(_qw("val"), srgb.get("val"))

        # `a:rPr/@sz` is hundredths of a point; `w:sz` is half-points.
        sz = a_rpr.get("sz")
        if sz:
            try:
                half_points = max(1, int(round(int(sz) / 50.0)))
            except ValueError:
                half_points = None
            if half_points is not None:
                for tag in ("sz", "szCs"):
                    etree.SubElement(w_rpr, _qw(tag)).set(_qw("val"), str(half_points))

    a_t = a_r.find(_qa("t"))
    w_t = etree.SubElement(w_r, _qw("t"))
    w_t.text = a_t.text if a_t is not None else ""
    # Tick labels such as " 2.5" carry meaningful leading whitespace.
    w_t.set(XML_SPACE, "preserve")
    return w_r


def _reparent_cnvpr(root: etree._Element) -> None:
    """Re-namespace every `cNvPr` to match the vocabulary of its parent.

    `wps:wsp` wants `wps:cNvPr`, `wpg:wgp` wants `wpg:cNvPr`, and
    `pic:nvPicPr` wants `pic:cNvPr`. This runs as a post-pass because the
    correct namespace depends on the parent, which is only settled once the
    tree has been reassembled.
    """
    owner = {
        _qwps("wsp"): _qwps("cNvPr"),
        _qwpg("wgp"): _qwpg("cNvPr"),
        _qpic("nvPicPr"): _qpic("cNvPr"),
    }
    for parent in root.iter():
        target = owner.get(parent.tag)
        if target is None:
            continue
        for child in parent:
            if child.tag == _qp("cNvPr"):
                child.tag = target


def presentationml_to_wordml(xml: str) -> List[etree._Element]:
    """Rewrite a PresentationML shape fragment into Word shape elements.

    Takes the raw fragment emitted by the Rust core and returns top-level
    `wps:wsp` / `wpg:wgp` / `pic:pic` elements, ready to be placed inside the
    `wpg:wgp` group of a `w:drawing` anchor.
    """
    converted: List[etree._Element] = []
    for el in parse_drawingml_fragment(xml):
        converted.extend(_convert_shape_element(el))
    for el in converted:
        _reparent_cnvpr(el)
    return converted


def _group_bounds(elements: Sequence[etree._Element]) -> Tuple[int, int, int, int]:
    """Union of the `a:off`/`a:ext` boxes of `elements`, in EMU.

    Returns `(x, y, cx, cy)`. Shapes without an `a:xfrm` (rare, but possible
    for degenerate geometry) are skipped; if nothing has one we fall back to a
    1x1 box so Word still gets a valid extent.
    """
    xs, ys, x2s, y2s = [], [], [], []
    for el in elements:
        for xfrm in el.iter(_qa("xfrm")):
            off = xfrm.find(_qa("off"))
            ext = xfrm.find(_qa("ext"))
            if off is None or ext is None:
                continue
            try:
                x, y = int(off.get("x", 0)), int(off.get("y", 0))
                cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
            except (TypeError, ValueError):
                continue
            xs.append(x)
            ys.append(y)
            x2s.append(x + cx)
            y2s.append(y + cy)
            break  # only the shape's own (outermost) xfrm counts

    if not xs:
        return 0, 0, 1, 1
    x, y = min(xs), min(ys)
    return x, y, max(1, max(x2s) - x), max(1, max(y2s) - y)


#: Wrap styles accepted by `build_word_drawing(anchor="floating")`, mapped to
#: the `wp:` element that expresses them. `square` and `tight` additionally
#: need a `wp:wrapPolygon`-free `wrapText` attribute, handled below.
_WRAP_ELEMENTS = {
    "square": "wrapSquare",
    "tight": "wrapTight",
    "through": "wrapThrough",
    "topAndBottom": "wrapTopAndBottom",
    "none": "wrapNone",
}

#: Anchor positions that Word resolves relative to a page feature rather than
#: an absolute offset, e.g. `position=("center", "paragraph")`.
_ALIGN_KEYWORDS = {"left", "right", "center", "top", "bottom", "inside", "outside"}


def _append_graphic(parent: etree._Element, shapes, width_emu, height_emu, name, shape_id):
    """Append the `a:graphic` payload shared by inline and floating anchors.

    This is the part that is identical either way: a single `wpg:wgp` group
    holding every converted shape, scaled from its own union bounds onto the
    requested on-page box.
    """
    ch_x, ch_y, ch_cx, ch_cy = _group_bounds(shapes)

    graphic = etree.SubElement(parent, _qa("graphic"), nsmap={"a": NS_A})
    graphic_data = etree.SubElement(graphic, _qa("graphicData"))
    graphic_data.set("uri", NS_WPG)

    wgp = etree.SubElement(graphic_data, _qwpg("wgp"), nsmap={"wpg": NS_WPG})
    # Word needs the group to carry its own non-visual identity; without a
    # `wpg:cNvPr` it treats the whole drawing as a locked group that cannot be
    # ungrouped or drilled into, which would defeat the point of emitting
    # native shapes at all.
    grp_cnv_pr = etree.SubElement(wgp, _qwpg("cNvPr"))
    grp_cnv_pr.set("id", str(shape_id))
    grp_cnv_pr.set("name", name)
    etree.SubElement(wgp, _qwpg("cNvGrpSpPr"))

    grp_sp_pr = etree.SubElement(wgp, _qwpg("grpSpPr"))
    xfrm = etree.SubElement(grp_sp_pr, _qa("xfrm"))
    etree.SubElement(xfrm, _qa("off")).attrib.update({"x": "0", "y": "0"})
    etree.SubElement(xfrm, _qa("ext")).attrib.update(
        {"cx": str(max(1, int(width_emu))), "cy": str(max(1, int(height_emu)))}
    )
    etree.SubElement(xfrm, _qa("chOff")).attrib.update({"x": str(ch_x), "y": str(ch_y)})
    etree.SubElement(xfrm, _qa("chExt")).attrib.update(
        {"cx": str(ch_cx), "cy": str(ch_cy)}
    )

    for shape in shapes:
        wgp.append(shape)


def _append_position(anchor: etree._Element, axis: str, value, relative_to: str) -> None:
    """Write a `wp:positionH` / `wp:positionV` child onto a floating anchor.

    `value` is either an EMU offset or one of `_ALIGN_KEYWORDS` (``"center"``,
    ``"right"``, …), which Word resolves against `relative_to` (``"page"``,
    ``"margin"``, ``"column"``, ``"paragraph"``, …).
    """
    tag = "positionH" if axis == "h" else "positionV"
    pos = etree.SubElement(anchor, "{%s}%s" % (NS_WP, tag))
    pos.set("relativeFrom", relative_to)
    if isinstance(value, str):
        if value not in _ALIGN_KEYWORDS:
            raise ValueError(
                "position must be an EMU offset or one of %s, got %r"
                % (sorted(_ALIGN_KEYWORDS), value)
            )
        etree.SubElement(pos, "{%s}align" % NS_WP).text = value
    else:
        etree.SubElement(pos, "{%s}posOffset" % NS_WP).text = str(int(value))


def build_word_drawing(
    xml: str,
    *,
    width_emu: int,
    height_emu: int,
    shape_id: int = 1,
    name: str = "Figure",
    anchor: str = "inline",
    position: Tuple = (0, 0),
    relative_to: Tuple[str, str] = ("column", "paragraph"),
    wrap: str = "square",
    behind_text: bool = False,
) -> etree._Element:
    """Build a complete `<w:drawing>` for a converted figure.

    Every shape is placed inside a single `wpg:wgp` group so the figure
    behaves as one object in Word — one click selects the whole chart, and
    its parts stay editable underneath.

    `width_emu` / `height_emu` give the on-page size; the group's child
    coordinate space (`a:chOff` / `a:chExt`) is set to the shapes' own union
    bounds so Word rescales the artwork onto the requested box.

    `anchor` picks the envelope:

    * ``"inline"`` (default) — a `wp:inline`, which sits in the text flow
      like a very large character.
    * ``"floating"`` — a `wp:anchor`, which is positioned independently of
      the text and has body text flow around it per `wrap`.

    The floating-only arguments (`position`, `relative_to`, `wrap`,
    `behind_text`) are ignored for an inline drawing.
    """
    if anchor not in ("inline", "floating"):
        raise ValueError("anchor must be 'inline' or 'floating', got %r" % (anchor,))

    shapes = presentationml_to_wordml(xml)
    cx = str(max(1, int(width_emu)))
    cy = str(max(1, int(height_emu)))

    drawing = etree.Element(_qw("drawing"), nsmap={"w": NS_W})

    if anchor == "inline":
        node = etree.SubElement(drawing, "{%s}inline" % NS_WP, nsmap={"wp": NS_WP})
        node.attrib.update({"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
    else:
        if wrap not in _WRAP_ELEMENTS:
            raise ValueError(
                "wrap must be one of %s, got %r" % (sorted(_WRAP_ELEMENTS), wrap)
            )
        node = etree.SubElement(drawing, "{%s}anchor" % NS_WP, nsmap={"wp": NS_WP})
        node.attrib.update(
            {
                "distT": "0",
                "distB": "0",
                "distL": "114300",  # Word's own default 0.125" side gutters
                "distR": "114300",
                "simplePos": "0",
                "relativeHeight": str(shape_id + 251658240),
                "behindDoc": "1" if behind_text else "0",
                "locked": "0",
                "layoutInCell": "1",
                "allowOverlap": "1",
            }
        )
        # `wp:simplePos` is required even though `simplePos="0"` tells Word to
        # ignore it in favour of the positionH/positionV pair below.
        etree.SubElement(node, "{%s}simplePos" % NS_WP).attrib.update(
            {"x": "0", "y": "0"}
        )
        _append_position(node, "h", position[0], relative_to[0])
        _append_position(node, "v", position[1], relative_to[1])

    etree.SubElement(node, "{%s}extent" % NS_WP).attrib.update({"cx": cx, "cy": cy})
    etree.SubElement(node, "{%s}effectExtent" % NS_WP).attrib.update(
        {"l": "0", "t": "0", "r": "0", "b": "0"}
    )

    if anchor == "floating":
        wrap_el = etree.SubElement(node, "{%s}%s" % (NS_WP, _WRAP_ELEMENTS[wrap]))
        if wrap in ("square", "tight", "through"):
            wrap_el.set("wrapText", "bothSides")
        if wrap in ("tight", "through"):
            # These two wrap styles need an explicit outline; a rectangle
            # around the figure matches what Word generates for a grouped
            # drawing, whose true silhouette it cannot compute either.
            polygon = etree.SubElement(wrap_el, "{%s}wrapPolygon" % NS_WP)
            polygon.set("edited", "0")
            etree.SubElement(polygon, "{%s}start" % NS_WP).attrib.update(
                {"x": "0", "y": "0"}
            )
            for x, y in (("0", "21600"), ("21600", "21600"), ("21600", "0"), ("0", "0")):
                etree.SubElement(polygon, "{%s}lineTo" % NS_WP).attrib.update(
                    {"x": x, "y": y}
                )

    doc_pr = etree.SubElement(node, "{%s}docPr" % NS_WP)
    doc_pr.set("id", str(shape_id))
    doc_pr.set("name", name)

    etree.SubElement(node, "{%s}cNvGraphicFramePr" % NS_WP)

    _append_graphic(node, shapes, width_emu, height_emu, name, shape_id)

    return drawing


def register_images_on_part(part, images: Sequence[ImageTuple]) -> dict:
    """Register image blobs on a `python-docx` story part; sentinel -> rId.

    The Word counterpart of :func:`register_images_on_slide`. `python-docx`
    exposes `StoryPart.get_or_add_image`, which allocates the media part name,
    creates the relationship, and reuses an existing part when the same bytes
    are added twice — the same deduplication `python-pptx` does on the slide
    side.
    """
    sentinel_to_rid: dict[str, str] = {}
    for sentinel, data, _format in images:
        rid, _image = part.get_or_add_image(BytesIO(data))
        sentinel_to_rid[sentinel] = rid
    return sentinel_to_rid


def append_drawing_to_paragraph(
    paragraph,
    xml: str,
    images: Sequence[ImageTuple],
    *,
    width_emu: int,
    height_emu: int,
    name: str = "Figure",
    **anchor_kwargs,
) -> etree._Element:
    """Inject a converted figure into a `python-docx` paragraph.

    Mirrors :func:`append_to_sptree_with_images`: register the images on the
    document part, rewrite the sentinel rIds, build the `w:drawing` anchor and
    append it inside a new run. Extra keyword arguments (`anchor`, `position`,
    `relative_to`, `wrap`, `behind_text`) pass through to
    :func:`build_word_drawing`.
    """
    if images:
        mapping = register_images_on_part(paragraph.part, images)
        xml = rewrite_sentinels(xml, mapping)

    # Word keys a floating drawing's stacking order off `relativeHeight` and
    # its identity off `docPr/@id`; both must be unique within the document,
    # so derive an id from how many drawings are already present.
    body = paragraph._p.getroottree().getroot()
    shape_id = len(body.findall(".//{%s}drawing" % NS_W)) + 1

    drawing = build_word_drawing(
        xml,
        width_emu=width_emu,
        height_emu=height_emu,
        name=name,
        shape_id=shape_id,
        **anchor_kwargs,
    )
    run = paragraph.add_run()
    run._r.append(drawing)
    return drawing
