
# mpl-office

**Native vector graphics from Python into Microsoft Office documents.**

> **Status (v0.1 scaffold):** Phases 0–6 are implemented in code under this
> repo. Paths, primitives, styles, text, transforms, linear gradients,
> `<use>`/`<defs>`, raster `<image>` embedding (imshow, colorbars), the
> pptx injection layer, template workflows, and the matplotlib backend
> all work end-to-end, with 55 Rust unit tests and 25 Python pytest
> tests. Deferred for later: `.docx` output and clip-path geometric
> intersection. The sections below have been updated to reflect the
> choices that were actually made while implementing.

## Problem Statement

The R ecosystem, via the `officer` + `rvg` packages, can insert fully editable vector graphics from any R plotting library into PowerPoint and Word documents. The Python ecosystem has no equivalent. Python users are stuck inserting raster PNGs into Office documents, losing editability, resolution independence, and producing bloated files.

This gap exists despite:

- matplotlib having a clean, well-structured SVG backend
- SVG and DrawingML (Office's native vector format) being structurally similar XML formats
- `python-pptx` and `python-docx` already handling all the OOXML packaging machinery
- Proof-of-concept SVG→DrawingML pipelines existing in projects like `typ2pptx` and `svg2pptx`

The missing piece is a robust, fast, general-purpose SVG→DrawingML converter exposed as a matplotlib backend and a composable Python API.

## Prior Art

### R: `officer` + `rvg`

The gold standard. `rvg` implements a custom R graphics device that emits DrawingML directly. Combined with `officer`, users can open a branded template, target a specific slide layout and placeholder, and insert a ggplot as native editable vector shapes. The result is production-grade Office output. Limitation: R-only, and `officer` is notably slow for large documents.

### Python: `svg2pptx` (benouinirachid)

Pure Python, pip-installable, clean API. Converts basic SVG elements to `python-pptx` shapes. **Critical flaw:** linearizes Bezier curves into line segments, producing jagged output on anything with curves. No gradient support. 5 GitHub stars, ~10 commits. MVP quality.

### Python: `typ2pptx` / `svg_to_shapes.py` (from `ppt-master`)

The `svg_to_shapes.py` module (originally from `niccolocorsani/ppt-master`, now preserved in `touying-typ/typ2pptx`) implements the correct algorithm: SVG path parsing → all curve types normalized to cubic Beziers → DrawingML `<a:custGeom>` emission. Supports gradients, opacity inheritance, stroke patterns, group shapes. 254 tests. **This is the reference implementation** for the conversion logic, but it's buried inside a Typst-specific tool and not packaged for standalone use.

### Python: `pyemf`

Pure Python EMF writer from 2006. Proves the concept but targets the legacy EMF format rather than DrawingML. Abandoned.

### Commercial: Spire.Presentation

Proprietary library with `AddFromSVGAsShapes()`. Not open source.

## Design Goals

1. **Rust core for speed and correctness.** The SVG→DrawingML conversion is CPU-bound XML transformation with complex path math (Bezier normalization, arc-to-cubic conversion, coordinate transforms). Rust gives us safe, fast code that can be shared across Python, future R bindings, and CLI usage. This directly addresses `officer`'s known performance problems.

2. **matplotlib backend as the primary interface.** `fig.savefig("chart.pptx")` should just work. The backend wraps matplotlib's own SVG backend—we render to SVG in-memory, then pipe it through the Rust converter. Zero reimplementation of rendering logic.

3. **Template and placeholder support from day one.** The API must support opening an existing `.pptx`/`.docx` template and inserting vector graphics into specific placeholders or positions. Without this, the tool is a toy.

4. **Modular architecture.** The SVG→DrawingML converter is a standalone library. The matplotlib backend, pptx integration, and docx integration are thin layers on top. Other SVG sources (plotly, altair, bokeh, hand-authored SVG) can use the converter directly.

5. **Pragmatic SVG subset.** matplotlib's SVG output uses a small, well-defined subset of SVG. We target that subset first and expand coverage as needed for other sources.

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                   User-facing APIs                   │
│                                                      │
│  fig.savefig("out.pptx")    fig_to_slide(fig, slide)│
│  fig.savefig("out.docx")    fig_to_placeholder(...)  │
│                                                      │
│  ┌─────────────┐  ┌──────────────┐  ┌────────────┐  │
│  │  matplotlib  │  │  pptx        │  │  docx       │  │
│  │  backend     │  │  integration │  │  integration│  │
│  │  (Python)    │  │  (Python)    │  │  (Python)   │  │
│  └──────┬───────┘  └──────┬───────┘  └──────┬──────┘  │
│         │                 │                 │         │
│         └────────┬────────┴────────┬────────┘         │
│                  │                 │                   │
│          SVG string         DrawingML XML             │
│                  │                 │                   │
│         ┌────────▼─────────────────▼────────┐         │
│         │      mpl_office (Python package)   │         │
│         │      PyO3 bindings                 │         │
│         └────────────────┬───────────────────┘         │
└──────────────────────────┼───────────────────────────┘
                           │
              ┌────────────▼────────────────┐
              │   mpl-office-core (Rust)     │
              │                              │
              │  ┌────────────────────────┐  │
              │  │   SVG Parser            │  │
              │  │   (quick-xml + custom)  │  │
              │  └───────────┬─────────────┘  │
              │              │                │
              │  ┌───────────▼─────────────┐  │
              │  │   IR (intermediate repr) │  │
              │  │   Shapes, paths, text,   │  │
              │  │   groups, styles         │  │
              │  └───────────┬─────────────┘  │
              │              │                │
              │  ┌───────────▼─────────────┐  │
              │  │   Path Normalizer        │  │
              │  │   All curves → cubic     │  │
              │  │   Beziers, arc→cubic,    │  │
              │  │   relative→absolute      │  │
              │  └───────────┬─────────────┘  │
              │              │                │
              │  ┌───────────▼─────────────┐  │
              │  │   DrawingML Emitter      │  │
              │  │   IR → OOXML fragments   │  │
              │  └────────────────────────┘  │
              └──────────────────────────────┘
```

### Concrete workspace layout

```
mpl-office/
├── Cargo.toml                      # Rust workspace root
├── pyproject.toml                  # Python project + maturin config
├── crates/
│   ├── mpl-office-core/             # Pure-Rust converter
│   │   └── src/
│   │       ├── lib.rs               # Public API: convert_svg_to_drawingml, ConvertOptions
│   │       ├── error.rs             # Error/Result types (thiserror)
│   │       ├── coord.rs             # EMU math, DPI helpers
│   │       ├── color.rs             # #RGB/#RRGGBB/rgb()/rgba()/named parsers
│   │       ├── transform.rs         # 2D affine + SVG transform-list parser
│   │       ├── path.rs              # Tokenize → absolutize → normalize → bbox
│   │       ├── style.rs             # Style struct + cascading rules
│   │       ├── ir.rs                # IrDocument, Node, NodeKind, gradients, defs
│   │       ├── parse.rs             # quick-xml streaming parser → IR
│   │       └── emit.rs              # IR → DrawingML XML
│   └── mpl-office-py/               # PyO3 crate → mpl_office._native
│       └── src/lib.rs               # #[pymodule] + ConvertOptions class
└── python/mpl_office/               # Python package source
    ├── __init__.py                  # Re-exports convert_svg_to_drawingml, ConvertOptions
    ├── _inject.py                   # lxml helpers for <p:spTree> injection
    ├── pptx.py                      # svg_to_slide, fig_to_slide, fig_to_placeholder
    └── backend.py                   # matplotlib backend (FigureCanvasOffice)
```

The native extension is exposed as `mpl_office._native` and re-exported
through `mpl_office.__init__` so users type `from mpl_office import ...`
without touching the underscore module.

### Rust Core (`mpl-office-core`)

**Input:** SVG string (or stream).

**Output:** DrawingML XML fragments — specifically, one or more `<p:sp>` (shape) or `<p:grpSp>` (group shape) elements ready to be inserted into an OOXML document's shape tree.

**Not responsible for:** OOXML packaging (ZIP structure, relationships, content types). That stays in Python via `python-pptx` / `python-docx`, which already handle it well.

#### Modules

**`svg_parse`** — Streaming SVG parser built on `quick-xml`. Extracts the elements and attributes we care about, ignores the rest. Handles `<defs>`, `<use>` references, `<clipPath>`, CSS `style` attributes, and `transform` attributes.

**`ir`** — Intermediate representation. A flat-ish tree of typed nodes:

```
enum Node {
    Group { children: Vec<Node>, transform: Transform, opacity: f64 },
    Path { commands: Vec<PathCmd>, style: Style },
    Rect { x, y, width, height, rx, ry, style: Style },
    Circle { cx, cy, r, style: Style },
    Ellipse { cx, cy, rx, ry, style: Style },
    Line { x1, y1, x2, y2, style: Style },
    Polyline { points: Vec<(f64,f64)>, style: Style },
    Polygon { points: Vec<(f64,f64)>, style: Style },
    Text { x, y, content: String, style: TextStyle },
    Image { href: String, x, y, width, height },
}
```

**`path_normalize`** — The mathematical core. Converts all SVG path commands to absolute coordinates, normalizes S/Q/T/A commands to cubic Beziers (C commands). Arc-to-cubic conversion uses the standard endpoint-to-center parameterization followed by cubic approximation (the same algorithm used in `typ2pptx` and every serious SVG implementation). This module must be exhaustively tested.

**`drawingml_emit`** — Converts IR nodes to DrawingML XML strings. Key mappings:

| SVG | DrawingML |
|-----|-----------|
| `<path>` | `<a:custGeom>` with `<a:path>` containing `<a:moveTo>`, `<a:lnTo>`, `<a:cubicBezTo>`, `<a:close>` |
| `<rect>` | `<a:prstGeom prst="rect">` or `<a:custGeom>` if rounded corners |
| `<circle>`, `<ellipse>` | `<a:prstGeom prst="ellipse">` |
| `<line>` | `<a:custGeom>` with moveTo + lnTo |
| `<text>` | `<p:sp>` with `<a:txBody>` containing `<a:r>` runs |
| `<g>` | `<p:grpSp>` |
| `transform` | `<a:xfrm>` with `<a:off>` and `<a:ext>`, rotation via `rot` attribute |
| `fill` | `<a:solidFill>` or `<a:gradFill>` |
| `stroke` | `<a:ln>` with `<a:solidFill>`, dash patterns via `<a:prstDash>` |
| `opacity` | `<a:solidFill><a:srgbClr val="..."><a:alpha val="..."/></a:srgbClr></a:solidFill>` |
| `clip-path` | Flattened (clip applied to path geometry) for v1; proper `<a:clip>` later |

**`coord`** — Coordinate system conversion. SVG uses pixels with Y-down. DrawingML uses EMUs (English Metric Units: 914400 EMU = 1 inch) with Y-down. The converter needs a configurable DPI assumption (matplotlib SVG defaults to 72 DPI) and a target bounding box (the placeholder or slide region) to compute the scaling transform.

### Implementation decisions (as built)

- **Target-bbox scaling is folded into the root affine.** When
  `target_width_emu` / `target_height_emu` are set, `EmitContext` builds a
  single affine that maps the SVG viewBox onto the EMU bbox, and every
  shape's local transform is composed on top of it. That way each shape
  emits final EMU coordinates directly — there is no second post-scaling
  pass. Without a target, we just apply a DPI correction (96/source_dpi).
- **Path math is hand-rolled, not `kurbo`.** The arc-to-cubic algorithm
  (SVG F.6 endpoint → center parameterization, followed by the standard
  4/3·tan(θ/4) approximation) lives in `path.rs::arc_to_cubics`. It's ~60
  lines and exhaustively tested; skipping `kurbo` saves a dependency and
  keeps the crate lean. See "Open Questions → kurbo vs. hand-rolled"
  below for the tradeoff.
- **`<text>` stays a single text box per `<text>` element.** Per-`<tspan>`
  runs become `<a:r>` runs inside one `<a:p>`, which matches the
  reference implementation and is exactly what matplotlib's SVG output
  needs (one `<text>` per tick label, one per axis title).
- **`<use>` is inlined at emit time.** When the emitter hits a
  `NodeKind::Use`, it looks up the referenced node in the defs table and
  recursively emits it with the use's `x`/`y` attributes composed into
  the current transform. This was a bug fix — matplotlib uses `<use>`
  for every tick mark, and the initial implementation ignored those
  `x`/`y` attributes, stacking all ticks at the origin.
- **Groups flatten when they only contain one shape.** The `<p:grpSp>`
  wrapper only appears when a group has two or more children, avoiding
  an unnecessary level of nesting for matplotlib's many one-shape
  subgroups. Group bounds are computed by scanning the emitted inner
  XML for `<a:off>`/`<a:ext>` pairs — a pragmatic pattern that avoids a
  second tree walk.
- **Rects and ellipses use `prstGeom` when possible.** If the effective
  transform is translate + axis-aligned scale (no rotation or skew) and
  the rect has no rounded corners, we emit `<a:prstGeom prst="rect">`
  instead of a custom path. It produces crisper, smaller, and more
  recognisable shapes in PowerPoint. Any rotation, skew, or rounded
  corner falls back to `<a:custGeom>` with a cubic-Bézier approximation
  (κ = 0.5522847498 for circle arcs).

### Python Package (`mpl-office`)

Built with PyO3 0.22 + maturin. ABI3 floor is **Python 3.9**
(`pyo3 = { features = ["abi3-py39"] }`), so one wheel per platform
covers every future Python ≥ 3.9. `requires-python = ">=3.9"` in
`pyproject.toml` matches.

Exposes:

#### Low-level: `convert_svg_to_drawingml(svg: str, options: ConvertOptions) -> str`

Returns DrawingML XML fragment(s). This is the universal entry point that any SVG source can use.

```python
from mpl_office import convert_svg_to_drawingml, ConvertOptions

opts = ConvertOptions(
    target_width_emu=6858000,   # 7.5 inches
    target_height_emu=4572000,  # ~5 inches
    source_dpi=72,
)
drawingml_xml = convert_svg_to_drawingml(svg_string, opts)
```

#### Mid-level: pptx and docx integration

```python
from mpl_office.pptx import fig_to_slide, fig_to_placeholder
from mpl_office.docx import fig_to_document
from pptx import Presentation
from docx import Document

# === PowerPoint ===

# Open a corporate template
prs = Presentation("quarterly_template.pptx")
slide = prs.slides.add_slide(prs.slide_layouts[1])

# Insert into a specific placeholder
fig_to_placeholder(fig, slide.placeholders[1])

# Or free-positioned
fig_to_slide(fig, slide,
             left=Inches(0.5), top=Inches(1.5),
             width=Inches(9), height=Inches(5))

prs.save("output.pptx")

# === Word ===

doc = Document("report_template.docx")
fig_to_document(fig, doc,
                width=Inches(6),
                anchor="inline")  # or "floating"
doc.save("output.docx")
```

These functions:
1. Render the figure to SVG via matplotlib's SVG backend (in-memory `StringIO`)
2. Call the Rust core to convert SVG→DrawingML
3. Inject the resulting XML into the `python-pptx` or `python-docx` object model

The injection step requires reaching into `python-pptx`/`python-docx` internals to append shape XML to the slide's `spTree` or the document's body. Both libraries use `lxml` internally, so this is `etree.SubElement()` operations. We will need to also register relationships (for images, if any) and handle the `<p:grpSp>` wrapper that positions and scales the entire figure within the target region.

**Implementation notes:**

- **Matplotlib renders with `svg.fonttype='none'` forced via rc_context.**
  Without this override, matplotlib's default SVG backend rasterizes text
  into glyph outlines, which defeats the whole point of native DrawingML
  output. `_render_fig_svg` in `python/mpl_office/pptx.py` wraps the
  `savefig` call in `mpl.rc_context({"svg.fonttype": "none"})` so users
  don't have to think about it.
- **Namespace-safe injection.** The Rust emitter writes bare `<p:sp>` and
  `<a:...>` tags without namespace declarations. `_inject.parse_drawingml_fragment`
  wraps the fragment in a throwaway `<root>` element that declares all
  OOXML namespaces, parses it, and returns the top-level children — ready
  to be appended via `spTree.append(el)`.
- **`fig_to_placeholder` takes the slide explicitly.** The first draft used
  `placeholder.part.slide` to look up the owner, but that attribute isn't
  stable across `python-pptx` versions. The signature is now
  `fig_to_placeholder(fig, slide, placeholder, *, remove_placeholder=True)`
  and the placeholder element is removed from the slide tree by default so
  the figure cleanly replaces it.
- **Image embedding uses a sentinel-rId protocol.** The Rust emitter has
  no knowledge of `python-pptx`'s relationship tables, so it writes
  `r:embed="__mpl_office_img_N__"` on each `<p:pic>` and returns the
  image blobs in a parallel `Vec<EmittedImage>`. The Python layer
  (`_inject.append_to_sptree_with_images`) then iterates the list,
  calls `slide.part.get_or_add_image_part(BytesIO(bytes))` for each,
  captures the real `rId` that python-pptx allocates, and does a literal
  string replacement on the XML (`f'r:embed="{sentinel}"'` →
  `f'r:embed="{rid}"'`) before parsing. The sentinels are opaque
  strings that cannot collide with real rIds or with any other XML
  content, so `str.replace` is both correct and fast. The legacy
  `convert_svg_to_drawingml` function is kept for back-compat — it
  emits the same XML (with unresolved sentinels) but discards the
  image list, which is the behaviour the old API documented.

#### High-level: matplotlib backend

```python
import matplotlib
matplotlib.use('module://mpl_office.backend')

import matplotlib.pyplot as plt

fig, ax = plt.subplots()
ax.plot([1, 2, 3], [1, 4, 9])

# Simple: one figure → one slide
fig.savefig("chart.pptx")

# With template
fig.savefig("chart.pptx",
            template="corporate.pptx",
            layout_index=1,
            placeholder_index=1)

# Word
fig.savefig("chart.docx",
            template="report.docx",
            width_inches=6)
```

The backend class:

```python
from matplotlib.backends.backend_svg import FigureCanvasSVG
from io import StringIO

class FigureCanvasOffice(FigureCanvasSVG):

    def print_pptx(self, filename, *,
                   template=None, layout_index=0,
                   placeholder_index=None, **kwargs):
        buf = StringIO()
        self.print_svg(buf, **kwargs)
        svg = buf.getvalue()
        _save_pptx(svg, filename, template,
                   layout_index, placeholder_index)

    def print_docx(self, filename, *,
                   template=None, width_inches=6.5, **kwargs):
        buf = StringIO()
        self.print_svg(buf, **kwargs)
        svg = buf.getvalue()
        _save_docx(svg, filename, template, width_inches)

FigureCanvas = FigureCanvasOffice
FigureManager = FigureManagerBase
```

## SVG Subset: matplotlib Coverage

matplotlib's SVG backend (`backend_svg.py`) emits a specific, well-behaved subset of SVG. This is our primary target. Audit of matplotlib 3.9's SVG output:

### Elements Used

| Element | Usage | Priority |
|---------|-------|----------|
| `<svg>` | Root with viewBox | P0 |
| `<defs>` | Clip paths, reusable elements | P0 |
| `<g>` | Grouping with transforms, clip-path refs | P0 |
| `<path>` | All line/area/curve geometry (plots, axes, ticks, legend boxes) | P0 |
| `<text>` | Axis labels, titles, tick labels, annotations | P0 |
| `<clipPath>` | Axes region clipping | P0 |
| `<use>` | Reused markers/symbols | P1 |
| `<image>` | Embedded raster (imshow, etc.) | P1 |
| `<rect>` | Occasional (legend backgrounds) | P1 |
| `<style>` | CSS class definitions | P1 |

### Attributes Used

| Attribute | Notes |
|-----------|-------|
| `d` | Path data — M, L, C, Z primarily; Q, A occasionally (arc markers) |
| `transform` | `translate(x,y)`, `scale(x,y)`, `rotate(deg)`, `matrix(...)` |
| `style` | Inline CSS: `fill`, `stroke`, `stroke-width`, `stroke-linecap`, `stroke-linejoin`, `stroke-dasharray`, `opacity`, `fill-opacity`, `stroke-opacity`, `font-size`, `font-family`, `font-style`, `font-weight` |
| `clip-path` | `url(#...)` references |
| `class` | CSS class references |

### What matplotlib does NOT use

- Filters (`<filter>`, `feGaussianBlur`, etc.)
- Masks (`<mask>`)
- Complex CSS (no external stylesheets, no `@import`)
- Animations
- `<foreignObject>`
- Radial gradients (linear gradients are rare, used in some colorbars)
- `<pattern>`

This is extremely favorable. The subset is small and well-defined.

## Implementation Plan

### Phase 0: Scaffolding (Week 1)

**Goal:** Project structure, build pipeline, "hello world" round-trip.

- [ ] Create Rust workspace: `mpl-office-core` library crate
- [ ] Create Python package with maturin + PyO3 bindings
- [ ] Set up CI (GitHub Actions): Rust tests, Python tests, cross-platform wheels
- [ ] Implement trivial round-trip: hardcoded SVG `<rect>` → DrawingML `<a:prstGeom prst="rect">` → injected into a `python-pptx` slide → valid .pptx file opens in PowerPoint/LibreOffice
- [ ] Verify: open the .pptx, click the shape, confirm it's editable

### Phase 1: Path Pipeline (Weeks 2–3)

**Goal:** The mathematical core. Correct conversion of arbitrary SVG paths to DrawingML custom geometry.

- [ ] SVG path `d` attribute tokenizer/parser
- [ ] Relative→absolute coordinate conversion
- [ ] Shorthand resolution (S→C, T→Q)
- [ ] Quadratic→cubic Bezier promotion (Q→C)
- [ ] Arc→cubic Bezier conversion (A→C sequence)
- [ ] DrawingML `<a:custGeom>` emitter: `<a:moveTo>`, `<a:lnTo>`, `<a:cubicBezTo>`, `<a:close>`
- [ ] EMU coordinate conversion with configurable DPI and target bounds
- [ ] Property tests: round-trip path parsing, known-good SVG→DML pairs
- [ ] Visual regression tests: render DrawingML output back to image, compare against SVG rendering

**Key reference:** `typ2pptx/scripts/svg_to_shapes.py` functions `parse_svg_path()`, `svg_path_to_absolute()`, `normalize_path_commands()`, `path_commands_to_drawingml()`. Rewrite in Rust, not port — different language idioms, different error handling, stricter types.

### Phase 2: Core SVG Elements + Styles (Weeks 3–4)

**Goal:** Handle the full set of SVG elements matplotlib emits, with correct styling.

- [ ] SVG parser (using `quick-xml`): streaming parse of elements, attributes, inline `style`, `class` + `<style>` CSS resolution
- [ ] Build IR tree from parsed SVG
- [ ] `<rect>`, `<circle>`, `<ellipse>`, `<line>`, `<polyline>`, `<polygon>` → IR → DrawingML
- [ ] `<g>` with transforms → `<p:grpSp>` with `<a:xfrm>`
- [ ] Transform parsing and composition: `translate`, `scale`, `rotate`, `matrix`
- [ ] Style resolution: inline `style` attribute → `Style` struct (fill, stroke, stroke-width, opacity, dash arrays, line caps, line joins)
- [ ] `<a:solidFill>` with color parsing (hex, `rgb()`, `rgba()`, named colors)
- [ ] `<a:ln>` with stroke properties
- [ ] Opacity → alpha channel on fills/strokes

### Phase 3: Text (Week 5)

**Goal:** Axis labels, titles, tick labels, annotations rendered as editable Office text.

- [ ] `<text>` and `<tspan>` parsing with `x`, `y`, `dx`, `dy` positioning
- [ ] Font property mapping: `font-family` → Office font name, `font-size` (SVG px → Office pt), `font-weight` → bold, `font-style` → italic
- [ ] Text color
- [ ] Rotation (common for y-axis labels: `transform="rotate(-90, ...)"`)
- [ ] DrawingML `<p:sp>` with `<a:txBody>`, `<a:p>`, `<a:r>` emission
- [ ] Text anchor/alignment mapping

Text is the highest-risk element. Font metrics differ between SVG rendering and Office rendering, so exact positioning will require care. matplotlib's SVG backend positions each text element with explicit `x,y` coordinates, which helps — we can create individual text boxes rather than trying to flow text.

### Phase 4: Clipping + Defs + Use (Week 5–6)

**Goal:** Proper axes clipping (so plot data doesn't overflow axes bounds) and reusable elements.

- [ ] `<defs>` parsing and symbol table
- [ ] `<clipPath>` → for v1, apply clip as a geometric intersection on contained paths (conservative but correct). DrawingML has no direct clip-path equivalent for arbitrary shapes on groups; the pragmatic approach is to clip at the path level.
- [ ] `<use>` → inline expansion (resolve `href`, apply transform, emit as if the element were inline)
- [ ] Linear gradients (`<linearGradient>`) → `<a:gradFill>` (used by some colorbars)

### Phase 5: Python Integration (Week 6–7)

**Goal:** `python-pptx` and `python-docx` injection, matplotlib backend, template support.

- [ ] PyO3 bindings: expose `convert_svg_to_drawingml()` accepting SVG string, returning DrawingML XML string
- [ ] `ConvertOptions` struct exposed to Python: target dimensions, DPI, optional bounding box
- [ ] pptx integration module:
  - [ ] `fig_to_slide(fig, slide, left, top, width, height)` — render fig → SVG → DrawingML → inject into slide's `spTree`
  - [ ] `fig_to_placeholder(fig, placeholder)` — same but auto-sized to placeholder bounds
  - [ ] Handle `<p:grpSp>` wrapper with `<a:xfrm>` for positioning/scaling within target region
  - [ ] Relationship registration for any embedded images
- [ ] docx integration module:
  - [ ] `fig_to_document(fig, doc, width, anchor)` — inject as inline or floating drawing
  - [ ] DrawingML wrapped in `<w:drawing><wp:inline>` or `<wp:anchor>` with appropriate extents
- [ ] matplotlib backend module:
  - [ ] `FigureCanvasOffice` class inheriting `FigureCanvasSVG`
  - [ ] `print_pptx()` and `print_docx()` methods
  - [ ] `backend` module registration for `matplotlib.use('module://mpl_office.backend')`
- [ ] Template support:
  - [ ] `template=` parameter on `savefig` and integration functions
  - [ ] Layout selection by index or name
  - [ ] Placeholder targeting by index or type

### Phase 6: Testing + Hardening (Week 7–8)

**Goal:** Confidence that output is correct across Office applications.

- [ ] Rust unit tests: path parsing, normalization, coordinate conversion, DrawingML emission (aim for >90% coverage on core modules)
- [ ] Python integration tests: end-to-end figure → .pptx/.docx → validate OOXML
- [ ] Visual regression suite:
  - Generate a gallery of matplotlib figures (line plots, scatter, bar, pie, subplots, legends, colorbars, annotations, log scales, polar plots, 3D projections)
  - Convert each to .pptx
  - Open in LibreOffice (headless), export to PNG
  - Compare against matplotlib's direct PNG output
  - Flag regressions above pixel-difference threshold
- [ ] Validate output .pptx/.docx files with Microsoft's OOXML SDK validator
- [ ] Test with: PowerPoint (Windows), PowerPoint (Mac), PowerPoint Online, LibreOffice Impress, Google Slides
- [ ] Template tests: test with several real-world corporate templates

### Phase 7: Polish + Release (Week 8–9)

- [ ] Documentation: README, API reference, cookbook with common patterns
- [ ] PyPI packaging: maturin-built wheels for Linux (manylinux), macOS (x86+arm), Windows
- [ ] Examples gallery: Jupyter notebooks showing common workflows
- [ ] Performance benchmarks vs. R `officer` + `rvg` on equivalent plots
- [ ] `--help` CLI tool: `mpl-office convert input.svg output.pptx` for non-Python users

## Future Work (Post-v1)

- **plotly/altair/bokeh SVG ingestion.** Their browser-rendered SVG is messier (inline styles, nested transforms, possibly rasterized text). Add an SVG normalization/cleanup pass or per-source adapters.
- **R bindings** via extendr. This would give R users a faster alternative to `rvg`'s DrawingML emission while keeping the same `officer` workflow.
- **Editable charts.** Instead of converting plot geometry to flat shapes, emit DrawingML `<c:chart>` objects with data tables. This would make the charts truly editable (change data in PowerPoint). Much harder — requires understanding the semantic structure of a matplotlib figure, not just its visual output. Probably a separate project.
- **xlsx integration.** Insert vector charts into Excel spreadsheets via `openpyxl`.
- **Bidirectional: DrawingML→SVG.** For reading Office vector graphics back into Python. The Rust core's IR is symmetric — add a DrawingML parser and SVG emitter.
- **WASM build.** The Rust core can compile to WASM for browser-side conversion (e.g., in JupyterLite or web apps).

## Dependencies

### Rust

| Crate | Purpose | Notes |
|-------|---------|-------|
| `quick-xml` 0.36 | SVG parsing | Streaming reader, hand-rolled attribute decoder |
| `thiserror` 1 | Error type derive | |
| `base64` 0.22 | `data:` URI decoding | Used by the parser to extract inline PNGs/JPEGs from `<image>` hrefs |
| `pyo3` 0.22 | Python bindings | `abi3-py39` + `extension-module` features |

**Dropped from the original spec:**

- **`cssparser`** — we only needed to parse inline `style="..."` declarations
  and a tiny subset of CSS class rules, which `style.rs` handles in ~30
  lines. A full CSS parser would be overkill.
- **`kurbo`** — see the Open Questions section. Hand-rolled arc-to-cubic
  is ~60 lines and avoids a chunky dependency for a crate that aims to
  stay lean.

### Python

| Package | Purpose | Notes |
|---------|---------|-------|
| `maturin` ≥ 1.5 | Build system | `python-source = "python"` layout |
| `lxml` | OOXML injection | Required for pptx integration |
| `python-pptx` | OOXML PowerPoint manipulation | Extra: `pptx` |
| `python-docx` | OOXML Word manipulation | Extra: `docx`, deferred |
| `matplotlib` ≥ 3.5 | SVG rendering | Extra: `matplotlib` |

`matplotlib`, `python-pptx`, and `python-docx` are peer/optional dependencies — the core converter works without them.

### Build and dev tooling

- **`uv`** is the project's package manager (`uv sync --all-extras`,
  `uv run pytest`). Native extension builds run through
  `.venv/Scripts/maturin.exe develop --release`.
- **CI** runs on GitHub Actions. `.github/workflows/ci.yml` runs
  `cargo fmt --check`, `cargo clippy -D warnings`, `cargo test`, and a
  pytest matrix over `{ubuntu, macos-14, windows} × {3.9, 3.12}`.
  `.github/workflows/release.yml` builds a release wheel matrix
  (manylinux / musllinux x86_64+aarch64, macOS universal2, windows x64,
  sdist) and publishes to PyPI via OIDC trusted publishing on `v*` tags.
  See `RELEASING.md` for the internal release procedure.

## Open Questions and Decisions

1. **Clipping strategy. (Deferred.)** DrawingML doesn't have a direct
   equivalent of SVG's `<clipPath>` applied to a group. The v1 emitter
   **parses `<clipPath>` into the defs table but does not apply it** — data
   that falls outside a manually-tightened `xlim`/`ylim` will spill past the
   axes frame. Acceptable because matplotlib's default "data clipping ≈
   axes bbox" matches the emitter's implicit output region, and visually
   correct clipping would require a geometric-intersection pass that's
   expensive for scatter/line plots with thousands of segments. Revisit
   when a user reports it.

2. **Text fidelity.** Font metrics will differ between matplotlib's
   rendering and Office's rendering. For tick labels and axis labels
   (short, positioned individually), this is fine — each is its own text
   box with explicit `x,y`. For multi-line annotations or wrapped text,
   positioning may drift. We use the same rough-estimate text-width
   function as the `typ2pptx` reference (charwise table + 1.15× fudge
   factor) and force `svg.fonttype='none'` in the matplotlib render
   helper so glyphs stay editable. Acceptable for v1.

3. **`python-pptx` / `python-docx` internals.** Injecting raw XML into
   these libraries' object models means depending on their internal
   structure (`slide.shapes._spTree`, `placeholder._element`). These
   aren't public APIs. The `_inject.py` module centralises every such
   access so a future version bump only needs to touch one file. `docx`
   injection is deferred (see #5).

4. **3D plots.** matplotlib's 3D projection produces SVG paths (it's all
   2D by the time it hits the backend). These will convert fine
   geometrically but may produce very large shape counts. May need a
   complexity threshold that falls back to raster for extremely dense
   plots.

5. **`.docx` output. (Deferred.)** The initial `backend.print_docx`
   raises `NotImplementedError`. Word's drawing model wraps shapes in
   `<w:drawing><wp:inline>…<wps:wsp>` rather than `<p:sp>` directly, so
   the pptx emitter's output needs a second rewrite pass. An early
   prototype lived in `_inject.py` but was removed to keep v1 focused.
   Revisit after a user actually asks for it.

6. **`<image>` emission. (Resolved.)** Inline `data:image/...;base64,...`
   URIs are decoded in the parser (see `parse::decode_data_uri`), stored
   as `ImageData { bytes, format }` on the IR node, and the emitter
   writes a `<p:pic>` shape with a sentinel `r:embed="__mpl_office_img_N__"`.
   `EmitContext::images` collects the image blobs in order; the caller
   uses `convert_svg_to_drawingml_with_images` (PyO3 exposes the same
   name) and receives `(xml, Vec<EmittedImage>)`. The Python
   `_inject.append_to_sptree_with_images` registers each blob as an
   OOXML image part via `slide.part.get_or_add_image_part(BytesIO(bytes))`,
   gets a real `rId` back, rewrites the sentinels in the XML by literal
   string replacement, and then appends the elements to `<p:spTree>`.
   Duplicate bytes are deduplicated automatically by python-pptx's
   image-part cache. External `<image>` file references (no `data:`
   prefix) are still dropped because the core crate has no filesystem
   access; that's the one remaining gap and not expected to bite
   matplotlib users since matplotlib's SVG backend always inlines
   rasters as data URIs.

7. **`kurbo` vs. hand-rolled path math. (Resolved: hand-rolled.)** The
   original recommendation was to use `kurbo`. In practice the total
   footprint of arc-to-cubic + quadratic-to-cubic + affine + bbox is
   ~80 lines of straightforward numerical code, fully covered by unit
   tests, and avoids pulling a geometry library (which also brings in
   `arrayvec`, `smallvec`, etc.). `path.rs::arc_to_cubics` is the only
   piece of numerically-interesting code and it's a direct transcription
   of the SVG 1.1 F.6 algorithm. Revisit if we ever need Bézier
   intersection, subdivision, or flattening — those are where `kurbo`
   really earns its keep.

8. **Group wrapping heuristic. (Resolved.)** We flatten groups with a
   single child and wrap groups with two or more children in `<p:grpSp>`.
   Group bounds are computed by scanning the emitted inner XML for
   `<a:off>`/`<a:ext>` pairs. This sidesteps a second tree walk and
   matches the reference implementation, at the cost of a tight coupling
   between emitter output shape and the bounds-extraction regex. The
   regex lives next to the emitter so a change in one forces a change
   in the other.