# mpl-office

**Native vector graphics from Python into Microsoft Office documents.**

```python
import matplotlib.pyplot as plt

fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [1, 4, 9, 16])
ax.set_title("Native DrawingML, not a PNG")

fig.savefig("chart.pptx", backend="module://mpl_office.backend")
```

Open `chart.pptx` in PowerPoint. Click any line, any tick label, any
title — it's all editable: real text, real shapes, real colors. No
rasterized images, no embedded bitmaps, no "convert to shape" hack. Just
DrawingML, like PowerPoint drew it itself.

## What this is

Python users who want editable vector graphics in Office documents have
historically had two bad options: insert raster PNGs (and live with jagged
zoom, bloated files, and opaque content), or drive Office via COM on
Windows. The R ecosystem solved this years ago with `officer` + `rvg`,
which emit DrawingML directly from R's graphics device. `mpl-office`
brings the same workflow to Python.

- **A Rust core** (`mpl-office-core`) converts SVG into DrawingML XML
  fragments. Path normalization (arc → cubic Bézier, relative → absolute,
  smooth curves expanded), style cascading, transform composition, and
  linear-gradient support are all handled here.
- **A Python layer** uses `matplotlib`'s SVG backend to render figures
  in memory, hands the string to the Rust core, and injects the resulting
  DrawingML into a `python-pptx` slide's shape tree.
- **A matplotlib backend** makes `fig.savefig("out.pptx")` work end-to-end,
  with optional template and placeholder targeting.

The Rust core is a standalone library — any SVG source (plotly, altair,
hand-authored SVG) can use it directly.

## Installing

### End users

```bash
pip install mpl-office
# or, for PowerPoint integration:
pip install "mpl-office[pptx,matplotlib]"
```

Wheels are published for Linux (x86_64, aarch64, manylinux + musllinux),
macOS (universal2), and Windows (x64). A Rust toolchain is only needed if
you want to build from source.

### Building from source

```bash
git clone <repo>
cd mpl-office
uv sync --all-extras
.venv/Scripts/maturin.exe develop --release   # Windows
# or
.venv/bin/maturin develop --release            # macOS/Linux
uv run pytest
```

Requires Rust ≥ 1.70 and Python ≥ 3.9.

## Usage

### The matplotlib backend

The simplest path: tell matplotlib to use the `mpl_office` backend for a
single save call.

```python
import matplotlib.pyplot as plt

fig, ax = plt.subplots(figsize=(8, 5))
ax.plot([1, 2, 3, 4, 5], [1, 4, 9, 16, 25], label="squares")
ax.plot([1, 2, 3, 4, 5], [1, 8, 27, 64, 125], label="cubes")
ax.legend()
ax.set_title("Polynomial growth")

fig.savefig("chart.pptx", backend="module://mpl_office.backend")
```

To make the backend the default for an entire script, register it at the
top:

```python
import matplotlib
matplotlib.use("module://mpl_office.backend")
```

### Inserting into an existing slide deck

When you already have a `python-pptx` `Presentation` open, use
`fig_to_slide` to place a figure at specific coordinates:

```python
from pptx import Presentation
from mpl_office.pptx import Inches, fig_to_slide
import matplotlib.pyplot as plt

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

fig, ax = plt.subplots(figsize=(10, 5))
ax.bar(["Q1", "Q2", "Q3", "Q4"], [4.2, 5.1, 6.8, 7.5],
       color="#2E86AB")
ax.set_title("Revenue")

fig_to_slide(
    fig, slide,
    left=Inches(1.5), top=Inches(1),
    width=Inches(10), height=Inches(5.5),
)
prs.save("report.pptx")
```

Left, top, width, and height are in EMU. `Inches()` is a small helper that
converts — you can also pass raw integers or use `python-pptx`'s own
`Inches` (they're interchangeable; both return EMU).

### Working with templates

Open a branded corporate template and append figure slides against it.
The template's title slide, master layout, theme colors, and existing
slides are all preserved — `mpl-office` just adds new content.

```python
from pptx import Presentation
from mpl_office.pptx import Inches, fig_to_slide

prs = Presentation("quarterly_template.pptx")

# Add a blank slide using one of the template's layouts
blank_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_layout)

fig_to_slide(
    fig, slide,
    left=Inches(1.5), top=Inches(1.25),
    width=Inches(10), height=Inches(5),
)
prs.save("quarterly_report.pptx")
```

You can also target a placeholder on a template slide — `fig_to_placeholder`
removes the placeholder and drops the figure into its exact bounds:

```python
from mpl_office.pptx import fig_to_placeholder

prs = Presentation("template.pptx")
slide = prs.slides[0]

content_ph = next(
    ph for ph in slide.placeholders
    if ph.placeholder_format.idx != 0  # skip title
)
fig_to_placeholder(fig, slide, content_ph)
prs.save("out.pptx")
```

The matplotlib backend understands templates too:

```python
fig.savefig(
    "out.pptx",
    backend="module://mpl_office.backend",
    template="quarterly_template.pptx",
    layout_index=6,
)
```

### Raster images (`imshow`, colorbars, heatmaps)

Figures with `imshow` and rasterized colorbar cells work out of the box.
matplotlib's SVG backend emits those regions as `<image>` elements with
inline `data:image/png;base64,...` URIs. `mpl-office` decodes them,
registers the bytes as native OOXML picture parts on the destination
slide, and wires up the relationship ids automatically — so a heatmap
arrives in PowerPoint as a real embedded PNG surrounded by vector axes,
tick labels, and titles.

```python
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from mpl_office.pptx import Inches, fig_to_slide

rng = np.random.default_rng(42)
fig, ax = plt.subplots(figsize=(8, 6))
im = ax.imshow(rng.standard_normal((40, 40)), cmap="viridis")
ax.set_title("Heatmap (raster data, vector chrome)")
fig.colorbar(im, ax=ax)

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
fig_to_slide(fig, slide, left=Inches(1), top=Inches(1),
             width=Inches(8), height=Inches(6))
prs.save("heatmap.pptx")
```

Open `heatmap.pptx` — the heatmap cells are a single editable picture
shape, but the title, axis labels, tick marks, and colorbar frame are
all individually selectable vector elements.

Duplicated images (e.g. two subplots with the same bitmap) are
deduplicated automatically by `python-pptx`'s image cache, so the output
file only stores each unique PNG once.

### Using the converter directly

The low-level API takes a raw SVG string and returns a DrawingML XML
fragment. This is the universal entry point for any SVG source — plotly,
altair, hand-authored SVG, the output of a server-rendered chart service.

```python
from mpl_office import ConvertOptions, convert_svg_to_drawingml

svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
  <circle cx="50" cy="50" r="40" fill="#3498db" stroke="#2c3e50" stroke-width="3"/>
  <text x="50" y="55" text-anchor="middle" font-family="Segoe UI"
        font-size="14" fill="white">Hello</text>
</svg>"""

opts = ConvertOptions(
    source_dpi=96.0,
    target_width_emu=3_657_600,   # 4 inches
    target_height_emu=3_657_600,
    offset_x_emu=914_400,          # 1 inch from left
    offset_y_emu=914_400,          # 1 inch from top
)
drawingml = convert_svg_to_drawingml(svg, opts)
# drawingml is a <p:sp>...</p:sp> fragment ready to inject
```

Combine with `mpl_office._inject.append_to_sptree(slide, drawingml)` to
add the result to a `python-pptx` slide.

## What's supported

`mpl-office` targets the SVG subset that matplotlib's SVG backend emits,
and covers enough of the broader SVG spec to be useful for other sources.

**SVG elements:**

- `<rect>`, `<circle>`, `<ellipse>`, `<line>` — mapped to DrawingML
  preset geometries where possible (for crispness and editability) and
  custom paths otherwise.
- `<path>` — full command set: `M`, `L`, `H`, `V`, `C`, `S`, `Q`, `T`,
  `A`, `Z`, both cases (absolute/relative). All curves normalize to
  cubic Béziers.
- `<polygon>`, `<polyline>`.
- `<g>` — nested groups with transform composition; single-child groups
  flatten, multi-child groups become `<p:grpSp>`.
- `<text>` with `<tspan>` — becomes an editable DrawingML text box with
  per-run styling. `svg.fonttype='none'` is forced automatically when
  rendering matplotlib figures so glyphs stay as text, not outlines.
- `<defs>`, `<linearGradient>`, `<use>` (with `x`/`y` translation),
  `<clipPath>` (parsed but not geometrically applied — see Limitations).
- `<image>` with `data:image/...;base64,...` URIs — embedded as native
  OOXML picture parts, so `imshow`, colorbar strips, and any other
  raster content survives the round-trip. See the "Raster images"
  subsection under Usage.
- `<style>` with CSS class rules — matches classes on child elements
  and feeds them into the style cascade.

**SVG attributes and styles:**

- `transform` — `translate`, `scale`, `rotate` (with optional center),
  `skewX`, `skewY`, `matrix`. Composes left-to-right per the SVG spec.
- `fill`, `stroke`, `stroke-width`, `stroke-dasharray`,
  `stroke-linecap`, `stroke-linejoin`, `opacity`, `fill-opacity`,
  `stroke-opacity`. Opacity multiplies through the cascade; other
  properties override.
- `font-family`, `font-size`, `font-weight`, `font-style`, `text-anchor`.
- Colors: `#RGB`, `#RRGGBB`, `rgb(r,g,b)`, `rgba(r,g,b,a)`,
  `rgb(50%,50%,50%)`, plus the CSS named colors.

**Output features:**

- EMU coordinate output with configurable source DPI (defaults to 96;
  matplotlib's 72 is handled automatically by the backend).
- Fit-to-box scaling: pass `target_width_emu` / `target_height_emu` to
  stretch the SVG viewBox into a specific EMU region of a slide.
- `prstGeom` emission for axis-aligned rectangles and ellipses (smaller
  output, crisper rendering, labelled as "Rectangle" / "Ellipse" in
  PowerPoint's selection pane).
- `custGeom` fallback with cubic-Bézier approximation for rotated or
  skewed shapes.
- Linear gradients → `<a:gradFill>` with stop-by-stop color + alpha.
- Nested group preservation so users can select and move entire figure
  parts in PowerPoint.
- Templates: open an existing `.pptx`, pick a layout, add slides, insert
  figures — all while preserving the template's branding, master slides,
  and theme.

## Limitations

- **`.docx` output is not yet implemented.** Word wraps shapes in a
  different container (`<wps:wsp>` inside `<w:drawing>`) and needs its
  own rewrite pass. `fig.savefig("x.docx", backend=...)` raises
  `NotImplementedError`.
- **External `<image>` file references are dropped.** Images embedded
  as `data:image/...;base64,...` URIs are fully supported (this is
  what matplotlib emits), but `<image xlink:href="photo.png"/>` style
  references to files on disk are silently skipped — the core crate
  has no filesystem access to resolve them.
- **`<clipPath>` is parsed but not applied.** In practice matplotlib
  clips to axis bounds, which matches our natural output region; if you
  manually tighten `xlim`/`ylim`, data outside the new window may spill
  past the frame. Revisit if it becomes a problem in real use.
- **Text positioning is approximate.** Font metrics differ between
  matplotlib's rendering and Office's, so very long or wrapped text may
  drift by a few pixels. Tick labels, axis titles, and annotations
  (which matplotlib positions individually and we render as one text
  box each) are accurate.
- **3D plots convert** (matplotlib projects them to 2D paths before
  hitting the backend) **but produce large shape counts.** A surface
  plot with 10k facets will give you 10k editable shapes in PowerPoint.
  Consider whether that's actually what you want.
- **No `<pattern>`, `<mask>`, or SVG filters.** matplotlib doesn't emit
  these, so they're not on the v1 roadmap.

## Architecture

`mpl-office` is a mixed Rust + Python project, structured as a Cargo
workspace and a maturin-built Python package:

```
mpl-office/
├── Cargo.toml                    # Rust workspace root
├── pyproject.toml                # Python project + maturin config
├── crates/
│   ├── mpl-office-core/           # Pure-Rust converter
│   │   └── src/
│   │       ├── lib.rs              # Public API
│   │       ├── coord.rs            # EMU math, DPI helpers
│   │       ├── color.rs            # SVG color parsing
│   │       ├── transform.rs        # 2D affine + transform-list parser
│   │       ├── path.rs             # Path tokenize → absolutize → normalize
│   │       ├── style.rs            # Style cascading
│   │       ├── ir.rs               # Intermediate representation
│   │       ├── parse.rs            # quick-xml streaming parser → IR
│   │       └── emit.rs             # IR → DrawingML XML
│   └── mpl-office-py/              # PyO3 bindings → mpl_office._native
│       └── src/lib.rs
└── python/mpl_office/              # Python package
    ├── __init__.py                 # Re-exports from _native
    ├── _inject.py                  # lxml helpers for spTree injection
    ├── pptx.py                     # svg_to_slide, fig_to_slide, fig_to_placeholder
    └── backend.py                  # matplotlib backend
```

The pipeline for `fig.savefig("out.pptx")`:

1. matplotlib renders the figure to an in-memory SVG string
   (`svg.fonttype='none'` is forced so text stays editable).
2. The Rust core parses the SVG via `quick-xml`, builds an intermediate
   tree, normalizes every path command to a cubic Bézier, cascades styles
   through the group hierarchy, and emits DrawingML XML fragments.
3. The Python layer wraps the fragments in a namespace-declaring root,
   parses with `lxml`, and appends the resulting elements to a
   `python-pptx` slide's `<p:spTree>`.
4. `python-pptx` saves the slide as a valid `.pptx` file.

The Rust core is I/O-free and has no Python dependency — you could use
it as a CLI tool or from any language that speaks C-FFI.

## Development

```bash
# Install everything (including dev extras)
uv sync --all-extras

# Rebuild the native extension after editing Rust code
.venv/Scripts/maturin.exe develop --release   # Windows
.venv/bin/maturin develop --release            # macOS/Linux

# Run Rust tests (pure-Rust, no Python needed)
cargo test -p mpl-office-core

# Run Python tests (exercises the full pipeline)
uv run pytest

# Build the example decks
uv run python examples/demo.py
uv run python examples/demo_template.py
```

Open `examples/demo.pptx` and `examples/from_template.pptx` in PowerPoint
to see the output visually.

### Test coverage

| Layer | Count | What it covers |
| --- | --- | --- |
| Rust unit tests | 55 | Color parsing, coord math, affine transforms, path tokenizer, path normalizer, arc-to-cubic, style cascading, parser, emitter, image data-URI decoding |
| Python smoke | 4 | Native extension loads; `ConvertOptions` wiring |
| Python pptx round-trip | 3 | Raw SVG → `.pptx` → re-opened via `python-pptx` |
| Python matplotlib e2e | 3 | Line / bar / scatter figures through the full pipeline |
| Python matplotlib gallery | 6 | Subplots, histogram, log scale, filled area, pie, annotated |
| Python images | 5 | `imshow` heatmaps, raw data-URI round-trip, image deduplication, legacy-API back-compat |
| Python templates | 3 | Template reuse, `fig_to_placeholder`, backend `template=` kwarg |
| Python matplotlib backend | 1 | `fig.savefig("x.pptx", backend=...)` |

All 80 tests run in CI on Linux, macOS, and Windows.

## License

MIT.

## Acknowledgements

The path-normalization algorithm is a machine-rolled-hand-rolled Rust reimplementation
of the approach used in
[`touying-typ/typ2pptx`](https://github.com/touying-typ/typ2pptx)'s
`svg_to_shapes.py`. That project's author deserves credit for proving
out the SVG → DrawingML pipeline in Python and for the clean algorithmic
structure that this project borrowed.

The R ecosystem's [`officer`](https://github.com/davidgohel/officer) +
[`rvg`](https://github.com/davidgohel/rvg) packages set the standard
for native vector graphics in Office documents, and are what I wish
Python had had for the last five years. This project is an attempt to
finally close that gap.
