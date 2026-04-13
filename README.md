# mpl-office

Native vector graphics from Python into Microsoft Office documents.

`fig.savefig("chart.pptx")` — editable DrawingML shapes, not rasterized PNGs.

## Status

Work in progress. See `Spec.md` for the full design.

## Layout

- `crates/mpl-office-core` — pure Rust SVG → DrawingML converter.
- `crates/mpl-office-py` — PyO3 bindings (builds into `mpl_office._native`).
- `python/mpl_office/` — Python package: `pptx` / `docx` injection and matplotlib backend.

## Build

```bash
uv sync --all-extras
uv run maturin develop --release
uv run pytest
```
