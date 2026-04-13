//! mpl-office-core
//!
//! SVG → DrawingML converter. Pure Rust, no I/O outside of the optional
//! thin adapters. The crate produces DrawingML XML fragments
//! (`<p:sp>`, `<p:grpSp>`, `<p:pic>`) that can be injected into the `spTree`
//! of a PowerPoint slide or the drawing anchor of a Word document.
//!
//! Pipeline:
//!
//! ```text
//! SVG string  →  parser  →  IR  →  normalizer  →  emitter  →  DrawingML
//! ```
//!
//! Coordinates inside DrawingML are EMU (English Metric Units):
//! 914_400 EMU = 1 inch, 9_525 EMU = 1 px at 96 DPI.

pub mod color;
pub mod coord;
pub mod emit;
pub mod error;
pub mod ir;
pub mod parse;
pub mod path;
pub mod style;
pub mod transform;

pub use error::{Error, Result};

/// Options controlling the SVG → DrawingML conversion.
#[derive(Debug, Clone)]
pub struct ConvertOptions {
    /// DPI assumed for the *source* SVG. matplotlib's SVG backend writes at
    /// 72 DPI but the DrawingML world assumes 96 DPI; `source_dpi` lets the
    /// caller re-scale SVG pixels onto that base.
    pub source_dpi: f64,
    /// Optional target width in EMU. When both target dimensions are set we
    /// scale the SVG's viewBox onto that bounding box instead of converting
    /// pixels → EMU naively.
    pub target_width_emu: Option<i64>,
    /// Optional target height in EMU.
    pub target_height_emu: Option<i64>,
    /// Top-left X offset in EMU on the destination slide/page.
    pub offset_x_emu: i64,
    /// Top-left Y offset in EMU on the destination slide/page.
    pub offset_y_emu: i64,
}

impl Default for ConvertOptions {
    fn default() -> Self {
        Self {
            source_dpi: 96.0,
            target_width_emu: None,
            target_height_emu: None,
            offset_x_emu: 0,
            offset_y_emu: 0,
        }
    }
}

/// Convert an SVG string to one or more DrawingML shape XML fragments.
///
/// Returns the concatenated XML — caller is responsible for wrapping/
/// inserting it into an OOXML document.
pub fn convert_svg_to_drawingml(svg: &str, options: &ConvertOptions) -> Result<String> {
    let document = parse::parse_svg(svg)?;
    let normalized = ir::normalize_document(document);
    let ctx = emit::EmitContext::from_options(&normalized, options);
    Ok(emit::emit_document(&normalized, &ctx))
}
