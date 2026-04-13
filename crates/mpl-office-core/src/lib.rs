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

pub use emit::EmittedImage;
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
/// inserting it into an OOXML document. Any embedded `<image>` elements
/// are silently dropped; if the SVG contains images you want preserved,
/// use [`convert_svg_to_drawingml_with_images`] instead.
pub fn convert_svg_to_drawingml(svg: &str, options: &ConvertOptions) -> Result<String> {
    let (xml, _images) = convert_svg_to_drawingml_with_images(svg, options)?;
    Ok(xml)
}

/// Convert an SVG string and also return any raster images found inside.
///
/// Each entry in the returned image list carries a **sentinel** string
/// that appears in the XML as `r:embed="{sentinel}"` — the caller is
/// expected to register `bytes` as an image part in the destination
/// OOXML document, obtain a real relationship id, and rewrite every
/// occurrence of the sentinel in the XML to that id.
///
/// Images found in SVG `<image xlink:href="data:image/png;base64,...">`
/// URIs are decoded automatically. External file references are
/// currently dropped (the core crate has no filesystem access).
pub fn convert_svg_to_drawingml_with_images(
    svg: &str,
    options: &ConvertOptions,
) -> Result<(String, Vec<EmittedImage>)> {
    let document = parse::parse_svg(svg)?;
    let normalized = ir::normalize_document(document);
    let ctx = emit::EmitContext::from_options(&normalized, options);
    let xml = emit::emit_document(&normalized, &ctx);
    let images = ctx.images.into_inner();
    Ok((xml, images))
}
