//! Intermediate representation for parsed SVG documents.
//!
//! The parser produces an `IrDocument` — a tree of `Node`s with style
//! inherited from parent groups and every path `d` attribute normalized to
//! a cubic-only `Vec<PathCmd>`.

use crate::path::PathCmd;
use crate::style::Style;
use crate::transform::Affine;

/// Top-level SVG document.
#[derive(Debug, Clone, Default)]
pub struct IrDocument {
    /// ViewBox (x, y, w, h). Defaults to None if not specified — in which
    /// case the parser fills it from `width`/`height` or [0, 0, 0, 0].
    pub view_box: Option<(f64, f64, f64, f64)>,
    /// Document `width` and `height` attributes in user units.
    pub width: Option<f64>,
    pub height: Option<f64>,
    pub root: Node,
    /// Flat definitions registry: `<linearGradient>`, `<clipPath>`, etc.
    /// Stored as raw sub-nodes keyed by `id`.
    pub defs: std::collections::HashMap<String, DefEntry>,
}

#[derive(Debug, Clone)]
pub enum DefEntry {
    LinearGradient(LinearGradient),
    ClipPath(ClipPath),
    /// Reusable symbol/element (for `<use>`).
    Symbol(Node),
}

#[derive(Debug, Clone)]
pub struct LinearGradient {
    pub x1: f64,
    pub y1: f64,
    pub x2: f64,
    pub y2: f64,
    pub stops: Vec<GradientStop>,
    /// `userSpaceOnUse` (false = objectBoundingBox, the SVG default).
    pub user_space_on_use: bool,
}

#[derive(Debug, Clone)]
pub struct GradientStop {
    pub offset: f64,
    pub color: [u8; 6],
    pub opacity: f64,
}

#[derive(Debug, Clone)]
pub struct ClipPath {
    pub children: Vec<Node>,
}

#[derive(Debug, Clone, Default)]
pub struct Node {
    pub id: Option<String>,
    pub transform: Affine,
    pub style: Style,
    pub kind: NodeKind,
}

#[derive(Debug, Clone)]
pub enum NodeKind {
    Group { children: Vec<Node> },
    Path { cmds: Vec<PathCmd> },
    Rect { x: f64, y: f64, w: f64, h: f64, rx: f64, ry: f64 },
    Ellipse { cx: f64, cy: f64, rx: f64, ry: f64 },
    Line { x1: f64, y1: f64, x2: f64, y2: f64 },
    Polygon { points: Vec<(f64, f64)> },
    Polyline { points: Vec<(f64, f64)> },
    Text(TextNode),
    Image { href: String, x: f64, y: f64, w: f64, h: f64 },
    /// `<use>` that the parser couldn't inline (e.g. forward reference);
    /// the emitter will resolve lazily.
    Use { href: String, x: f64, y: f64 },
}

impl Default for NodeKind {
    fn default() -> Self {
        NodeKind::Group { children: Vec::new() }
    }
}

#[derive(Debug, Clone, Default)]
pub struct TextNode {
    pub x: f64,
    pub y: f64,
    /// Each run is a (text, optional per-run style override) pair. This is
    /// flattened from `<text>` + `<tspan>` children during parsing.
    pub runs: Vec<TextRun>,
}

#[derive(Debug, Clone, Default)]
pub struct TextRun {
    pub text: String,
    pub style: Style,
    /// Optional per-run absolute X (for individually positioned glyphs).
    pub x: Option<f64>,
    pub y: Option<f64>,
}

/// Post-parse pass. For now just a hook — the parser already returns a
/// fully-normalized tree. Later we may fold `<use>` references, propagate
/// CSS `<style>` rules, resolve currentColor, etc.
pub fn normalize_document(doc: IrDocument) -> IrDocument {
    doc
}
