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
    /// Reusable symbol/element (for `<use>`). Boxed because `Node` is ~500
    /// bytes and dominates the enum's layout otherwise.
    Symbol(Box<Node>),
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
    Group {
        children: Vec<Node>,
    },
    Path {
        cmds: Vec<PathCmd>,
    },
    Rect {
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        rx: f64,
        ry: f64,
    },
    Ellipse {
        cx: f64,
        cy: f64,
        rx: f64,
        ry: f64,
    },
    Line {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
    },
    Polygon {
        points: Vec<(f64, f64)>,
    },
    Polyline {
        points: Vec<(f64, f64)>,
    },
    Text(TextNode),
    Image {
        /// Original `href` / `xlink:href` attribute. Kept for diagnostics;
        /// the emitter uses `data` when present.
        href: String,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        /// Decoded raster bytes when the href was a `data:image/...;base64,`
        /// URI. `None` for external file references, which the core crate
        /// doesn't resolve (the Python layer may fill these in later).
        data: Option<ImageData>,
    },
    /// `<use>` that the parser couldn't inline (e.g. forward reference);
    /// the emitter will resolve lazily.
    Use {
        href: String,
        x: f64,
        y: f64,
    },
}

/// Decoded raster image payload extracted from a `data:` URI.
#[derive(Debug, Clone)]
pub struct ImageData {
    pub bytes: Vec<u8>,
    /// `png`, `jpeg`, `gif`, etc. — stored lowercase, without the leading dot.
    pub format: String,
}

impl Default for NodeKind {
    fn default() -> Self {
        NodeKind::Group {
            children: Vec::new(),
        }
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
    /// Vertical alignment relative to the run's baseline, recovered from the
    /// glyph geometry by [`normalize_text_runs`]. SVG has no markup for this
    /// — matplotlib's mathtext expresses a subscript purely as a smaller
    /// font at a shifted `y` — so it is inferred rather than parsed.
    pub vert_align: VertAlign,
}

/// Baseline-relative alignment of a text run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VertAlign {
    #[default]
    Baseline,
    Subscript,
    Superscript,
}

/// Rebuild a text node's runs into the fewest runs that preserve its meaning.
///
/// Matplotlib's SVG backend writes plain text as a single `<text>` with one
/// string, but renders *mathtext* by laying out every glyph itself: one
/// `<tspan>` per character, each with an absolute `x`, and subscripts and
/// superscripts expressed only as a smaller `font-size` at a shifted `y`.
/// Emitted literally that becomes one Office run per glyph, in whatever order
/// the layout engine happened to emit them — `log$_{10}$(PF)` arrives as the
/// nine runs `l o g ( P F ) 1 0`, reading `log(PF)10`.
///
/// This pass recovers the structure:
///
/// 1. Runs with an explicit `x` are sorted into visual reading order, since
///    mathtext emits the subscript after the group it belongs under.
/// 2. The dominant font size is taken as the baseline; a run that is smaller
///    *and* vertically displaced is marked sub- or superscript by the sign of
///    the displacement (SVG `y` grows downward, so lower `y` is a superscript).
/// 3. Adjacent runs that agree on size and alignment are merged back into one.
///
/// Text that was not glyph-split is unaffected: a single run has nothing to
/// sort, nothing to compare against, and nothing to merge.
pub fn normalize_text_runs(runs: Vec<TextRun>) -> Vec<TextRun> {
    if runs.len() < 2 {
        return runs;
    }

    // Only reorder when every run is absolutely positioned, which is what
    // marks matplotlib's glyph-by-glyph mathtext layout. A `<text>` whose
    // tspans flow normally must keep its document order.
    let all_positioned = runs.iter().all(|r| r.x.is_some());
    let mut runs = runs;
    if all_positioned {
        runs.sort_by(|a, b| {
            let ax = a.x.unwrap_or(0.0);
            let bx = b.x.unwrap_or(0.0);
            ax.partial_cmp(&bx).unwrap_or(std::cmp::Ordering::Equal)
        });
    }

    // The baseline is the size most of the text is set in. Taking the maximum
    // rather than the mode keeps a label that is *mostly* subscript (`H_2 O`
    // is half small) from treating the small size as its baseline.
    let base_size = runs
        .iter()
        .filter_map(|r| r.style.font_size)
        .fold(f64::NEG_INFINITY, f64::max);
    let base_y = runs
        .iter()
        .filter(|r| {
            r.style
                .font_size
                .map(|s| (s - base_size).abs() < f64::EPSILON)
                .unwrap_or(false)
        })
        .filter_map(|r| r.y)
        .next();

    if base_size.is_finite() {
        for run in runs.iter_mut() {
            let size = match run.style.font_size {
                Some(s) => s,
                None => continue,
            };
            // A sub/superscript is always set smaller than its base. Requiring
            // that as well as the y-shift avoids misreading an ordinary
            // multi-line or manually-kerned label as a script.
            if size >= base_size * SCRIPT_SIZE_RATIO {
                continue;
            }
            let (y, by) = match (run.y, base_y) {
                (Some(y), Some(by)) => (y, by),
                _ => continue,
            };
            // Tolerance scales with the font so it holds at any figure size.
            let threshold = base_size * SCRIPT_SHIFT_RATIO;
            if y > by + threshold {
                run.vert_align = VertAlign::Subscript;
            } else if y < by - threshold {
                run.vert_align = VertAlign::Superscript;
            }
        }
    }

    // Merge neighbours that would produce identical run properties. Office
    // renders `l` + `o` + `g` and `log` the same, but the merged form is what
    // a human gets when they click into the text box to edit it.
    let mut merged: Vec<TextRun> = Vec::with_capacity(runs.len());
    for run in runs {
        match merged.last_mut() {
            Some(prev) if runs_mergeable(prev, &run) => {
                prev.text.push_str(&run.text);
            }
            _ => merged.push(run),
        }
    }
    merged
}

/// Fraction of the base font size below which a run may be a sub/superscript.
/// Matplotlib sets scripts at 0.7 of their base, so 0.9 separates them
/// cleanly while leaving ordinary size jitter alone.
const SCRIPT_SIZE_RATIO: f64 = 0.9;

/// Vertical displacement, as a fraction of the base font size, beyond which a
/// smaller run counts as shifted off the baseline rather than merely rounded.
const SCRIPT_SHIFT_RATIO: f64 = 0.05;

/// Whether two adjacent runs carry the same formatting and can become one.
fn runs_mergeable(a: &TextRun, b: &TextRun) -> bool {
    if a.vert_align != b.vert_align {
        return false;
    }
    fn same_size(a: Option<f64>, b: Option<f64>) -> bool {
        match (a, b) {
            (Some(x), Some(y)) => (x - y).abs() < 1e-6,
            (None, None) => true,
            _ => false,
        }
    }
    same_size(a.style.font_size, b.style.font_size)
        && a.style.font_family == b.style.font_family
        && a.style.font_weight == b.style.font_weight
        && a.style.font_style == b.style.font_style
        && a.style.fill == b.style.fill
        && a.style.fill_opacity == b.style.fill_opacity
}

/// Post-parse pass: rebuild glyph-split mathtext into structured runs.
/// Beyond that the parser already returns a fully-normalized tree. Later we
/// may fold `<use>` references, propagate CSS `<style>` rules, resolve
/// currentColor, etc.
pub fn normalize_document(mut doc: IrDocument) -> IrDocument {
    fn walk(node: &mut Node) {
        match &mut node.kind {
            NodeKind::Text(text) => {
                let runs = std::mem::take(&mut text.runs);
                text.runs = normalize_text_runs(runs);
            }
            NodeKind::Group { children } => {
                for child in children.iter_mut() {
                    walk(child);
                }
            }
            _ => {}
        }
    }
    walk(&mut doc.root);
    doc
}

#[cfg(test)]
mod tests {
    use super::*;

    /// One glyph as matplotlib's mathtext emits it: absolutely positioned,
    /// with its size on the run itself.
    fn glyph(text: &str, x: f64, y: f64, size: f64) -> TextRun {
        TextRun {
            text: text.to_string(),
            style: Style {
                font_size: Some(size),
                ..Style::default()
            },
            x: Some(x),
            y: Some(y),
            vert_align: VertAlign::Baseline,
        }
    }

    /// `log$_{10}$(PF)` exactly as the SVG backend writes it — note the
    /// subscript glyphs come last but belong at x=15.3.
    fn log10_pf() -> Vec<TextRun> {
        vec![
            glyph("l", 0.0, -0.40625, 10.0),
            glyph("o", 2.77832, -0.40625, 10.0),
            glyph("g", 8.896484, -0.40625, 10.0),
            glyph("(", 24.520508, -0.40625, 10.0),
            glyph("P", 28.421875, -0.40625, 10.0),
            glyph("F", 34.452148, -0.40625, 10.0),
            glyph(")", 40.204102, -0.40625, 10.0),
            glyph("1", 15.339844, 1.093745, 7.0),
            glyph("0", 19.793457, 1.093745, 7.0),
        ]
    }

    #[test]
    fn mathtext_subscript_is_reordered_merged_and_marked() {
        let runs = normalize_text_runs(log10_pf());
        let text: String = runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(text, "log10(PF)");
        assert_eq!(runs.len(), 3, "glyphs should merge into log | 10 | (PF)");
        assert_eq!(runs[0].text, "log");
        assert_eq!(runs[1].text, "10");
        assert_eq!(runs[2].text, "(PF)");
        assert_eq!(runs[1].vert_align, VertAlign::Subscript);
        assert_eq!(runs[0].vert_align, VertAlign::Baseline);
        assert_eq!(runs[2].vert_align, VertAlign::Baseline);
        // Each run keeps the size its glyphs carried.
        assert_eq!(runs[0].style.font_size, Some(10.0));
        assert_eq!(runs[1].style.font_size, Some(7.0));
    }

    #[test]
    fn superscript_is_distinguished_by_sign_of_shift() {
        // `$R^2$`: the script sits above the baseline, so its y is smaller.
        let runs = normalize_text_runs(vec![
            glyph("R", 0.0, -0.809625, 12.0),
            glyph("2", 9.180078, -5.765625, 8.4),
        ]);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].vert_align, VertAlign::Baseline);
        assert_eq!(runs[1].vert_align, VertAlign::Superscript);
    }

    #[test]
    fn plain_single_run_is_untouched() {
        let runs = normalize_text_runs(vec![TextRun {
            text: "log Kow".to_string(),
            style: Style {
                font_size: Some(10.0),
                ..Style::default()
            },
            ..TextRun::default()
        }]);
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "log Kow");
        assert_eq!(runs[0].vert_align, VertAlign::Baseline);
    }

    #[test]
    fn same_size_runs_are_never_marked_as_scripts() {
        // Two runs at a shifted y but the *same* size are manual positioning,
        // not a script — marking them would corrupt ordinary labels.
        let runs =
            normalize_text_runs(vec![glyph("a", 0.0, 0.0, 10.0), glyph("b", 5.0, 4.0, 10.0)]);
        assert!(runs.iter().all(|r| r.vert_align == VertAlign::Baseline));
    }

    #[test]
    fn unpositioned_runs_keep_document_order() {
        // Without an explicit x there is no visual order to recover, so the
        // source order must be preserved rather than guessed at.
        let mut a = glyph("second", 0.0, 0.0, 10.0);
        let mut b = glyph("first", 0.0, 0.0, 10.0);
        a.x = None;
        b.x = None;
        let runs = normalize_text_runs(vec![a, b]);
        let text: String = runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(text, "secondfirst");
    }

    #[test]
    fn differing_styles_are_not_merged() {
        let mut bold = glyph("b", 5.0, 0.0, 10.0);
        bold.style.font_weight = Some("bold".to_string());
        let runs = normalize_text_runs(vec![glyph("a", 0.0, 0.0, 10.0), bold]);
        assert_eq!(runs.len(), 2, "a bold run must stay its own run");
    }
}
