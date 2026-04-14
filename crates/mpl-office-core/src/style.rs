//! Style resolution: inline CSS `style` attribute → `Style` struct,
//! with cascading/inheritance through `<g>` groups.

use crate::color::{parse_color, Color};
use std::collections::HashMap;

/// Fill paint — solid color, gradient reference, or none.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum Paint {
    None,
    Color(Color),
    /// `url(#id)` reference, e.g. to a `<linearGradient>`.
    Ref(String),
    /// Not specified at this level (inherit).
    #[default]
    Inherit,
}

/// Stroke dash kind — either a preset or a raw array.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum DashKind {
    #[default]
    Solid,
    Preset(String),
    Array(Vec<f64>),
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    pub fill: Paint,
    pub stroke: Paint,
    pub stroke_width: Option<f64>,
    pub opacity: Option<f64>,
    pub fill_opacity: Option<f64>,
    pub stroke_opacity: Option<f64>,
    pub stroke_dasharray: Option<String>,
    pub stroke_linecap: Option<String>,
    pub stroke_linejoin: Option<String>,
    pub font_family: Option<String>,
    pub font_size: Option<f64>,
    pub font_weight: Option<String>,
    pub font_style: Option<String>,
    pub text_anchor: Option<String>,
    pub clip_path: Option<String>,
}

impl Style {
    /// Resolve paint into an effective default for rendering (SVG default fill
    /// is black, default stroke is none).
    pub fn effective_fill(&self) -> Paint {
        match &self.fill {
            Paint::Inherit => Paint::Color(parse_color("#000000").unwrap()),
            other => other.clone(),
        }
    }

    pub fn effective_stroke(&self) -> Paint {
        match &self.stroke {
            Paint::Inherit => Paint::None,
            other => other.clone(),
        }
    }

    pub fn effective_fill_opacity(&self) -> f64 {
        let op = self.opacity.unwrap_or(1.0);
        let fop = self.fill_opacity.unwrap_or(1.0);
        (op * fop).clamp(0.0, 1.0)
    }

    pub fn effective_stroke_opacity(&self) -> f64 {
        let op = self.opacity.unwrap_or(1.0);
        let sop = self.stroke_opacity.unwrap_or(1.0);
        (op * sop).clamp(0.0, 1.0)
    }

    pub fn effective_stroke_width(&self) -> f64 {
        self.stroke_width.unwrap_or(1.0)
    }

    /// Apply a child style on top of `self`. Child values override parent,
    /// except for `opacity` / `fill-opacity` / `stroke-opacity` which
    /// multiply (SVG cascading rule).
    pub fn cascade(&self, child: &Style) -> Style {
        fn pick<T: Clone>(parent: &Option<T>, child: &Option<T>) -> Option<T> {
            child.clone().or_else(|| parent.clone())
        }
        fn pick_paint(parent: &Paint, child: &Paint) -> Paint {
            match child {
                Paint::Inherit => parent.clone(),
                other => other.clone(),
            }
        }
        fn mul(parent: &Option<f64>, child: &Option<f64>) -> Option<f64> {
            match (parent, child) {
                (Some(a), Some(b)) => Some(a * b),
                (Some(a), None) => Some(*a),
                (None, Some(b)) => Some(*b),
                (None, None) => None,
            }
        }
        Style {
            fill: pick_paint(&self.fill, &child.fill),
            stroke: pick_paint(&self.stroke, &child.stroke),
            stroke_width: pick(&self.stroke_width, &child.stroke_width),
            opacity: mul(&self.opacity, &child.opacity),
            fill_opacity: mul(&self.fill_opacity, &child.fill_opacity),
            stroke_opacity: mul(&self.stroke_opacity, &child.stroke_opacity),
            stroke_dasharray: pick(&self.stroke_dasharray, &child.stroke_dasharray),
            stroke_linecap: pick(&self.stroke_linecap, &child.stroke_linecap),
            stroke_linejoin: pick(&self.stroke_linejoin, &child.stroke_linejoin),
            font_family: pick(&self.font_family, &child.font_family),
            font_size: pick(&self.font_size, &child.font_size),
            font_weight: pick(&self.font_weight, &child.font_weight),
            font_style: pick(&self.font_style, &child.font_style),
            text_anchor: pick(&self.text_anchor, &child.text_anchor),
            clip_path: pick(&self.clip_path, &child.clip_path),
        }
    }
}

/// Parse `style="fill:red;stroke:blue"` into a map.
pub fn parse_style_decl(s: &str) -> HashMap<String, String> {
    let mut out = HashMap::new();
    for decl in s.split(';') {
        let decl = decl.trim();
        if decl.is_empty() {
            continue;
        }
        if let Some((k, v)) = decl.split_once(':') {
            out.insert(k.trim().to_ascii_lowercase(), v.trim().to_string());
        }
    }
    out
}

/// Build a `Style` from a collection of presentation attributes / style
/// declarations. `lookup` is called for each known property name.
pub fn style_from_attrs<F: Fn(&str) -> Option<String>>(lookup: F) -> Style {
    let mut s = Style::default();

    if let Some(v) = lookup("fill") {
        s.fill = parse_paint(&v);
    }
    if let Some(v) = lookup("stroke") {
        s.stroke = parse_paint(&v);
    }
    if let Some(v) = lookup("stroke-width") {
        s.stroke_width = parse_length(&v);
    }
    if let Some(v) = lookup("opacity") {
        s.opacity = v.parse().ok();
    }
    if let Some(v) = lookup("fill-opacity") {
        s.fill_opacity = v.parse().ok();
    }
    if let Some(v) = lookup("stroke-opacity") {
        s.stroke_opacity = v.parse().ok();
    }
    if let Some(v) = lookup("stroke-dasharray") {
        if v != "none" {
            s.stroke_dasharray = Some(v);
        }
    }
    if let Some(v) = lookup("stroke-linecap") {
        s.stroke_linecap = Some(v);
    }
    if let Some(v) = lookup("stroke-linejoin") {
        s.stroke_linejoin = Some(v);
    }
    if let Some(v) = lookup("font-family") {
        s.font_family = Some(v);
    }
    if let Some(v) = lookup("font-size") {
        s.font_size = parse_length(&v);
    }
    if let Some(v) = lookup("font-weight") {
        s.font_weight = Some(v);
    }
    if let Some(v) = lookup("font-style") {
        s.font_style = Some(v);
    }
    if let Some(v) = lookup("text-anchor") {
        s.text_anchor = Some(v);
    }
    if let Some(v) = lookup("clip-path") {
        s.clip_path = Some(v);
    }

    s
}

fn parse_paint(v: &str) -> Paint {
    let v = v.trim();
    if v.is_empty() || v == "inherit" {
        return Paint::Inherit;
    }
    if v.starts_with("url(") {
        // url(#id)
        if let Some(inner) = v.strip_prefix("url(").and_then(|r| r.strip_suffix(')')) {
            let id = inner.trim().trim_start_matches('#').to_string();
            return Paint::Ref(id);
        }
    }
    match parse_color(v) {
        Some(c) => {
            if c.is_none() {
                Paint::None
            } else {
                Paint::Color(c)
            }
        }
        None => Paint::Inherit,
    }
}

/// Parse an SVG length (plain number or number with px/pt/em/etc. — we strip
/// the unit and return the numeric value). Returns None on failure.
pub fn parse_length(v: &str) -> Option<f64> {
    let v = v.trim();
    if v.is_empty() {
        return None;
    }
    let end = v
        .char_indices()
        .find(|(_, c)| {
            !c.is_ascii_digit() && *c != '.' && *c != '-' && *c != '+' && *c != 'e' && *c != 'E'
        })
        .map(|(i, _)| i)
        .unwrap_or(v.len());
    v[..end].parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_declarations() {
        let m = parse_style_decl("fill:red;stroke:#00ff00;stroke-width:2.5");
        assert_eq!(m.get("fill"), Some(&"red".to_string()));
        assert_eq!(m.get("stroke-width"), Some(&"2.5".to_string()));
    }

    #[test]
    fn style_from_attrs_basic() {
        let m: HashMap<String, String> = [
            ("fill", "#ff0000"),
            ("stroke", "none"),
            ("stroke-width", "2"),
        ]
        .iter()
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect();
        let s = style_from_attrs(|k| m.get(k).cloned());
        assert!(matches!(s.fill, Paint::Color(_)));
        assert!(matches!(s.stroke, Paint::None));
        assert_eq!(s.stroke_width, Some(2.0));
    }

    #[test]
    fn cascade_opacity_multiplies() {
        let parent = Style {
            opacity: Some(0.5),
            ..Default::default()
        };
        let child = Style {
            opacity: Some(0.5),
            ..Default::default()
        };
        let merged = parent.cascade(&child);
        assert_eq!(merged.opacity, Some(0.25));
    }

    #[test]
    fn cascade_fill_overrides() {
        let parent = Style {
            fill: Paint::Color(parse_color("#ff0000").unwrap()),
            ..Default::default()
        };
        let child = Style {
            fill: Paint::Color(parse_color("#00ff00").unwrap()),
            ..Default::default()
        };
        let merged = parent.cascade(&child);
        if let Paint::Color(c) = merged.fill {
            assert_eq!(c.hex(), Some("00FF00"));
        } else {
            panic!();
        }
    }

    #[test]
    fn parse_length_with_unit() {
        assert_eq!(parse_length("12.5px"), Some(12.5));
        assert_eq!(parse_length("3pt"), Some(3.0));
        assert_eq!(parse_length("-1.5"), Some(-1.5));
    }

    #[test]
    fn parse_paint_url_ref() {
        let p = parse_paint("url(#grad1)");
        assert_eq!(p, Paint::Ref("grad1".to_string()));
    }
}
