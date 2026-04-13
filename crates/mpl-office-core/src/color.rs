//! Color parsing: SVG color strings → DrawingML `srgbClr` hex.

/// CSS named colors (the core set used by matplotlib's SVG backend).
/// Full X11 list would be overkill — extend as needed.
const NAMED_COLORS: &[(&str, &str)] = &[
    ("black", "000000"),
    ("white", "FFFFFF"),
    ("red", "FF0000"),
    ("green", "008000"),
    ("blue", "0000FF"),
    ("yellow", "FFFF00"),
    ("cyan", "00FFFF"),
    ("aqua", "00FFFF"),
    ("magenta", "FF00FF"),
    ("fuchsia", "FF00FF"),
    ("gray", "808080"),
    ("grey", "808080"),
    ("silver", "C0C0C0"),
    ("maroon", "800000"),
    ("olive", "808000"),
    ("lime", "00FF00"),
    ("teal", "008080"),
    ("navy", "000080"),
    ("purple", "800080"),
    ("orange", "FFA500"),
    ("none", ""),
    ("transparent", ""),
];

/// Parsed SVG color.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Color {
    /// 6-hex-digit RGB, uppercase. Alpha stored separately.
    Rgb { hex: [u8; 6], alpha: f64 },
    /// Explicit "none" — no paint.
    None,
    /// `currentColor` — resolved from cascade elsewhere.
    CurrentColor,
}

impl Color {
    pub fn hex(&self) -> Option<&str> {
        match self {
            Color::Rgb { hex, .. } => {
                // SAFETY: we only store ASCII bytes
                Some(std::str::from_utf8(hex).unwrap())
            }
            _ => None,
        }
    }

    pub fn alpha(&self) -> f64 {
        match self {
            Color::Rgb { alpha, .. } => *alpha,
            _ => 1.0,
        }
    }

    pub fn is_none(&self) -> bool {
        matches!(self, Color::None)
    }
}

/// Parse an SVG color string into a `Color`.
///
/// Accepts: `#RGB`, `#RRGGBB`, `rgb(r,g,b)`, `rgba(r,g,b,a)`, `none`,
/// and a small set of CSS named colors. Returns `None` on unrecognised input
/// (caller can fall back to element default).
pub fn parse_color(s: &str) -> Option<Color> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    if s.eq_ignore_ascii_case("none") || s.eq_ignore_ascii_case("transparent") {
        return Some(Color::None);
    }
    if s.eq_ignore_ascii_case("currentcolor") {
        return Some(Color::CurrentColor);
    }

    if let Some(rest) = s.strip_prefix('#') {
        return parse_hex(rest);
    }

    if let Some(rest) = s.strip_prefix("rgb(").and_then(|r| r.strip_suffix(')')) {
        return parse_rgb_func(rest, false);
    }
    if let Some(rest) = s.strip_prefix("rgba(").and_then(|r| r.strip_suffix(')')) {
        return parse_rgb_func(rest, true);
    }

    // Named colors
    let lower = s.to_ascii_lowercase();
    for (name, hex) in NAMED_COLORS {
        if *name == lower.as_str() {
            if hex.is_empty() {
                return Some(Color::None);
            }
            return parse_hex(hex);
        }
    }

    None
}

fn parse_hex(raw: &str) -> Option<Color> {
    let raw = raw.trim();
    let expanded: String;
    let hex_str = match raw.len() {
        3 => {
            expanded = raw.chars().flat_map(|c| [c, c]).collect();
            expanded.as_str()
        }
        6 => raw,
        _ => return None,
    };
    if !hex_str.bytes().all(|b| b.is_ascii_hexdigit()) {
        return None;
    }
    let mut hex = [0u8; 6];
    for (i, b) in hex_str.bytes().enumerate() {
        hex[i] = b.to_ascii_uppercase();
    }
    Some(Color::Rgb { hex, alpha: 1.0 })
}

fn parse_rgb_func(inner: &str, with_alpha: bool) -> Option<Color> {
    let parts: Vec<&str> = inner.split(',').map(str::trim).collect();
    let expected = if with_alpha { 4 } else { 3 };
    if parts.len() != expected {
        return None;
    }
    let r = parse_component(parts[0])?;
    let g = parse_component(parts[1])?;
    let b = parse_component(parts[2])?;
    let alpha = if with_alpha {
        parts[3].parse::<f64>().ok()?.clamp(0.0, 1.0)
    } else {
        1.0
    };
    let hex_str = format!("{:02X}{:02X}{:02X}", r, g, b);
    let mut hex = [0u8; 6];
    hex.copy_from_slice(hex_str.as_bytes());
    Some(Color::Rgb { hex, alpha })
}

fn parse_component(s: &str) -> Option<u8> {
    if let Some(pct) = s.strip_suffix('%') {
        let v: f64 = pct.trim().parse().ok()?;
        Some((v / 100.0 * 255.0).round().clamp(0.0, 255.0) as u8)
    } else {
        let v: f64 = s.parse().ok()?;
        Some(v.round().clamp(0.0, 255.0) as u8)
    }
}

/// Format an alpha in [0,1] as a DrawingML per-100000 integer.
#[inline]
pub fn alpha_to_drawingml(a: f64) -> i64 {
    (a.clamp(0.0, 1.0) * 100_000.0).round() as i64
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hex_long() {
        let c = parse_color("#1a2B3c").unwrap();
        assert_eq!(c.hex(), Some("1A2B3C"));
    }

    #[test]
    fn hex_short() {
        let c = parse_color("#fab").unwrap();
        assert_eq!(c.hex(), Some("FFAABB"));
    }

    #[test]
    fn rgb_func() {
        let c = parse_color("rgb(255, 128, 0)").unwrap();
        assert_eq!(c.hex(), Some("FF8000"));
    }

    #[test]
    fn rgba_func() {
        let c = parse_color("rgba(0,0,0,0.5)").unwrap();
        assert_eq!(c.hex(), Some("000000"));
        assert!((c.alpha() - 0.5).abs() < 1e-9);
    }

    #[test]
    fn rgb_percent() {
        let c = parse_color("rgb(100%, 0%, 50%)").unwrap();
        assert_eq!(c.hex(), Some("FF0080"));
    }

    #[test]
    fn named() {
        assert_eq!(parse_color("red").unwrap().hex(), Some("FF0000"));
        assert_eq!(parse_color("Black").unwrap().hex(), Some("000000"));
    }

    #[test]
    fn none_color() {
        assert!(parse_color("none").unwrap().is_none());
    }

    #[test]
    fn alpha_conv() {
        assert_eq!(alpha_to_drawingml(0.5), 50_000);
        assert_eq!(alpha_to_drawingml(1.0), 100_000);
        assert_eq!(alpha_to_drawingml(0.0), 0);
    }
}
