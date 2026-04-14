//! SVG parser. Uses `quick-xml` for streaming XML tokenization and builds an
//! `IrDocument` tree.
//!
//! Scope: the matplotlib SVG subset described in `Spec.md`:
//! `<svg>`, `<g>`, `<path>`, `<rect>`, `<circle>`, `<ellipse>`, `<line>`,
//! `<polyline>`, `<polygon>`, `<text>` / `<tspan>`, `<defs>`, `<clipPath>`,
//! `<linearGradient>`, `<image>`, `<use>`, `<style>` CSS rules.

use std::collections::HashMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::error::{Error, Result};
use crate::ir::{
    ClipPath, DefEntry, GradientStop, ImageData, IrDocument, LinearGradient, Node, NodeKind,
    TextNode, TextRun,
};
use crate::path::parse_and_normalize;
use crate::style::{parse_length, parse_style_decl, style_from_attrs, Style};
use crate::transform::{parse_transform, Affine};

/// Parse an SVG document from a string.
pub fn parse_svg(xml: &str) -> Result<IrDocument> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = true;

    let mut doc = IrDocument::default();
    let mut css_rules: HashMap<String, Style> = HashMap::new();

    // We walk the event stream with a stack of open elements, building child
    // node lists bottom-up.
    let mut stack: Vec<StackFrame> = Vec::new();
    let mut buf = Vec::new();
    let mut text_buf = String::new();
    let mut in_style = false;
    let mut style_content = String::new();

    let mut svg_seen = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = local_name(e.name().as_ref());
                if tag == "svg" && !svg_seen {
                    svg_seen = true;
                    parse_root_attrs(e, &mut doc)?;
                    stack.push(StackFrame::new("svg".to_string(), Attrs::from_start(e)?));
                    continue;
                }
                if tag == "style" {
                    in_style = true;
                    style_content.clear();
                    continue;
                }
                let attrs = Attrs::from_start(e)?;
                stack.push(StackFrame::new(tag, attrs));
            }
            Ok(Event::End(ref e)) => {
                let tag = local_name(e.name().as_ref());
                if tag == "style" {
                    parse_css_rules(&style_content, &mut css_rules);
                    in_style = false;
                    continue;
                }
                // Pop the matching frame and collect its children into its parent.
                let frame = stack
                    .pop()
                    .ok_or_else(|| Error::Parse(format!("unexpected </{}>", tag)))?;
                if frame.tag != tag {
                    return Err(Error::Parse(format!(
                        "mismatched close: expected </{}>, got </{}>",
                        frame.tag, tag
                    )));
                }

                if frame.tag == "svg" {
                    doc.root = build_group_node(&frame, &css_rules)?;
                    continue;
                }

                // Handle special containers that go into defs
                if frame.tag == "defs" {
                    for (id, entry) in frame.defs_buffer {
                        doc.defs.insert(id, entry);
                    }
                    // Also propagate any gradients/clip-paths defined directly
                    for child in &frame.children {
                        if let Some(id) = &child.id {
                            if let Some(entry) = node_to_def_entry(child) {
                                doc.defs.insert(id.clone(), entry);
                            }
                        }
                    }
                    continue;
                }

                if frame.tag == "text" {
                    let node = build_text_node(&frame, &css_rules)?;
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    }
                    continue;
                }

                if frame.tag == "tspan" {
                    // Handled inline by the parent <text> via text_runs field.
                    if let Some(parent) = stack.last_mut() {
                        parent.text_runs.push(TextRun {
                            text: frame.text_content.clone(),
                            style: resolve_node_style(&frame.attrs, &css_rules),
                            x: frame.attrs.get_f64("x"),
                            y: frame.attrs.get_f64("y"),
                        });
                    }
                    continue;
                }

                if frame.tag == "linearGradient" {
                    let grad = build_linear_gradient(&frame)?;
                    if let Some(id) = frame.attrs.get("id") {
                        let entry = DefEntry::LinearGradient(grad);
                        if let Some(parent) = stack.last_mut() {
                            parent.defs_buffer.insert(id.to_string(), entry);
                        } else {
                            doc.defs.insert(id.to_string(), entry);
                        }
                    }
                    continue;
                }

                if frame.tag == "clipPath" {
                    let clip = ClipPath {
                        children: frame.children.clone(),
                    };
                    if let Some(id) = frame.attrs.get("id") {
                        let entry = DefEntry::ClipPath(clip);
                        if let Some(parent) = stack.last_mut() {
                            parent.defs_buffer.insert(id.to_string(), entry);
                        } else {
                            doc.defs.insert(id.to_string(), entry);
                        }
                    }
                    continue;
                }

                if frame.tag == "stop" {
                    // Stops are consumed by the parent gradient frame directly.
                    if let Some(parent) = stack.last_mut() {
                        if let Some(stop) = parse_stop(&frame.attrs) {
                            parent.gradient_stops.push(stop);
                        }
                    }
                    continue;
                }

                // Normal element — build Node and attach to parent.
                let node_opt = build_element_node(&frame, &css_rules)?;
                if let Some(node) = node_opt {
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if in_style {
                    style_content.push_str(&t.unescape().unwrap_or_default());
                } else if let Some(frame) = stack.last_mut() {
                    let unescaped = t.unescape().unwrap_or_default();
                    frame.text_content.push_str(&unescaped);
                    text_buf.push_str(&unescaped);
                }
            }
            Ok(Event::CData(t)) => {
                if in_style {
                    style_content.push_str(&String::from_utf8_lossy(&t));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(Error::Parse(format!("quick-xml: {}", e))),
        }
        buf.clear();
    }

    Ok(doc)
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

fn local_name(full: &[u8]) -> String {
    // Strip any XML namespace prefix ("svg:rect" → "rect").
    if let Some(pos) = full.iter().position(|b| *b == b':') {
        String::from_utf8_lossy(&full[pos + 1..]).into_owned()
    } else {
        String::from_utf8_lossy(full).into_owned()
    }
}

#[derive(Debug, Default)]
struct Attrs {
    map: HashMap<String, String>,
}

impl Attrs {
    fn from_start(e: &BytesStart<'_>) -> Result<Self> {
        let mut map = HashMap::new();
        for attr in e.attributes() {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            // Decode bytes → str (SVG files are UTF-8). We apply a minimal
            // entity unescape below; for the matplotlib subset this handles
            // every case in practice.
            let raw = std::str::from_utf8(&attr.value)
                .map_err(|err| Error::Parse(format!("attr utf8: {}", err)))?;
            let val = unescape_xml(raw);
            map.insert(key, val);
        }
        Ok(Self { map })
    }

    fn get(&self, k: &str) -> Option<&str> {
        self.map.get(k).map(|s| s.as_str())
    }

    fn get_f64(&self, k: &str) -> Option<f64> {
        self.map.get(k).and_then(|v| parse_length(v))
    }

    fn get_f64_or(&self, k: &str, default: f64) -> f64 {
        self.get_f64(k).unwrap_or(default)
    }
}

#[derive(Debug)]
struct StackFrame {
    tag: String,
    attrs: Attrs,
    children: Vec<Node>,
    text_content: String,
    text_runs: Vec<TextRun>,
    gradient_stops: Vec<GradientStop>,
    defs_buffer: HashMap<String, DefEntry>,
}

impl StackFrame {
    fn new(tag: String, attrs: Attrs) -> Self {
        Self {
            tag,
            attrs,
            children: Vec::new(),
            text_content: String::new(),
            text_runs: Vec::new(),
            gradient_stops: Vec::new(),
            defs_buffer: HashMap::new(),
        }
    }
}

fn parse_root_attrs(e: &BytesStart<'_>, doc: &mut IrDocument) -> Result<()> {
    let attrs = Attrs::from_start(e)?;
    if let Some(v) = attrs.get("viewBox") {
        let nums: Vec<f64> = v
            .split(|c: char| c.is_whitespace() || c == ',')
            .filter(|t| !t.is_empty())
            .filter_map(|t| t.parse::<f64>().ok())
            .collect();
        if nums.len() == 4 {
            doc.view_box = Some((nums[0], nums[1], nums[2], nums[3]));
        }
    }
    doc.width = attrs.get_f64("width");
    doc.height = attrs.get_f64("height");
    Ok(())
}

fn resolve_node_style(attrs: &Attrs, css: &HashMap<String, Style>) -> Style {
    let mut style_map = HashMap::new();
    // CSS classes first (least specific)
    if let Some(classes) = attrs.get("class") {
        for class in classes.split_whitespace() {
            if let Some(cls_style) = css.get(class) {
                // Merge — child (attribute) overrides class.
                return attrs_and_style_to_style(attrs, cls_style, &mut style_map);
            }
        }
    }
    attrs_and_style_to_style(attrs, &Style::default(), &mut style_map)
}

fn attrs_and_style_to_style(
    attrs: &Attrs,
    base: &Style,
    scratch: &mut HashMap<String, String>,
) -> Style {
    scratch.clear();
    if let Some(decl) = attrs.get("style") {
        let parsed = parse_style_decl(decl);
        for (k, v) in parsed {
            scratch.insert(k, v);
        }
    }
    // Direct presentation attributes override style="" per SVG? Actually
    // SVG says `style=""` wins over presentation attrs. But we insert
    // presentation attrs only if not already set from style.
    let presentation = &[
        "fill",
        "stroke",
        "stroke-width",
        "opacity",
        "fill-opacity",
        "stroke-opacity",
        "stroke-dasharray",
        "stroke-linecap",
        "stroke-linejoin",
        "font-family",
        "font-size",
        "font-weight",
        "font-style",
        "text-anchor",
        "clip-path",
    ];
    for k in presentation {
        if !scratch.contains_key(*k) {
            if let Some(v) = attrs.get(k) {
                scratch.insert(k.to_string(), v.to_string());
            }
        }
    }
    let s = style_from_attrs(|k| scratch.get(k).cloned());
    base.cascade(&s)
}

fn parse_css_rules(css: &str, rules: &mut HashMap<String, Style>) {
    // Very small CSS parser: handles `.class { prop: value; }` rules only.
    // Enough for matplotlib's `<style>` block which defines per-class fills.
    let mut i = 0;
    let bytes = css.as_bytes();
    while i < bytes.len() {
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        // Selector
        let sel_start = i;
        while i < bytes.len() && bytes[i] != b'{' {
            i += 1;
        }
        if i >= bytes.len() {
            break;
        }
        let selectors = css[sel_start..i].trim().to_string();
        i += 1; // skip {
        let body_start = i;
        while i < bytes.len() && bytes[i] != b'}' {
            i += 1;
        }
        if i >= bytes.len() {
            break;
        }
        let body = &css[body_start..i];
        i += 1; // skip }

        let decls = parse_style_decl(body);
        let style = style_from_attrs(|k| decls.get(k).cloned());
        for selector in selectors.split(',') {
            let sel = selector.trim();
            if let Some(class) = sel.strip_prefix('.') {
                rules.insert(class.to_string(), style.clone());
            }
        }
    }
}

fn build_group_node(frame: &StackFrame, _css: &HashMap<String, Style>) -> Result<Node> {
    let transform = frame
        .attrs
        .get("transform")
        .map(parse_transform)
        .unwrap_or(Affine::IDENTITY);
    let style = resolve_node_style(&frame.attrs, _css);

    Ok(Node {
        id: frame.attrs.get("id").map(str::to_string),
        transform,
        style,
        kind: NodeKind::Group {
            children: frame.children.clone(),
        },
    })
}

fn build_element_node(frame: &StackFrame, css: &HashMap<String, Style>) -> Result<Option<Node>> {
    let tag = frame.tag.as_str();
    let transform = frame
        .attrs
        .get("transform")
        .map(parse_transform)
        .unwrap_or(Affine::IDENTITY);
    let style = resolve_node_style(&frame.attrs, css);
    let id = frame.attrs.get("id").map(str::to_string);

    let kind = match tag {
        "g" => NodeKind::Group {
            children: frame.children.clone(),
        },
        "path" => {
            let d = frame.attrs.get("d").unwrap_or("");
            let cmds = parse_and_normalize(d)?;
            NodeKind::Path { cmds }
        }
        "rect" => NodeKind::Rect {
            x: frame.attrs.get_f64_or("x", 0.0),
            y: frame.attrs.get_f64_or("y", 0.0),
            w: frame.attrs.get_f64_or("width", 0.0),
            h: frame.attrs.get_f64_or("height", 0.0),
            rx: frame.attrs.get_f64_or("rx", 0.0),
            ry: frame.attrs.get_f64_or("ry", 0.0),
        },
        "circle" => {
            let r = frame.attrs.get_f64_or("r", 0.0);
            NodeKind::Ellipse {
                cx: frame.attrs.get_f64_or("cx", 0.0),
                cy: frame.attrs.get_f64_or("cy", 0.0),
                rx: r,
                ry: r,
            }
        }
        "ellipse" => NodeKind::Ellipse {
            cx: frame.attrs.get_f64_or("cx", 0.0),
            cy: frame.attrs.get_f64_or("cy", 0.0),
            rx: frame.attrs.get_f64_or("rx", 0.0),
            ry: frame.attrs.get_f64_or("ry", 0.0),
        },
        "line" => NodeKind::Line {
            x1: frame.attrs.get_f64_or("x1", 0.0),
            y1: frame.attrs.get_f64_or("y1", 0.0),
            x2: frame.attrs.get_f64_or("x2", 0.0),
            y2: frame.attrs.get_f64_or("y2", 0.0),
        },
        "polygon" => NodeKind::Polygon {
            points: parse_points(frame.attrs.get("points").unwrap_or("")),
        },
        "polyline" => NodeKind::Polyline {
            points: parse_points(frame.attrs.get("points").unwrap_or("")),
        },
        "image" => {
            let href = frame
                .attrs
                .get("href")
                .or_else(|| frame.attrs.get("xlink:href"))
                .unwrap_or("")
                .to_string();
            let data = decode_data_uri(&href);
            NodeKind::Image {
                href,
                x: frame.attrs.get_f64_or("x", 0.0),
                y: frame.attrs.get_f64_or("y", 0.0),
                w: frame.attrs.get_f64_or("width", 0.0),
                h: frame.attrs.get_f64_or("height", 0.0),
                data,
            }
        }
        "use" => NodeKind::Use {
            href: frame
                .attrs
                .get("href")
                .or_else(|| frame.attrs.get("xlink:href"))
                .unwrap_or("")
                .trim_start_matches('#')
                .to_string(),
            x: frame.attrs.get_f64_or("x", 0.0),
            y: frame.attrs.get_f64_or("y", 0.0),
        },
        // Silently skip known non-visual elements
        "title" | "desc" | "metadata" | "style" | "filter" => return Ok(None),
        _ => return Ok(None),
    };

    Ok(Some(Node {
        id,
        transform,
        style,
        kind,
    }))
}

fn build_text_node(frame: &StackFrame, css: &HashMap<String, Style>) -> Result<Node> {
    let transform = frame
        .attrs
        .get("transform")
        .map(parse_transform)
        .unwrap_or(Affine::IDENTITY);
    let style = resolve_node_style(&frame.attrs, css);
    let id = frame.attrs.get("id").map(str::to_string);

    let x = frame.attrs.get_f64_or("x", 0.0);
    let y = frame.attrs.get_f64_or("y", 0.0);

    let mut runs: Vec<TextRun> = Vec::new();
    // Direct text content (before any tspan) becomes the first run.
    let direct = frame.text_content.trim();
    if !direct.is_empty() && frame.text_runs.is_empty() {
        runs.push(TextRun {
            text: normalize_ws(direct),
            style: Style::default(),
            x: None,
            y: None,
        });
    } else if !direct.is_empty() {
        // Mix: keep both — text content was split into tspans inline.
    }
    for r in &frame.text_runs {
        runs.push(r.clone());
    }

    Ok(Node {
        id,
        transform,
        style,
        kind: NodeKind::Text(TextNode { x, y, runs }),
    })
}

fn build_linear_gradient(frame: &StackFrame) -> Result<LinearGradient> {
    fn parse_grad_coord(v: &str, default: f64) -> f64 {
        let v = v.trim();
        if let Some(p) = v.strip_suffix('%') {
            p.trim().parse().unwrap_or(default) / 100.0
        } else {
            v.parse().unwrap_or(default)
        }
    }
    let x1 = frame
        .attrs
        .get("x1")
        .map(|v| parse_grad_coord(v, 0.0))
        .unwrap_or(0.0);
    let y1 = frame
        .attrs
        .get("y1")
        .map(|v| parse_grad_coord(v, 0.0))
        .unwrap_or(0.0);
    let x2 = frame
        .attrs
        .get("x2")
        .map(|v| parse_grad_coord(v, 1.0))
        .unwrap_or(1.0);
    let y2 = frame
        .attrs
        .get("y2")
        .map(|v| parse_grad_coord(v, 0.0))
        .unwrap_or(0.0);
    let user_space = frame
        .attrs
        .get("gradientUnits")
        .map(|v| v == "userSpaceOnUse")
        .unwrap_or(false);
    Ok(LinearGradient {
        x1,
        y1,
        x2,
        y2,
        stops: frame.gradient_stops.clone(),
        user_space_on_use: user_space,
    })
}

fn parse_stop(attrs: &Attrs) -> Option<GradientStop> {
    let offset_str = attrs.get("offset")?;
    let offset = if let Some(p) = offset_str.strip_suffix('%') {
        p.parse::<f64>().ok()? / 100.0
    } else {
        offset_str.parse::<f64>().ok()?
    };

    // stop-color may be direct or inside style=""
    let mut color = attrs.get("stop-color").map(String::from);
    let mut opacity = attrs
        .get("stop-opacity")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(1.0);
    if let Some(decl) = attrs.get("style") {
        let decls = parse_style_decl(decl);
        if color.is_none() {
            color = decls.get("stop-color").cloned();
        }
        if let Some(o) = decls
            .get("stop-opacity")
            .and_then(|v| v.parse::<f64>().ok())
        {
            opacity = o;
        }
    }
    let c = color.as_deref().unwrap_or("#000000");
    let parsed = crate::color::parse_color(c)?;
    let mut hex = [0u8; 6];
    hex.copy_from_slice(parsed.hex().unwrap_or("000000").as_bytes());
    Some(GradientStop {
        offset,
        color: hex,
        opacity,
    })
}

fn node_to_def_entry(_node: &Node) -> Option<DefEntry> {
    // Placeholder: we only accept linearGradient / clipPath explicitly during
    // parsing, so regular nodes under <defs> become reusable symbols.
    Some(DefEntry::Symbol(_node.clone()))
}

fn parse_points(s: &str) -> Vec<(f64, f64)> {
    let nums: Vec<f64> = s
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|t| !t.is_empty())
        .filter_map(|t| t.parse::<f64>().ok())
        .collect();
    nums.chunks_exact(2).map(|c| (c[0], c[1])).collect()
}

/// Decode a `data:image/<fmt>;base64,...` URI into its raw bytes.
///
/// Returns `None` for non-data URIs, unsupported encodings, or malformed
/// payloads. Whitespace inside the base64 payload is tolerated (matplotlib
/// does not add any, but some SVG tools do line-wrap long data URIs).
fn decode_data_uri(href: &str) -> Option<ImageData> {
    let rest = href.strip_prefix("data:image/")?;
    // rest looks like "png;base64,iVBORw0..."
    let (format, after) = rest.split_once(';')?;
    let format = format.trim().to_ascii_lowercase();
    if format.is_empty() {
        return None;
    }
    // Normalise jpeg/jpg so the emitter can pass it straight to PowerPoint.
    let format = if format == "jpg" {
        "jpeg".to_string()
    } else {
        format
    };

    let after = after.trim_start_matches("base64,");
    // Strip any embedded whitespace (some pretty-printers wrap long URIs).
    let cleaned: String = after.chars().filter(|c| !c.is_whitespace()).collect();

    use base64::engine::general_purpose::STANDARD;
    use base64::Engine as _;
    let bytes = STANDARD.decode(cleaned.as_bytes()).ok()?;
    Some(ImageData { bytes, format })
}

fn unescape_xml(s: &str) -> String {
    // Tiny XML entity unescaper — covers the five predefined entities plus
    // numeric decimal/hex references. Falls back to the raw match if the
    // reference is malformed.
    if !s.contains('&') {
        return s.to_string();
    }
    let mut out = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'&' {
            if let Some(semi) = s[i..].find(';') {
                let entity = &s[i + 1..i + semi];
                let replacement = match entity {
                    "lt" => Some('<'),
                    "gt" => Some('>'),
                    "amp" => Some('&'),
                    "quot" => Some('"'),
                    "apos" => Some('\''),
                    e if e.starts_with("#x") || e.starts_with("#X") => {
                        u32::from_str_radix(&e[2..], 16)
                            .ok()
                            .and_then(char::from_u32)
                    }
                    e if e.starts_with('#') => e[1..].parse::<u32>().ok().and_then(char::from_u32),
                    _ => None,
                };
                if let Some(ch) = replacement {
                    out.push(ch);
                    i += semi + 1;
                    continue;
                }
            }
        }
        out.push(s[i..].chars().next().unwrap());
        i += s[i..].chars().next().unwrap().len_utf8();
    }
    out
}

fn normalize_ws(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut prev_ws = false;
    for c in s.chars() {
        if c.is_whitespace() {
            if !prev_ws && !out.is_empty() {
                out.push(' ');
            }
            prev_ws = true;
        } else {
            out.push(c);
            prev_ws = false;
        }
    }
    while out.ends_with(' ') {
        out.pop();
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::path::PathCmd;
    use crate::style::Paint;

    #[test]
    fn parse_empty_svg() {
        let doc =
            parse_svg(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"/>"##)
                .unwrap();
        assert_eq!(doc.width, Some(100.0));
        assert_eq!(doc.height, Some(50.0));
    }

    #[test]
    fn parse_rect() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <rect x="1" y="2" width="30" height="40" fill="#ff0000"/>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        if let NodeKind::Group { children } = &doc.root.kind {
            assert_eq!(children.len(), 1);
            match &children[0].kind {
                NodeKind::Rect { x, y, w, h, .. } => {
                    assert_eq!(*x, 1.0);
                    assert_eq!(*y, 2.0);
                    assert_eq!(*w, 30.0);
                    assert_eq!(*h, 40.0);
                }
                _ => panic!(),
            }
            if let Paint::Color(c) = &children[0].style.fill {
                assert_eq!(c.hex(), Some("FF0000"));
            } else {
                panic!("expected color fill");
            }
        } else {
            panic!("root should be a group");
        }
    }

    #[test]
    fn parse_nested_groups() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <g transform="translate(10,20)">
                <g transform="scale(2)">
                    <rect x="0" y="0" width="5" height="5"/>
                </g>
            </g>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        if let NodeKind::Group { children } = &doc.root.kind {
            assert_eq!(children.len(), 1);
            let outer = &children[0];
            assert_eq!(outer.transform.e, 10.0);
            if let NodeKind::Group { children: inner } = &outer.kind {
                assert_eq!(inner.len(), 1);
                let g2 = &inner[0];
                assert_eq!(g2.transform.a, 2.0);
            } else {
                panic!();
            }
        } else {
            panic!();
        }
    }

    #[test]
    fn parse_style_attr() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <rect x="0" y="0" width="10" height="10" style="fill:#00ff00;stroke:#ff0000;stroke-width:2"/>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        if let NodeKind::Group { children } = &doc.root.kind {
            let rect = &children[0];
            if let Paint::Color(c) = &rect.style.fill {
                assert_eq!(c.hex(), Some("00FF00"));
            } else {
                panic!();
            }
            if let Paint::Color(c) = &rect.style.stroke {
                assert_eq!(c.hex(), Some("FF0000"));
            } else {
                panic!();
            }
            assert_eq!(rect.style.stroke_width, Some(2.0));
        } else {
            panic!();
        }
    }

    #[test]
    fn parse_path_element() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <path d="M0 0 L10 0 L10 10 Z" fill="blue"/>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        if let NodeKind::Group { children } = &doc.root.kind {
            if let NodeKind::Path { cmds } = &children[0].kind {
                assert_eq!(cmds.len(), 4);
                assert!(matches!(cmds[0], PathCmd::MoveTo { .. }));
                assert!(matches!(cmds[3], PathCmd::Close));
            } else {
                panic!();
            }
        }
    }

    #[test]
    fn parse_viewbox() {
        let doc = parse_svg(r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 80"/>"##)
            .unwrap();
        assert_eq!(doc.view_box, Some((0.0, 0.0, 100.0, 80.0)));
    }

    #[test]
    fn parse_css_class() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <style>.foo { fill: #123456 }</style>
            <rect class="foo" width="10" height="10"/>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        if let NodeKind::Group { children } = &doc.root.kind {
            if let Paint::Color(c) = &children[0].style.fill {
                assert_eq!(c.hex(), Some("123456"));
            } else {
                panic!();
            }
        }
    }

    #[test]
    fn parse_linear_gradient() {
        let xml = r##"<svg xmlns="http://www.w3.org/2000/svg">
            <defs>
                <linearGradient id="g1" x1="0%" y1="0%" x2="100%" y2="0%">
                    <stop offset="0%" stop-color="#ff0000"/>
                    <stop offset="100%" stop-color="#0000ff"/>
                </linearGradient>
            </defs>
            <rect width="10" height="10" fill="url(#g1)"/>
        </svg>"##;
        let doc = parse_svg(xml).unwrap();
        assert!(doc.defs.contains_key("g1"));
        if let Some(DefEntry::LinearGradient(g)) = doc.defs.get("g1") {
            assert_eq!(g.stops.len(), 2);
        } else {
            panic!();
        }
    }
}
