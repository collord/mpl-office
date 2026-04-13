//! Emit DrawingML XML fragments from an `IrDocument`.
//!
//! Output is a concatenated sequence of `<p:sp>` / `<p:grpSp>` / `<p:pic>`
//! elements. Callers wrap the result into the appropriate container
//! (`<p:spTree>` for a slide, `<w:drawing>` for a Word inline anchor).

use std::cell::{Cell, RefCell};
use std::fmt::Write as _;

use crate::color::{alpha_to_drawingml, parse_color, Color};
use crate::coord::{px_to_emu, ANGLE_UNIT, EMU_PER_INCH, FONT_PX_TO_HUNDREDTHS_PT};
use crate::ir::{DefEntry, ImageData, IrDocument, LinearGradient, Node, NodeKind, TextNode, TextRun};
use crate::path::{bbox, transform_cmds, PathCmd};
use crate::style::{Paint, Style};
use crate::transform::Affine;
use crate::ConvertOptions;

/// One image extracted from an `<image>` element and placed on the slide.
///
/// The emitter writes a `<p:pic>` shape with `r:embed="{sentinel}"`; the
/// caller (Python layer) is responsible for registering `bytes` as an image
/// part on the target slide, obtaining the real relationship id, and
/// replacing every occurrence of `sentinel` in the emitted XML with it.
#[derive(Debug, Clone)]
pub struct EmittedImage {
    pub sentinel: String,
    pub bytes: Vec<u8>,
    pub format: String,
}

/// Shared state held across the emission of one document.
pub struct EmitContext {
    pub id_counter: Cell<i64>,
    /// Top-level affine applied to every point *after* SVG transforms.
    /// Used to map the source viewBox onto a target EMU bounding box.
    pub root: Affine,
    /// Additional EMU offset applied to every emitted shape (post-root).
    pub offset_x_emu: i64,
    pub offset_y_emu: i64,
    /// Per-pixel EMU factor to account for `source_dpi != 96`.
    pub emu_per_px: f64,
    /// Defs table (shared borrow into the document).
    pub defs: std::collections::HashMap<String, DefEntry>,
    /// Images emitted during this pass, in the order their sentinels appear
    /// in the XML. Callers drain this after `emit_document` returns.
    pub images: RefCell<Vec<EmittedImage>>,
}

impl EmitContext {
    pub fn from_options(doc: &IrDocument, opts: &ConvertOptions) -> Self {
        let emu_per_px = EMU_PER_INCH as f64 / 96.0; // 9_525 — matches px_to_emu

        // Compute a scale that maps SVG user units onto the requested target
        // bounding box. When no target is given we just rely on source_dpi
        // to convert px → EMU on a 1:1 basis.
        let root = compute_root_affine(doc, opts);

        Self {
            id_counter: Cell::new(2),
            root,
            offset_x_emu: opts.offset_x_emu,
            offset_y_emu: opts.offset_y_emu,
            emu_per_px,
            defs: doc.defs.clone(),
            images: RefCell::new(Vec::new()),
        }
    }

    /// Register an image for emission and return the sentinel rId that the
    /// caller must later rewrite to a real relationship id.
    pub fn next_image_sentinel(&self, data: &ImageData) -> String {
        let mut images = self.images.borrow_mut();
        let n = images.len();
        let sentinel = format!("__mpl_office_img_{n}__");
        images.push(EmittedImage {
            sentinel: sentinel.clone(),
            bytes: data.bytes.clone(),
            format: data.format.clone(),
        });
        sentinel
    }

    pub fn next_id(&self) -> i64 {
        let id = self.id_counter.get();
        self.id_counter.set(id + 1);
        id
    }

    /// Convert a user-unit length to EMU using the current DPI scale.
    pub fn u_to_emu(&self, v: f64) -> i64 {
        (v * self.emu_per_px).round() as i64
    }
}

fn compute_root_affine(doc: &IrDocument, opts: &ConvertOptions) -> Affine {
    // Source size:
    //  - prefer viewBox dimensions if present;
    //  - otherwise fall back to width/height;
    //  - otherwise identity.
    let (src_w, src_h, vb_x, vb_y) = if let Some((x, y, w, h)) = doc.view_box {
        (w, h, x, y)
    } else {
        (doc.width.unwrap_or(0.0), doc.height.unwrap_or(0.0), 0.0, 0.0)
    };

    // Source-DPI rescale (matplotlib writes 72 DPI but DrawingML assumes 96).
    let dpi_scale = 96.0 / opts.source_dpi;

    // If a target bounding box is set we fit-to-box (ignoring aspect).
    let fit_scale_x: f64;
    let fit_scale_y: f64;
    if let (Some(tw), Some(th)) = (opts.target_width_emu, opts.target_height_emu) {
        if src_w > 0.0 && src_h > 0.0 {
            // target EMU / (source px) → final affine scale factor in EMU per user unit
            fit_scale_x = tw as f64 / src_w;
            fit_scale_y = th as f64 / src_h;
            // The rest of the pipeline converts user units → EMU using u_to_emu
            // (which multiplies by EMU_PER_PX). To cancel that out and instead
            // honour fit_scale_*, we embed (fit_scale_x / emu_per_px) into the
            // root affine.
            let emu_per_px = EMU_PER_INCH as f64 / 96.0;
            return Affine::translate(-vb_x, -vb_y)
                .then(Affine::scale(fit_scale_x / emu_per_px, fit_scale_y / emu_per_px));
        }
    }

    // No explicit target — just apply source-DPI correction.
    Affine::translate(-vb_x, -vb_y).then(Affine::scale(dpi_scale, dpi_scale))
}

/// Emit the entire document as a concatenated DrawingML fragment.
pub fn emit_document(doc: &IrDocument, ctx: &EmitContext) -> String {
    let mut out = String::new();
    emit_node(&doc.root, ctx, ctx.root, &Style::default(), &mut out);
    out
}

fn emit_node(node: &Node, ctx: &EmitContext, parent_transform: Affine, parent_style: &Style, out: &mut String) {
    let transform = parent_transform.then(node.transform);
    let style = parent_style.cascade(&node.style);

    match &node.kind {
        NodeKind::Group { children } => {
            let mut inner = String::new();
            for child in children {
                emit_node(child, ctx, transform, &style, &mut inner);
            }
            if inner.is_empty() {
                return;
            }
            // Wrap in <p:grpSp> only when there's >1 shape; otherwise flatten.
            // We measure by counting top-level <p:sp> / <p:pic> / <p:grpSp>.
            if count_top_level_shapes(&inner) <= 1 {
                out.push_str(&inner);
                return;
            }
            let bounds = extract_bounds(&inner).unwrap_or((0, 0, 0, 0));
            let gid = ctx.next_id();
            write!(
                out,
                concat!(
                    "<p:grpSp>",
                    "<p:nvGrpSpPr>",
                    "<p:cNvPr id=\"{gid}\" name=\"Group {gid}\"/>",
                    "<p:cNvGrpSpPr/><p:nvPr/>",
                    "</p:nvGrpSpPr>",
                    "<p:grpSpPr>",
                    "<a:xfrm>",
                    "<a:off x=\"{x}\" y=\"{y}\"/>",
                    "<a:ext cx=\"{w}\" cy=\"{h}\"/>",
                    "<a:chOff x=\"{x}\" y=\"{y}\"/>",
                    "<a:chExt cx=\"{w}\" cy=\"{h}\"/>",
                    "</a:xfrm>",
                    "</p:grpSpPr>",
                    "{inner}",
                    "</p:grpSp>"
                ),
                gid = gid,
                x = bounds.0,
                y = bounds.1,
                w = bounds.2.max(1),
                h = bounds.3.max(1),
                inner = inner
            )
            .unwrap();
        }
        NodeKind::Rect { x, y, w, h, rx, ry } => {
            emit_rect(ctx, transform, &style, *x, *y, *w, *h, *rx, *ry, out);
        }
        NodeKind::Ellipse { cx, cy, rx, ry } => {
            emit_ellipse(ctx, transform, &style, *cx, *cy, *rx, *ry, out);
        }
        NodeKind::Line { x1, y1, x2, y2 } => {
            emit_line(ctx, transform, &style, *x1, *y1, *x2, *y2, out);
        }
        NodeKind::Polygon { points } => {
            emit_poly(ctx, transform, &style, points, true, out);
        }
        NodeKind::Polyline { points } => {
            emit_poly(ctx, transform, &style, points, false, out);
        }
        NodeKind::Path { cmds } => {
            emit_path(ctx, transform, &style, cmds, out);
        }
        NodeKind::Text(t) => {
            emit_text(ctx, transform, &style, t, out);
        }
        NodeKind::Image { x, y, w, h, data, .. } => {
            if let Some(img) = data {
                emit_image(ctx, transform, *x, *y, *w, *h, img, out);
            }
            // External file references with `data == None` are dropped — the
            // core crate has no filesystem access to resolve them. A future
            // enhancement could let the Python layer supply them through
            // `ConvertOptions`.
        }
        NodeKind::Use { href, x, y } => {
            if let Some(DefEntry::Symbol(sym)) = ctx.defs.get(href) {
                // A `<use>`'s x/y attributes translate the referenced symbol
                // (they are NOT part of its `transform` attribute).
                let use_shift = Affine::translate(*x, *y);
                emit_node(sym, ctx, transform.then(use_shift), &style, out);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Shape emitters
// ---------------------------------------------------------------------------

#[allow(clippy::too_many_arguments)]
fn emit_rect(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    x: f64, y: f64, w: f64, h: f64,
    rx: f64, ry: f64,
    out: &mut String,
) {
    if w <= 0.0 || h <= 0.0 {
        return;
    }

    // When the effective transform is translate+uniform-scale only, we can
    // use a DrawingML preset rect geometry (which is crisper and editable).
    // Otherwise we fall back to a path.
    if t.is_axis_aligned_scale() && rx == 0.0 && ry == 0.0 {
        let (x0, y0) = t.transform_point(x, y);
        let (x1, y1) = t.transform_point(x + w, y + h);
        let emu_x = ctx.u_to_emu(x0.min(x1)) + ctx.offset_x_emu;
        let emu_y = ctx.u_to_emu(y0.min(y1)) + ctx.offset_y_emu;
        let emu_w = (ctx.u_to_emu(x1) - ctx.u_to_emu(x0)).abs().max(1);
        let emu_h = (ctx.u_to_emu(y1) - ctx.u_to_emu(y0)).abs().max(1);
        let geom = "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>";
        let fill = build_fill(&style.effective_fill(), style.effective_fill_opacity(), &ctx.defs);
        let stroke = build_stroke(&style.effective_stroke(), style.effective_stroke_opacity(),
                                  style.effective_stroke_width() * scale_factor(&t), style);
        let id = ctx.next_id();
        write_shape(out, id, "Rectangle", emu_x, emu_y, emu_w, emu_h, 0, geom, &fill, &stroke, "");
        return;
    }

    // Fallback: build a closed path and emit as a custom geom.
    let cmds = vec![
        PathCmd::MoveTo { x, y },
        PathCmd::LineTo { x: x + w, y },
        PathCmd::LineTo { x: x + w, y: y + h },
        PathCmd::LineTo { x, y: y + h },
        PathCmd::Close,
    ];
    emit_path(ctx, t, style, &cmds, out);
}

fn emit_ellipse(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    cx: f64, cy: f64, rx: f64, ry: f64,
    out: &mut String,
) {
    if rx <= 0.0 || ry <= 0.0 {
        return;
    }
    if t.is_axis_aligned_scale() {
        let (x0, y0) = t.transform_point(cx - rx, cy - ry);
        let (x1, y1) = t.transform_point(cx + rx, cy + ry);
        let emu_x = ctx.u_to_emu(x0.min(x1)) + ctx.offset_x_emu;
        let emu_y = ctx.u_to_emu(y0.min(y1)) + ctx.offset_y_emu;
        let emu_w = (ctx.u_to_emu(x1) - ctx.u_to_emu(x0)).abs().max(1);
        let emu_h = (ctx.u_to_emu(y1) - ctx.u_to_emu(y0)).abs().max(1);
        let geom = "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>";
        let fill = build_fill(&style.effective_fill(), style.effective_fill_opacity(), &ctx.defs);
        let stroke = build_stroke(&style.effective_stroke(), style.effective_stroke_opacity(),
                                  style.effective_stroke_width() * scale_factor(&t), style);
        let id = ctx.next_id();
        write_shape(out, id, "Ellipse", emu_x, emu_y, emu_w, emu_h, 0, geom, &fill, &stroke, "");
        return;
    }

    // Approximate an ellipse with four cubic Béziers and emit as custom path.
    const K: f64 = 0.5522847498307936;
    let ox = rx * K;
    let oy = ry * K;
    let cmds = vec![
        PathCmd::MoveTo { x: cx + rx, y: cy },
        PathCmd::CubicTo { x1: cx + rx, y1: cy + oy, x2: cx + ox, y2: cy + ry, x: cx, y: cy + ry },
        PathCmd::CubicTo { x1: cx - ox, y1: cy + ry, x2: cx - rx, y2: cy + oy, x: cx - rx, y: cy },
        PathCmd::CubicTo { x1: cx - rx, y1: cy - oy, x2: cx - ox, y2: cy - ry, x: cx, y: cy - ry },
        PathCmd::CubicTo { x1: cx + ox, y1: cy - ry, x2: cx + rx, y2: cy - oy, x: cx + rx, y: cy },
        PathCmd::Close,
    ];
    emit_path(ctx, t, style, &cmds, out);
}

fn emit_line(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    x1: f64, y1: f64, x2: f64, y2: f64,
    out: &mut String,
) {
    let cmds = vec![
        PathCmd::MoveTo { x: x1, y: y1 },
        PathCmd::LineTo { x: x2, y: y2 },
    ];
    // Lines ignore fill.
    let mut style = style.clone();
    style.fill = Paint::None;
    emit_path(ctx, t, &style, &cmds, out);
}

fn emit_poly(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    points: &[(f64, f64)],
    closed: bool,
    out: &mut String,
) {
    if points.len() < 2 {
        return;
    }
    let mut cmds = Vec::with_capacity(points.len() + 1);
    cmds.push(PathCmd::MoveTo { x: points[0].0, y: points[0].1 });
    for (px, py) in &points[1..] {
        cmds.push(PathCmd::LineTo { x: *px, y: *py });
    }
    let mut style = style.clone();
    if closed {
        cmds.push(PathCmd::Close);
    } else {
        // Polyline has no fill.
        style.fill = Paint::None;
    }
    emit_path(ctx, t, &style, &cmds, out);
}

fn emit_path(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    cmds: &[PathCmd],
    out: &mut String,
) {
    if cmds.is_empty() {
        return;
    }
    let transformed = transform_cmds(cmds, &t);
    let (min_x, min_y, w, h) = bbox(&transformed);
    if w <= 0.0 && h <= 0.0 {
        return;
    }

    let emu_x = ctx.u_to_emu(min_x) + ctx.offset_x_emu;
    let emu_y = ctx.u_to_emu(min_y) + ctx.offset_y_emu;
    let emu_w = ctx.u_to_emu(w).max(1);
    let emu_h = ctx.u_to_emu(h).max(1);

    // Local-coordinate path XML (relative to the shape's bounding box).
    let mut path_inner = String::new();
    for c in &transformed {
        match c {
            PathCmd::MoveTo { x, y } => {
                let lx = ctx.u_to_emu(x - min_x);
                let ly = ctx.u_to_emu(y - min_y);
                write!(path_inner, "<a:moveTo><a:pt x=\"{lx}\" y=\"{ly}\"/></a:moveTo>").unwrap();
            }
            PathCmd::LineTo { x, y } => {
                let lx = ctx.u_to_emu(x - min_x);
                let ly = ctx.u_to_emu(y - min_y);
                write!(path_inner, "<a:lnTo><a:pt x=\"{lx}\" y=\"{ly}\"/></a:lnTo>").unwrap();
            }
            PathCmd::CubicTo { x1, y1, x2, y2, x, y } => {
                let c1x = ctx.u_to_emu(x1 - min_x);
                let c1y = ctx.u_to_emu(y1 - min_y);
                let c2x = ctx.u_to_emu(x2 - min_x);
                let c2y = ctx.u_to_emu(y2 - min_y);
                let ex = ctx.u_to_emu(x - min_x);
                let ey = ctx.u_to_emu(y - min_y);
                write!(
                    path_inner,
                    "<a:cubicBezTo><a:pt x=\"{c1x}\" y=\"{c1y}\"/><a:pt x=\"{c2x}\" y=\"{c2y}\"/><a:pt x=\"{ex}\" y=\"{ey}\"/></a:cubicBezTo>"
                )
                .unwrap();
            }
            PathCmd::Close => path_inner.push_str("<a:close/>"),
        }
    }

    let geom = format!(
        "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"l\" t=\"t\" r=\"r\" b=\"b\"/><a:pathLst><a:path w=\"{emu_w}\" h=\"{emu_h}\">{path_inner}</a:path></a:pathLst></a:custGeom>",
    );

    let fill = build_fill(&style.effective_fill(), style.effective_fill_opacity(), &ctx.defs);
    let sw_scaled = style.effective_stroke_width() * scale_factor(&t);
    let stroke = build_stroke(&style.effective_stroke(), style.effective_stroke_opacity(), sw_scaled, style);

    let id = ctx.next_id();
    write_shape(out, id, "Freeform", emu_x, emu_y, emu_w, emu_h, 0, &geom, &fill, &stroke, "");
}

fn emit_image(
    ctx: &EmitContext,
    t: Affine,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    data: &ImageData,
    out: &mut String,
) {
    if w <= 0.0 || h <= 0.0 {
        return;
    }

    // Transform the four corners. matplotlib's imshow typically emits a
    // translate+scale transform (often with a negative Y scale to flip the
    // image), so we compute the axis-aligned bounding box of the transformed
    // rect and drop any rotation component for now. Rotated images would
    // need DrawingML's `rot` attribute, which we can add later.
    let (x0, y0) = t.transform_point(x, y);
    let (x1, y1) = t.transform_point(x + w, y);
    let (x2, y2) = t.transform_point(x + w, y + h);
    let (x3, y3) = t.transform_point(x, y + h);
    let min_x = x0.min(x1).min(x2).min(x3);
    let max_x = x0.max(x1).max(x2).max(x3);
    let min_y = y0.min(y1).min(y2).min(y3);
    let max_y = y0.max(y1).max(y2).max(y3);

    let emu_x = ctx.u_to_emu(min_x) + ctx.offset_x_emu;
    let emu_y = ctx.u_to_emu(min_y) + ctx.offset_y_emu;
    let emu_w = (ctx.u_to_emu(max_x - min_x)).max(1);
    let emu_h = (ctx.u_to_emu(max_y - min_y)).max(1);

    // If the effective Y-scale is negative (imshow's default), tell DrawingML
    // to flip the image vertically so it renders the right way up.
    let flip_v = t.d < 0.0;
    let flip_h = t.a < 0.0;
    let mut flip_attrs = String::new();
    if flip_h {
        flip_attrs.push_str(" flipH=\"1\"");
    }
    if flip_v {
        flip_attrs.push_str(" flipV=\"1\"");
    }

    let sentinel = ctx.next_image_sentinel(data);
    let id = ctx.next_id();

    write!(
        out,
        concat!(
            "<p:pic>",
            "<p:nvPicPr>",
            "<p:cNvPr id=\"{id}\" name=\"Image {id}\"/>",
            "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>",
            "<p:nvPr/>",
            "</p:nvPicPr>",
            "<p:blipFill>",
            "<a:blip r:embed=\"{sentinel}\"/>",
            "<a:stretch><a:fillRect/></a:stretch>",
            "</p:blipFill>",
            "<p:spPr>",
            "<a:xfrm{flip_attrs}>",
            "<a:off x=\"{x}\" y=\"{y}\"/>",
            "<a:ext cx=\"{w}\" cy=\"{h}\"/>",
            "</a:xfrm>",
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
            "</p:spPr>",
            "</p:pic>"
        ),
        id = id,
        sentinel = sentinel,
        flip_attrs = flip_attrs,
        x = emu_x,
        y = emu_y,
        w = emu_w,
        h = emu_h,
    )
    .unwrap();
}

fn emit_text(
    ctx: &EmitContext,
    t: Affine,
    style: &Style,
    text: &TextNode,
    out: &mut String,
) {
    // We map text anchor position through the transform; font sizing uses
    // the transform's overall scale so rotated labels still read correctly.
    let (tx, ty) = t.transform_point(text.x, text.y);
    let scale = scale_factor(&t);
    let font_size = style.font_size.unwrap_or(16.0) * scale;
    let sz = (font_size * FONT_PX_TO_HUNDREDTHS_PT).round() as i64;

    // Very crude width estimate — matches the reference implementation.
    let full_text: String = text.runs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>().join("");
    if full_text.trim().is_empty() {
        return;
    }
    let width_px = estimate_text_width(&full_text, font_size, style.font_weight.as_deref());
    let padding = font_size * 0.1;
    let height_px = font_size * 1.5;
    let text_anchor = style.text_anchor.as_deref().unwrap_or("start");

    let box_x = match text_anchor {
        "middle" => tx - width_px / 2.0 - padding,
        "end" => tx - width_px - padding,
        _ => tx - padding,
    };
    let box_y = ty - font_size * 0.85;
    let box_w = width_px + padding * 2.0;
    let box_h = height_px + padding;

    let emu_x = ctx.u_to_emu(box_x) + ctx.offset_x_emu;
    let emu_y = ctx.u_to_emu(box_y) + ctx.offset_y_emu;
    let emu_w = ctx.u_to_emu(box_w).max(1);
    let emu_h = ctx.u_to_emu(box_h).max(1);

    // Rotation extracted from the affine.
    let rot_rad = t.b.atan2(t.a);
    let rot = (rot_rad.to_degrees() * ANGLE_UNIT as f64).round() as i64;

    let fill_color = match style.effective_fill() {
        Paint::Color(c) => c,
        _ => parse_color("#000000").unwrap(),
    };

    let algn = match text_anchor {
        "middle" => "ctr",
        "end" => "r",
        _ => "l",
    };

    let mut runs_xml = String::new();
    for run in &text.runs {
        write_text_run(&mut runs_xml, run, style, fill_color, sz);
    }
    if runs_xml.is_empty() {
        // Single-run fallback when runs are empty but text has content.
        let default_run = TextRun { text: full_text.clone(), style: Style::default(), x: None, y: None };
        write_text_run(&mut runs_xml, &default_run, style, fill_color, sz);
    }

    let id = ctx.next_id();
    write!(out,
        concat!(
            "<p:sp>",
            "<p:nvSpPr>",
            "<p:cNvPr id=\"{id}\" name=\"TextBox {id}\"/>",
            "<p:cNvSpPr txBox=\"1\"/><p:nvPr/>",
            "</p:nvSpPr>",
            "<p:spPr>",
            "<a:xfrm{rot_attr}>",
            "<a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/>",
            "</a:xfrm>",
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
            "<a:noFill/><a:ln><a:noFill/></a:ln>",
            "</p:spPr>",
            "<p:txBody>",
            "<a:bodyPr wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"t\" anchorCtr=\"0\"><a:spAutoFit/></a:bodyPr>",
            "<a:lstStyle/>",
            "<a:p><a:pPr algn=\"{algn}\"/>{runs}</a:p>",
            "</p:txBody>",
            "</p:sp>"
        ),
        id = id,
        rot_attr = if rot != 0 { format!(" rot=\"{}\"", rot) } else { String::new() },
        x = emu_x,
        y = emu_y,
        w = emu_w,
        h = emu_h,
        algn = algn,
        runs = runs_xml,
    ).unwrap();
}

fn write_text_run(
    out: &mut String,
    run: &TextRun,
    parent_style: &Style,
    default_color: Color,
    sz: i64,
) {
    let cascaded = parent_style.cascade(&run.style);
    let color = match cascaded.effective_fill() {
        Paint::Color(c) => c,
        _ => default_color,
    };
    let hex = color.hex().unwrap_or("000000");
    let fw = cascaded.font_weight.as_deref().unwrap_or("400");
    let bold = matches!(fw, "bold" | "600" | "700" | "800" | "900");
    let italic = cascaded.font_style.as_deref() == Some("italic");
    let family = cascaded.font_family.as_deref().unwrap_or("Segoe UI");
    let latin = pick_latin_font(family);

    let text_escaped = xml_escape(&run.text);

    let alpha = cascaded.effective_fill_opacity();
    let alpha_xml = if alpha < 1.0 {
        format!("<a:alpha val=\"{}\"/>", alpha_to_drawingml(alpha))
    } else {
        String::new()
    };

    write!(out,
        "<a:r><a:rPr lang=\"en-US\" sz=\"{sz}\"{b}{i} dirty=\"0\"><a:solidFill><a:srgbClr val=\"{hex}\">{alpha_xml}</a:srgbClr></a:solidFill><a:latin typeface=\"{latin}\"/><a:cs typeface=\"{latin}\"/></a:rPr><a:t>{text}</a:t></a:r>",
        sz = sz,
        b = if bold { " b=\"1\"" } else { "" },
        i = if italic { " i=\"1\"" } else { "" },
        hex = hex,
        alpha_xml = alpha_xml,
        latin = xml_escape(latin),
        text = text_escaped,
    ).unwrap();
}

fn pick_latin_font(family: &str) -> &str {
    // Take first font in the list, strip quotes.
    family
        .split(',')
        .next()
        .map(|s| s.trim().trim_matches(|c| c == '\'' || c == '"'))
        .filter(|s| !s.is_empty())
        .unwrap_or("Segoe UI")
}

fn estimate_text_width(text: &str, font_size: f64, font_weight: Option<&str>) -> f64 {
    let mut w = 0.0;
    for ch in text.chars() {
        w += match ch {
            ' ' => font_size * 0.3,
            'm' | 'M' | 'w' | 'W' | 'O' | 'Q' => font_size * 0.75,
            'i' | 'I' | 'l' | 'j' | '1' | '!' | '|' => font_size * 0.3,
            _ => font_size * 0.55,
        };
    }
    if matches!(font_weight, Some("bold" | "600" | "700" | "800" | "900")) {
        w *= 1.05;
    }
    w * 1.15
}

// ---------------------------------------------------------------------------
// Fill / stroke builders
// ---------------------------------------------------------------------------

fn build_fill(
    paint: &Paint,
    opacity: f64,
    defs: &std::collections::HashMap<String, DefEntry>,
) -> String {
    match paint {
        Paint::None => "<a:noFill/>".to_string(),
        Paint::Color(c) => {
            let hex = c.hex().unwrap_or("000000");
            let effective = opacity * c.alpha();
            if effective < 1.0 {
                let a = alpha_to_drawingml(effective);
                format!("<a:solidFill><a:srgbClr val=\"{hex}\"><a:alpha val=\"{a}\"/></a:srgbClr></a:solidFill>")
            } else {
                format!("<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>")
            }
        }
        Paint::Ref(id) => {
            if let Some(DefEntry::LinearGradient(g)) = defs.get(id) {
                build_linear_gradient_fill(g, opacity)
            } else {
                "<a:noFill/>".to_string()
            }
        }
        Paint::Inherit => {
            // Default SVG fill is black.
            "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>".to_string()
        }
    }
}

fn build_linear_gradient_fill(g: &LinearGradient, opacity: f64) -> String {
    let mut stops = String::new();
    for s in &g.stops {
        let pos = (s.offset * 100_000.0).round() as i64;
        let hex = std::str::from_utf8(&s.color).unwrap_or("000000");
        let a = alpha_to_drawingml((s.opacity * opacity).clamp(0.0, 1.0));
        if a < 100_000 {
            write!(
                stops,
                "<a:gs pos=\"{pos}\"><a:srgbClr val=\"{hex}\"><a:alpha val=\"{a}\"/></a:srgbClr></a:gs>"
            ).unwrap();
        } else {
            write!(
                stops,
                "<a:gs pos=\"{pos}\"><a:srgbClr val=\"{hex}\"/></a:gs>"
            ).unwrap();
        }
    }
    if stops.is_empty() {
        return "<a:noFill/>".to_string();
    }
    let angle_rad = (g.y2 - g.y1).atan2(g.x2 - g.x1);
    let ang = ((angle_rad.to_degrees().rem_euclid(360.0)) * ANGLE_UNIT as f64).round() as i64;
    format!("<a:gradFill><a:gsLst>{stops}</a:gsLst><a:lin ang=\"{ang}\" scaled=\"1\"/></a:gradFill>")
}

fn build_stroke(
    paint: &Paint,
    opacity: f64,
    width_px: f64,
    style: &Style,
) -> String {
    match paint {
        Paint::None => "<a:ln><a:noFill/></a:ln>".to_string(),
        Paint::Color(c) => {
            let hex = c.hex().unwrap_or("000000");
            let emu_w = px_to_emu(width_px);
            let effective = opacity * c.alpha();
            let alpha_xml = if effective < 1.0 {
                format!("<a:alpha val=\"{}\"/>", alpha_to_drawingml(effective))
            } else {
                String::new()
            };
            let dash = build_dash(&style.stroke_dasharray);
            let cap_attr = match style.stroke_linecap.as_deref() {
                Some("round") => " cap=\"rnd\"",
                Some("square") => " cap=\"sq\"",
                Some("butt") => " cap=\"flat\"",
                _ => "",
            };
            let join = match style.stroke_linejoin.as_deref() {
                Some("round") => "<a:round/>",
                Some("bevel") => "<a:bevel/>",
                Some("miter") => "<a:miter lim=\"800000\"/>",
                _ => "",
            };
            format!(
                "<a:ln w=\"{emu_w}\"{cap_attr}><a:solidFill><a:srgbClr val=\"{hex}\">{alpha_xml}</a:srgbClr></a:solidFill>{join}{dash}</a:ln>"
            )
        }
        Paint::Ref(_) => "<a:ln><a:noFill/></a:ln>".to_string(),
        Paint::Inherit => "<a:ln><a:noFill/></a:ln>".to_string(),
    }
}

fn build_dash(dash: &Option<String>) -> String {
    match dash.as_deref() {
        None | Some("") | Some("none") => String::new(),
        Some(d) => {
            let preset = match d.trim() {
                "4,4" | "4 4" => "dash",
                "6,3" | "6 3" => "dash",
                "2,2" | "2 2" => "sysDot",
                "8,4" | "8 4" => "lgDash",
                "8,4,2,4" | "8 4 2 4" => "lgDashDot",
                _ => "dash",
            };
            format!("<a:prstDash val=\"{preset}\"/>")
        }
    }
}

// ---------------------------------------------------------------------------
// Shape wrapper
// ---------------------------------------------------------------------------

#[allow(clippy::too_many_arguments)]
fn write_shape(
    out: &mut String,
    id: i64,
    name: &str,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    rot: i64,
    geom: &str,
    fill: &str,
    stroke: &str,
    extra: &str,
) {
    let rot_attr = if rot != 0 { format!(" rot=\"{}\"", rot) } else { String::new() };
    write!(
        out,
        concat!(
            "<p:sp>",
            "<p:nvSpPr>",
            "<p:cNvPr id=\"{id}\" name=\"{name} {id}\"/>",
            "<p:cNvSpPr/><p:nvPr/>",
            "</p:nvSpPr>",
            "<p:spPr>",
            "<a:xfrm{rot_attr}>",
            "<a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/>",
            "</a:xfrm>",
            "{geom}{fill}{stroke}{extra}",
            "</p:spPr>",
            "</p:sp>"
        ),
        id = id,
        name = xml_escape(name),
        rot_attr = rot_attr,
        x = x,
        y = y,
        cx = cx,
        cy = cy,
        geom = geom,
        fill = fill,
        stroke = stroke,
        extra = extra,
    )
    .unwrap();
}

// ---------------------------------------------------------------------------
// Small utilities
// ---------------------------------------------------------------------------

fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

fn scale_factor(t: &Affine) -> f64 {
    // Geometric-mean singular value — good enough for stroke widths and
    // font sizes under a uniform scale, and reasonable for mild skew.
    let sx = (t.a * t.a + t.b * t.b).sqrt();
    let sy = (t.c * t.c + t.d * t.d).sqrt();
    ((sx * sy).max(1e-9)).sqrt()
}

fn count_top_level_shapes(s: &str) -> usize {
    // Count top-level <p:sp>, <p:grpSp>, <p:pic> by scanning depth.
    let mut count = 0usize;
    let bytes = s.as_bytes();
    let mut i = 0;
    let mut depth = 0i32;
    while i + 1 < bytes.len() {
        if bytes[i] == b'<' {
            let rest = &s[i..];
            if rest.starts_with("</p:sp>") || rest.starts_with("</p:grpSp>") || rest.starts_with("</p:pic>") {
                depth -= 1;
                if depth == 0 { /* end of top-level element already counted */ }
                i += 1;
                continue;
            }
            if rest.starts_with("<p:sp>") || rest.starts_with("<p:grpSp>") || rest.starts_with("<p:pic>") {
                if depth == 0 {
                    count += 1;
                }
                depth += 1;
                i += 1;
                continue;
            }
        }
        i += 1;
    }
    count
}

/// Extract the outer bounding box (in EMU) from a chunk of already-emitted
/// shape XML — used for group wrapping.
fn extract_bounds(s: &str) -> Option<(i64, i64, i64, i64)> {
    // Find every top-level <a:off .../><a:ext .../> pair, accumulate min/max.
    // This is a scan, not a full XML parse — acceptable because we produce
    // the XML ourselves and know the layout.
    let mut min_x = i64::MAX;
    let mut min_y = i64::MAX;
    let mut max_x = i64::MIN;
    let mut max_y = i64::MIN;
    let mut i = 0;
    let bytes = s.as_bytes();
    let mut pending_off: Option<(i64, i64)> = None;
    while i < bytes.len() {
        if bytes[i] == b'<' {
            let rest = &s[i..];
            if let Some(after) = rest.strip_prefix("<a:off ") {
                if let Some((x, y)) = parse_x_y(after) {
                    pending_off = Some((x, y));
                }
                i += 1;
                continue;
            }
            if let Some(after) = rest.strip_prefix("<a:ext ") {
                if let Some((cx, cy)) = parse_cx_cy(after) {
                    if let Some((ox, oy)) = pending_off.take() {
                        if ox < min_x { min_x = ox; }
                        if oy < min_y { min_y = oy; }
                        if ox + cx > max_x { max_x = ox + cx; }
                        if oy + cy > max_y { max_y = oy + cy; }
                    }
                }
                i += 1;
                continue;
            }
        }
        i += 1;
    }
    if min_x == i64::MAX {
        None
    } else {
        Some((min_x, min_y, max_x - min_x, max_y - min_y))
    }
}

fn parse_x_y(s: &str) -> Option<(i64, i64)> {
    let x = extract_attr(s, "x")?;
    let y = extract_attr(s, "y")?;
    Some((x, y))
}

fn parse_cx_cy(s: &str) -> Option<(i64, i64)> {
    let cx = extract_attr(s, "cx")?;
    let cy = extract_attr(s, "cy")?;
    Some((cx, cy))
}

fn extract_attr(s: &str, name: &str) -> Option<i64> {
    let needle = format!("{}=\"", name);
    let start = s.find(&needle)? + needle.len();
    let rest = &s[start..];
    let end = rest.find('"')?;
    rest[..end].parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse::parse_svg;

    fn convert(xml: &str) -> String {
        let doc = parse_svg(xml).unwrap();
        let ctx = EmitContext::from_options(&doc, &ConvertOptions::default());
        emit_document(&doc, &ctx)
    }

    #[test]
    fn emit_rect_preset_geom() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">
            <rect x="10" y="20" width="30" height="40" fill="#ff0000"/>
        </svg>"##);
        assert!(out.contains("<p:sp>"));
        assert!(out.contains("a:prstGeom prst=\"rect\""));
        assert!(out.contains("srgbClr val=\"FF0000\""));
    }

    #[test]
    fn emit_ellipse_preset() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">
            <circle cx="50" cy="50" r="20" fill="#00ff00"/>
        </svg>"##);
        assert!(out.contains("a:prstGeom prst=\"ellipse\""));
    }

    #[test]
    fn emit_path_custom_geom() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">
            <path d="M0 0 L10 0 L10 10 Z" fill="blue"/>
        </svg>"##);
        assert!(out.contains("a:custGeom"));
        assert!(out.contains("a:moveTo"));
        assert!(out.contains("a:lnTo"));
        assert!(out.contains("a:close"));
    }

    #[test]
    fn emit_stroke_only_line() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">
            <line x1="0" y1="0" x2="10" y2="10" stroke="#000000" stroke-width="2"/>
        </svg>"##);
        assert!(out.contains("a:ln"));
        assert!(out.contains("srgbClr val=\"000000\""));
    }

    #[test]
    fn emit_multiple_shapes_group_wraps() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">
            <g>
                <rect x="0" y="0" width="10" height="10"/>
                <rect x="20" y="20" width="10" height="10"/>
            </g>
        </svg>"##);
        assert!(out.contains("<p:grpSp>"));
        assert_eq!(out.matches("<p:sp>").count(), 2);
    }

    #[test]
    fn emit_image_data_uri_round_trip() {
        // 1×1 transparent PNG (smallest valid PNG we can embed)
        const PNG_BASE64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII=";
        let svg = format!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="100" height="100">
                <image x="10" y="20" width="50" height="40" xlink:href="data:image/png;base64,{PNG_BASE64}"/>
            </svg>"##
        );
        let doc = crate::parse::parse_svg(&svg).unwrap();
        let ctx = EmitContext::from_options(&doc, &ConvertOptions::default());
        let xml = emit_document(&doc, &ctx);
        let images = ctx.images.into_inner();

        assert!(xml.contains("<p:pic>"));
        assert!(xml.contains("r:embed=\"__mpl_office_img_0__\""));
        assert_eq!(images.len(), 1, "expected one image to be extracted");
        assert_eq!(images[0].sentinel, "__mpl_office_img_0__");
        assert_eq!(images[0].format, "png");
        assert!(!images[0].bytes.is_empty());
        // Sanity-check the PNG magic number
        assert_eq!(&images[0].bytes[..8], b"\x89PNG\r\n\x1a\n");
    }

    #[test]
    fn emit_image_negative_y_scale_sets_flipv() {
        // imshow typically emits a transform that flips Y. Make sure flipV
        // ends up on the <a:xfrm>.
        const PNG_BASE64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII=";
        let svg = format!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="100" height="100">
                <g transform="matrix(1 0 0 -1 0 100)">
                    <image x="0" y="0" width="100" height="100" xlink:href="data:image/png;base64,{PNG_BASE64}"/>
                </g>
            </svg>"##
        );
        let doc = crate::parse::parse_svg(&svg).unwrap();
        let ctx = EmitContext::from_options(&doc, &ConvertOptions::default());
        let xml = emit_document(&doc, &ctx);
        assert!(xml.contains("flipV=\"1\""), "expected flipV on inverted-Y image");
    }

    #[test]
    fn emit_honours_target_bbox() {
        let out = convert(r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
            <rect x="0" y="0" width="100" height="100"/>
        </svg>"##);
        // Without target, 100 user units → 100 * 9525 = 952500 EMU
        assert!(out.contains("cx=\"952500\""));

        let doc = parse_svg(r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
            <rect x="0" y="0" width="100" height="100"/>
        </svg>"##).unwrap();
        let opts = ConvertOptions {
            target_width_emu: Some(6_858_000),
            target_height_emu: Some(4_572_000),
            ..Default::default()
        };
        let ctx = EmitContext::from_options(&doc, &opts);
        let out2 = emit_document(&doc, &ctx);
        assert!(out2.contains("cx=\"6858000\""));
    }
}
