//! SVG path parsing and normalization.
//!
//! The pipeline:
//!
//! ```text
//! parse_d(&str)          → Vec<RawCmd>
//! to_absolute(...)       → Vec<RawCmd>   (only M/L/C/Q/S/T/A/Z)
//! normalize(...)         → Vec<PathCmd>  (only M/L/C/Z)
//! ```
//!
//! After normalization every curve is a cubic Bézier, which maps 1:1 to
//! DrawingML's `<a:cubicBezTo>`.

use crate::error::{Error, Result};

/// A single normalized path command. After `normalize()` we only emit these
/// four variants — the rest are converted into them.
#[derive(Debug, Clone, PartialEq)]
pub enum PathCmd {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    Close,
}

/// Raw commands as they appear in the SVG `d` attribute, absolutized but
/// not yet normalized.
#[derive(Debug, Clone, PartialEq)]
pub struct RawCmd {
    cmd: char,
    args: Vec<f64>,
}

/// Parse an SVG `d` attribute into raw commands (absolute only).
pub fn parse_d(d: &str) -> Result<Vec<RawCmd>> {
    let tokens = tokenize(d);
    let raw = collect_commands(tokens)?;
    Ok(to_absolute(&raw))
}

/// Parse → absolutize → normalize. Returns the cubic-only command list.
pub fn parse_and_normalize(d: &str) -> Result<Vec<PathCmd>> {
    let raw = parse_d(d)?;
    Ok(normalize(&raw))
}

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Cmd(char),
    Num(f64),
}

fn tokenize(d: &str) -> Vec<Token> {
    // Single-pass scanner. Numbers may be negative, decimal, or scientific;
    // they may be followed by a sign to start the next number with no
    // whitespace (e.g. "10-5" == [10, -5]).
    let bytes = d.as_bytes();
    let mut out = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b.is_ascii_whitespace() || b == b',' {
            i += 1;
            continue;
        }
        if b.is_ascii_alphabetic() {
            out.push(Token::Cmd(b as char));
            i += 1;
            continue;
        }
        // number
        let start = i;
        let mut saw_dot = false;
        let mut saw_exp = false;
        if b == b'+' || b == b'-' {
            i += 1;
        }
        while i < bytes.len() {
            let c = bytes[i];
            if c.is_ascii_digit() {
                i += 1;
            } else if c == b'.' && !saw_dot && !saw_exp {
                saw_dot = true;
                i += 1;
            } else if (c == b'e' || c == b'E') && !saw_exp {
                saw_exp = true;
                i += 1;
                if i < bytes.len() && (bytes[i] == b'+' || bytes[i] == b'-') {
                    i += 1;
                }
            } else {
                break;
            }
        }
        if start == i {
            // Unknown char — skip it.
            i += 1;
            continue;
        }
        if let Ok(v) = d[start..i].parse::<f64>() {
            out.push(Token::Num(v));
        }
    }
    out
}

fn arg_count(c: char) -> Option<usize> {
    Some(match c {
        'M' | 'm' | 'L' | 'l' | 'T' | 't' => 2,
        'H' | 'h' | 'V' | 'v' => 1,
        'C' | 'c' => 6,
        'S' | 's' | 'Q' | 'q' => 4,
        'A' | 'a' => 7,
        'Z' | 'z' => 0,
        _ => return None,
    })
}

fn collect_commands(tokens: Vec<Token>) -> Result<Vec<RawCmd>> {
    let mut out = Vec::new();
    let mut i = 0;
    let mut current: Option<char> = None;
    let mut pending: Vec<f64> = Vec::new();

    // Helper to flush pending args under the current command.
    // M/m chains implicitly become L/l after the first pair.
    fn flush(
        out: &mut Vec<RawCmd>,
        current: &mut Option<char>,
        pending: &mut Vec<f64>,
    ) -> Result<()> {
        let Some(cmd) = *current else {
            pending.clear();
            return Ok(());
        };
        let n =
            arg_count(cmd).ok_or_else(|| Error::Path(format!("unknown path command '{}'", cmd)))?;
        if n == 0 {
            out.push(RawCmd {
                cmd,
                args: Vec::new(),
            });
            pending.clear();
            // Clear current so a subsequent end-of-input flush doesn't re-emit.
            *current = None;
            return Ok(());
        }
        if pending.len() < n {
            if pending.is_empty() {
                return Ok(());
            }
            return Err(Error::Path(format!(
                "command '{}' expected {} args, got {}",
                cmd,
                n,
                pending.len()
            )));
        }
        let mut idx = 0;
        let mut cur = cmd;
        while idx + n <= pending.len() {
            out.push(RawCmd {
                cmd: cur,
                args: pending[idx..idx + n].to_vec(),
            });
            idx += n;
            // M → L / m → l for implicit repeats
            if cur == 'M' {
                cur = 'L';
            } else if cur == 'm' {
                cur = 'l';
            }
        }
        if idx != pending.len() {
            return Err(Error::Path(format!(
                "command '{}' has leftover args: {}",
                cmd,
                pending.len() - idx
            )));
        }
        *current = Some(cur);
        pending.clear();
        Ok(())
    }

    while i < tokens.len() {
        match &tokens[i] {
            Token::Cmd(c) => {
                flush(&mut out, &mut current, &mut pending)?;
                current = Some(*c);
                if arg_count(*c) == Some(0) {
                    flush(&mut out, &mut current, &mut pending)?;
                }
            }
            Token::Num(v) => {
                pending.push(*v);
            }
        }
        i += 1;
    }
    flush(&mut out, &mut current, &mut pending)?;
    Ok(out)
}

fn to_absolute(raw: &[RawCmd]) -> Vec<RawCmd> {
    let mut out = Vec::with_capacity(raw.len());
    let mut cx = 0.0f64;
    let mut cy = 0.0f64;
    let mut sx = 0.0f64;
    let mut sy = 0.0f64;

    for r in raw {
        let a = &r.args;
        match r.cmd {
            'M' => {
                cx = a[0];
                cy = a[1];
                sx = cx;
                sy = cy;
                out.push(RawCmd {
                    cmd: 'M',
                    args: vec![cx, cy],
                });
            }
            'm' => {
                cx += a[0];
                cy += a[1];
                sx = cx;
                sy = cy;
                out.push(RawCmd {
                    cmd: 'M',
                    args: vec![cx, cy],
                });
            }
            'L' => {
                cx = a[0];
                cy = a[1];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'l' => {
                cx += a[0];
                cy += a[1];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'H' => {
                cx = a[0];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'h' => {
                cx += a[0];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'V' => {
                cy = a[0];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'v' => {
                cy += a[0];
                out.push(RawCmd {
                    cmd: 'L',
                    args: vec![cx, cy],
                });
            }
            'C' => {
                out.push(RawCmd {
                    cmd: 'C',
                    args: a.clone(),
                });
                cx = a[4];
                cy = a[5];
            }
            'c' => {
                let abs = vec![
                    cx + a[0],
                    cy + a[1],
                    cx + a[2],
                    cy + a[3],
                    cx + a[4],
                    cy + a[5],
                ];
                cx = abs[4];
                cy = abs[5];
                out.push(RawCmd {
                    cmd: 'C',
                    args: abs,
                });
            }
            'S' => {
                out.push(RawCmd {
                    cmd: 'S',
                    args: a.clone(),
                });
                cx = a[2];
                cy = a[3];
            }
            's' => {
                let abs = vec![cx + a[0], cy + a[1], cx + a[2], cy + a[3]];
                cx = abs[2];
                cy = abs[3];
                out.push(RawCmd {
                    cmd: 'S',
                    args: abs,
                });
            }
            'Q' => {
                out.push(RawCmd {
                    cmd: 'Q',
                    args: a.clone(),
                });
                cx = a[2];
                cy = a[3];
            }
            'q' => {
                let abs = vec![cx + a[0], cy + a[1], cx + a[2], cy + a[3]];
                cx = abs[2];
                cy = abs[3];
                out.push(RawCmd {
                    cmd: 'Q',
                    args: abs,
                });
            }
            'T' => {
                out.push(RawCmd {
                    cmd: 'T',
                    args: a.clone(),
                });
                cx = a[0];
                cy = a[1];
            }
            't' => {
                let abs = vec![cx + a[0], cy + a[1]];
                cx = abs[0];
                cy = abs[1];
                out.push(RawCmd {
                    cmd: 'T',
                    args: abs,
                });
            }
            'A' => {
                out.push(RawCmd {
                    cmd: 'A',
                    args: a.clone(),
                });
                cx = a[5];
                cy = a[6];
            }
            'a' => {
                let abs = vec![a[0], a[1], a[2], a[3], a[4], cx + a[5], cy + a[6]];
                cx = abs[5];
                cy = abs[6];
                out.push(RawCmd {
                    cmd: 'A',
                    args: abs,
                });
            }
            'Z' | 'z' => {
                out.push(RawCmd {
                    cmd: 'Z',
                    args: Vec::new(),
                });
                cx = sx;
                cy = sy;
            }
            _ => {}
        }
    }
    out
}

#[inline]
fn reflect_cp(cpx: f64, cpy: f64, cx: f64, cy: f64) -> (f64, f64) {
    (2.0 * cx - cpx, 2.0 * cy - cpy)
}

#[inline]
fn quad_to_cubic(qx: f64, qy: f64, p0x: f64, p0y: f64, p3x: f64, p3y: f64) -> [f64; 6] {
    let cp1x = p0x + 2.0 / 3.0 * (qx - p0x);
    let cp1y = p0y + 2.0 / 3.0 * (qy - p0y);
    let cp2x = p3x + 2.0 / 3.0 * (qx - p3x);
    let cp2y = p3y + 2.0 / 3.0 * (qy - p3y);
    [cp1x, cp1y, cp2x, cp2y, p3x, p3y]
}

/// Convert a single SVG elliptical arc (endpoint parameterization, already
/// in absolute coordinates) into a sequence of cubic Bézier segments.
///
/// Matches the algorithm in SVG 1.1 F.6 (elliptical arc → center param) and
/// the standard 4/3·tan(θ/4) control-point rule for the unit circle.
#[allow(clippy::too_many_arguments)]
fn arc_to_cubics(
    x1: f64,
    y1: f64,
    mut rx: f64,
    mut ry: f64,
    phi_deg: f64,
    large_arc: bool,
    sweep: bool,
    x2: f64,
    y2: f64,
) -> Vec<[f64; 6]> {
    if (x1 - x2).abs() < 1e-12 && (y1 - y2).abs() < 1e-12 {
        return Vec::new();
    }
    rx = rx.abs();
    ry = ry.abs();
    if rx < 1e-12 || ry < 1e-12 {
        // Degenerate: emit a straight line as a single cubic.
        return vec![[x1, y1, x2, y2, x2, y2]];
    }

    let phi = phi_deg.to_radians();
    let (sin_phi, cos_phi) = phi.sin_cos();

    // Step 1: compute (x1', y1')
    let dx = (x1 - x2) / 2.0;
    let dy = (y1 - y2) / 2.0;
    let x1p = cos_phi * dx + sin_phi * dy;
    let y1p = -sin_phi * dx + cos_phi * dy;

    // Step 2: enforce radius constraint
    let mut rx2 = rx * rx;
    let mut ry2 = ry * ry;
    let x1p2 = x1p * x1p;
    let y1p2 = y1p * y1p;

    let lam = x1p2 / rx2 + y1p2 / ry2;
    if lam > 1.0 {
        let lam_sqrt = lam.sqrt();
        rx *= lam_sqrt;
        ry *= lam_sqrt;
        rx2 = rx * rx;
        ry2 = ry * ry;
    }

    let num = (rx2 * ry2 - rx2 * y1p2 - ry2 * x1p2).max(0.0);
    let den = rx2 * y1p2 + ry2 * x1p2;
    let mut sq = if den > 1e-20 { (num / den).sqrt() } else { 0.0 };
    if large_arc == sweep {
        sq = -sq;
    }
    let cxp = sq * rx * y1p / ry;
    let cyp = -sq * ry * x1p / rx;

    // Step 3: back to original coordinate system
    let arc_cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2.0;
    let arc_cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2.0;

    // Step 4: theta1, dtheta
    fn angle(ux: f64, uy: f64, vx: f64, vy: f64) -> f64 {
        let n = ((ux * ux + uy * uy) * (vx * vx + vy * vy)).sqrt();
        if n < 1e-20 {
            return 0.0;
        }
        let c = ((ux * vx + uy * vy) / n).clamp(-1.0, 1.0);
        let a = c.acos();
        if ux * vy - uy * vx < 0.0 {
            -a
        } else {
            a
        }
    }

    let ux1 = (x1p - cxp) / rx;
    let uy1 = (y1p - cyp) / ry;
    let ux2 = (-x1p - cxp) / rx;
    let uy2 = (-y1p - cyp) / ry;

    let theta1 = angle(1.0, 0.0, ux1, uy1);
    let mut dtheta = angle(ux1, uy1, ux2, uy2);

    if !sweep && dtheta > 0.0 {
        dtheta -= 2.0 * std::f64::consts::PI;
    } else if sweep && dtheta < 0.0 {
        dtheta += 2.0 * std::f64::consts::PI;
    }

    // Split into segments of ≤ 90°
    let n_segs = (dtheta.abs() / (std::f64::consts::PI / 2.0))
        .ceil()
        .max(1.0) as usize;
    let d_per_seg = dtheta / n_segs as f64;
    let alpha = 4.0 / 3.0 * (d_per_seg / 4.0).tan();

    let mut out = Vec::with_capacity(n_segs);
    for i in 0..n_segs {
        let t1 = theta1 + i as f64 * d_per_seg;
        let t2 = theta1 + (i + 1) as f64 * d_per_seg;
        let (sin_t1, cos_t1) = t1.sin_cos();
        let (sin_t2, cos_t2) = t2.sin_cos();

        let ep1x = cos_t1 - alpha * sin_t1;
        let ep1y = sin_t1 + alpha * cos_t1;
        let ep2x = cos_t2 + alpha * sin_t2;
        let ep2y = sin_t2 - alpha * cos_t2;
        let epx = cos_t2;
        let epy = sin_t2;

        let transform_pt = |px: f64, py: f64| -> (f64, f64) {
            let x = rx * px;
            let y = ry * py;
            let xr = cos_phi * x - sin_phi * y + arc_cx;
            let yr = sin_phi * x + cos_phi * y + arc_cy;
            (xr, yr)
        };
        let (c1x, c1y) = transform_pt(ep1x, ep1y);
        let (c2x, c2y) = transform_pt(ep2x, ep2y);
        let (ex, ey) = transform_pt(epx, epy);
        out.push([c1x, c1y, c2x, c2y, ex, ey]);
    }
    out
}

/// Normalize absolute path commands into M/L/C/Z only.
pub fn normalize(absolute: &[RawCmd]) -> Vec<PathCmd> {
    let mut out = Vec::with_capacity(absolute.len());
    let mut cx = 0.0f64;
    let mut cy = 0.0f64;
    let mut last_cp_x = 0.0f64; // last cubic control for S
    let mut last_cp_y = 0.0f64;
    let mut last_qp_x = 0.0f64; // last quadratic control for T
    let mut last_qp_y = 0.0f64;
    let mut last_cmd = ' ';

    for cmd in absolute {
        let a = &cmd.args;
        match cmd.cmd {
            'M' => {
                cx = a[0];
                cy = a[1];
                last_cp_x = cx;
                last_cp_y = cy;
                last_qp_x = cx;
                last_qp_y = cy;
                out.push(PathCmd::MoveTo { x: cx, y: cy });
            }
            'L' => {
                cx = a[0];
                cy = a[1];
                last_cp_x = cx;
                last_cp_y = cy;
                last_qp_x = cx;
                last_qp_y = cy;
                out.push(PathCmd::LineTo { x: cx, y: cy });
            }
            'C' => {
                last_cp_x = a[2];
                last_cp_y = a[3];
                cx = a[4];
                cy = a[5];
                out.push(PathCmd::CubicTo {
                    x1: a[0],
                    y1: a[1],
                    x2: a[2],
                    y2: a[3],
                    x: a[4],
                    y: a[5],
                });
                last_qp_x = cx;
                last_qp_y = cy;
            }
            'S' => {
                let (rcp_x, rcp_y) = if last_cmd == 'C' || last_cmd == 'S' {
                    reflect_cp(last_cp_x, last_cp_y, cx, cy)
                } else {
                    (cx, cy)
                };
                last_cp_x = a[0];
                last_cp_y = a[1];
                let nx = a[2];
                let ny = a[3];
                out.push(PathCmd::CubicTo {
                    x1: rcp_x,
                    y1: rcp_y,
                    x2: a[0],
                    y2: a[1],
                    x: nx,
                    y: ny,
                });
                cx = nx;
                cy = ny;
                last_qp_x = cx;
                last_qp_y = cy;
            }
            'Q' => {
                let cubic = quad_to_cubic(a[0], a[1], cx, cy, a[2], a[3]);
                last_qp_x = a[0];
                last_qp_y = a[1];
                out.push(PathCmd::CubicTo {
                    x1: cubic[0],
                    y1: cubic[1],
                    x2: cubic[2],
                    y2: cubic[3],
                    x: cubic[4],
                    y: cubic[5],
                });
                cx = a[2];
                cy = a[3];
                last_cp_x = cx;
                last_cp_y = cy;
            }
            'T' => {
                let (qx, qy) = if last_cmd == 'Q' || last_cmd == 'T' {
                    reflect_cp(last_qp_x, last_qp_y, cx, cy)
                } else {
                    (cx, cy)
                };
                last_qp_x = qx;
                last_qp_y = qy;
                let cubic = quad_to_cubic(qx, qy, cx, cy, a[0], a[1]);
                out.push(PathCmd::CubicTo {
                    x1: cubic[0],
                    y1: cubic[1],
                    x2: cubic[2],
                    y2: cubic[3],
                    x: cubic[4],
                    y: cubic[5],
                });
                cx = a[0];
                cy = a[1];
                last_cp_x = cx;
                last_cp_y = cy;
            }
            'A' => {
                let segs = arc_to_cubics(
                    cx,
                    cy,
                    a[0],
                    a[1],
                    a[2],
                    a[3] != 0.0,
                    a[4] != 0.0,
                    a[5],
                    a[6],
                );
                for s in segs {
                    out.push(PathCmd::CubicTo {
                        x1: s[0],
                        y1: s[1],
                        x2: s[2],
                        y2: s[3],
                        x: s[4],
                        y: s[5],
                    });
                }
                cx = a[5];
                cy = a[6];
                last_cp_x = cx;
                last_cp_y = cy;
                last_qp_x = cx;
                last_qp_y = cy;
            }
            'Z' => {
                out.push(PathCmd::Close);
            }
            _ => {}
        }
        last_cmd = cmd.cmd;
    }
    out
}

/// Apply an affine transform to every point in a normalized command list.
pub fn transform_cmds(cmds: &[PathCmd], t: &crate::transform::Affine) -> Vec<PathCmd> {
    cmds.iter()
        .map(|c| match c {
            PathCmd::MoveTo { x, y } => {
                let (nx, ny) = t.transform_point(*x, *y);
                PathCmd::MoveTo { x: nx, y: ny }
            }
            PathCmd::LineTo { x, y } => {
                let (nx, ny) = t.transform_point(*x, *y);
                PathCmd::LineTo { x: nx, y: ny }
            }
            PathCmd::CubicTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                let (a, b) = t.transform_point(*x1, *y1);
                let (c, d) = t.transform_point(*x2, *y2);
                let (e, f) = t.transform_point(*x, *y);
                PathCmd::CubicTo {
                    x1: a,
                    y1: b,
                    x2: c,
                    y2: d,
                    x: e,
                    y: f,
                }
            }
            PathCmd::Close => PathCmd::Close,
        })
        .collect()
}

/// Bounding box of a normalized command list (using only on-curve points
/// and control points — conservative but adequate for shape sizing).
pub fn bbox(cmds: &[PathCmd]) -> (f64, f64, f64, f64) {
    let mut minx = f64::INFINITY;
    let mut miny = f64::INFINITY;
    let mut maxx = f64::NEG_INFINITY;
    let mut maxy = f64::NEG_INFINITY;
    let push = |x: f64, y: f64, minx: &mut f64, miny: &mut f64, maxx: &mut f64, maxy: &mut f64| {
        if x < *minx {
            *minx = x;
        }
        if y < *miny {
            *miny = y;
        }
        if x > *maxx {
            *maxx = x;
        }
        if y > *maxy {
            *maxy = y;
        }
    };
    for c in cmds {
        match c {
            PathCmd::MoveTo { x, y } | PathCmd::LineTo { x, y } => {
                push(*x, *y, &mut minx, &mut miny, &mut maxx, &mut maxy);
            }
            PathCmd::CubicTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                push(*x1, *y1, &mut minx, &mut miny, &mut maxx, &mut maxy);
                push(*x2, *y2, &mut minx, &mut miny, &mut maxx, &mut maxy);
                push(*x, *y, &mut minx, &mut miny, &mut maxx, &mut maxy);
            }
            PathCmd::Close => {}
        }
    }
    if !minx.is_finite() {
        return (0.0, 0.0, 0.0, 0.0);
    }
    (minx, miny, maxx - minx, maxy - miny)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn approx(a: f64, b: f64) -> bool {
        (a - b).abs() < 1e-6
    }

    #[test]
    fn tokenize_mixed() {
        let toks = tokenize("M10 20L30,40");
        assert_eq!(
            toks,
            vec![
                Token::Cmd('M'),
                Token::Num(10.0),
                Token::Num(20.0),
                Token::Cmd('L'),
                Token::Num(30.0),
                Token::Num(40.0),
            ]
        );
    }

    #[test]
    fn tokenize_no_space_negatives() {
        let toks = tokenize("M10-5-3.5");
        assert_eq!(
            toks,
            vec![
                Token::Cmd('M'),
                Token::Num(10.0),
                Token::Num(-5.0),
                Token::Num(-3.5),
            ]
        );
    }

    #[test]
    fn parse_simple_mlz() {
        let cmds = parse_and_normalize("M0 0 L10 0 L10 10 Z").unwrap();
        assert_eq!(cmds.len(), 4);
        assert!(matches!(cmds[0], PathCmd::MoveTo { x: 0.0, y: 0.0 }));
        assert!(matches!(cmds[3], PathCmd::Close));
    }

    #[test]
    fn implicit_lineto_after_moveto() {
        // "M0 0 10 0 20 0" → M, L, L
        let cmds = parse_and_normalize("M0 0 10 0 20 0").unwrap();
        assert_eq!(cmds.len(), 3);
        match cmds[1] {
            PathCmd::LineTo { x, y } => {
                assert_eq!(x, 10.0);
                assert_eq!(y, 0.0);
            }
            _ => panic!("expected LineTo"),
        }
    }

    #[test]
    fn relative_moveto_lineto() {
        let cmds = parse_and_normalize("m10 10 l5 5").unwrap();
        match cmds[0] {
            PathCmd::MoveTo { x, y } => {
                assert_eq!(x, 10.0);
                assert_eq!(y, 10.0);
            }
            _ => panic!(),
        }
        match cmds[1] {
            PathCmd::LineTo { x, y } => {
                assert_eq!(x, 15.0);
                assert_eq!(y, 15.0);
            }
            _ => panic!(),
        }
    }

    #[test]
    fn horizontal_vertical() {
        let cmds = parse_and_normalize("M0 0 H10 V10 Z").unwrap();
        // M, L, L, Z
        match cmds[1] {
            PathCmd::LineTo { x, y } => {
                assert_eq!(x, 10.0);
                assert_eq!(y, 0.0);
            }
            _ => panic!(),
        }
        match cmds[2] {
            PathCmd::LineTo { x, y } => {
                assert_eq!(x, 10.0);
                assert_eq!(y, 10.0);
            }
            _ => panic!(),
        }
    }

    #[test]
    fn cubic_passthrough() {
        let cmds = parse_and_normalize("M0 0 C 10 0 20 10 30 10").unwrap();
        match cmds[1] {
            PathCmd::CubicTo { x, y, .. } => {
                assert_eq!(x, 30.0);
                assert_eq!(y, 10.0);
            }
            _ => panic!(),
        }
    }

    #[test]
    fn quadratic_to_cubic() {
        let cmds = parse_and_normalize("M0 0 Q10 0 20 10").unwrap();
        assert_eq!(cmds.len(), 2);
        match cmds[1] {
            PathCmd::CubicTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                // Q(0,0)→(10,0)→(20,10). Cubic control points = 2/3 of the way.
                assert!(approx(x1, 0.0 + 2.0 / 3.0 * 10.0));
                assert!(approx(y1, 0.0));
                assert!(approx(x2, 20.0 + 2.0 / 3.0 * (10.0 - 20.0)));
                assert!(approx(y2, 10.0 + 2.0 / 3.0 * (0.0 - 10.0)));
                assert_eq!(x, 20.0);
                assert_eq!(y, 10.0);
            }
            _ => panic!(),
        }
    }

    #[test]
    fn smooth_cubic_reflection() {
        // M0 0 C 0 10 10 10 10 0 S 20 -10 20 0
        // After first C, last_cp=(10,10), cur=(10,0). Reflect → (10, -10).
        let cmds = parse_and_normalize("M0 0 C 0 10 10 10 10 0 S 20 -10 20 0").unwrap();
        assert_eq!(cmds.len(), 3);
        if let PathCmd::CubicTo { x1, y1, .. } = cmds[2] {
            assert!(approx(x1, 10.0));
            assert!(approx(y1, -10.0));
        } else {
            panic!();
        }
    }

    #[test]
    fn arc_quarter_circle() {
        // Quarter circle: from (10, 0) arc to (0, 10), rx=ry=10, sweep=0 (CCW in y-down? — just check endpoints)
        let cmds = parse_and_normalize("M10 0 A10 10 0 0 0 0 10").unwrap();
        let last = cmds.last().unwrap();
        if let PathCmd::CubicTo { x, y, .. } = last {
            assert!(approx(*x, 0.0));
            assert!(approx(*y, 10.0));
        } else {
            panic!();
        }
    }

    #[test]
    fn arc_full_semicircle_splits_into_two() {
        // Semicircle from (0,0) to (20,0), r=10. >90°, should split.
        let cmds = parse_and_normalize("M0 0 A10 10 0 0 1 20 0").unwrap();
        // 1 MoveTo + at least 2 CubicTo (180° → 2 segments)
        let cubics = cmds
            .iter()
            .filter(|c| matches!(c, PathCmd::CubicTo { .. }))
            .count();
        assert!(cubics >= 2);
    }

    #[test]
    fn bbox_simple() {
        let cmds = parse_and_normalize("M0 0 L10 0 L10 5 Z").unwrap();
        let (x, y, w, h) = bbox(&cmds);
        assert_eq!(x, 0.0);
        assert_eq!(y, 0.0);
        assert_eq!(w, 10.0);
        assert_eq!(h, 5.0);
    }

    #[test]
    fn transform_cmds_translate() {
        let cmds = parse_and_normalize("M0 0 L10 0").unwrap();
        let t = crate::transform::Affine::translate(5.0, 5.0);
        let out = transform_cmds(&cmds, &t);
        if let PathCmd::MoveTo { x, y } = out[0] {
            assert_eq!(x, 5.0);
            assert_eq!(y, 5.0);
        } else {
            panic!();
        }
    }
}
