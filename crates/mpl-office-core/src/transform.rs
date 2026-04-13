//! SVG transform parser and 2D affine matrix.
//!
//! Supports: `translate(tx [,ty])`, `scale(sx [,sy])`, `rotate(a [,cx,cy])`,
//! `skewX(a)`, `skewY(a)`, `matrix(a,b,c,d,e,f)`. Multiple transforms
//! compose left-to-right (the SVG 1.1 convention).

/// 2D affine transform, row-major:
///
/// ```text
/// | a c e |
/// | b d f |
/// | 0 0 1 |
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Affine {
    pub a: f64,
    pub b: f64,
    pub c: f64,
    pub d: f64,
    pub e: f64,
    pub f: f64,
}

impl Affine {
    pub const IDENTITY: Affine = Affine {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: 0.0,
        f: 0.0,
    };

    pub fn translate(tx: f64, ty: f64) -> Self {
        Affine {
            a: 1.0,
            b: 0.0,
            c: 0.0,
            d: 1.0,
            e: tx,
            f: ty,
        }
    }

    pub fn scale(sx: f64, sy: f64) -> Self {
        Affine {
            a: sx,
            b: 0.0,
            c: 0.0,
            d: sy,
            e: 0.0,
            f: 0.0,
        }
    }

    /// Rotate by `deg` degrees around the origin.
    pub fn rotate_deg(deg: f64) -> Self {
        let r = deg.to_radians();
        let (s, c) = r.sin_cos();
        Affine {
            a: c,
            b: s,
            c: -s,
            d: c,
            e: 0.0,
            f: 0.0,
        }
    }

    pub fn skew_x_deg(deg: f64) -> Self {
        Affine {
            a: 1.0,
            b: 0.0,
            c: deg.to_radians().tan(),
            d: 1.0,
            e: 0.0,
            f: 0.0,
        }
    }

    pub fn skew_y_deg(deg: f64) -> Self {
        Affine {
            a: 1.0,
            b: deg.to_radians().tan(),
            c: 0.0,
            d: 1.0,
            e: 0.0,
            f: 0.0,
        }
    }

    /// SVG transform-list compose: `self * other` in matrix math.
    ///
    /// The SVG spec says a list `"A B"` behaves as if `B` is the inner
    /// transform — a point `p` becomes `A * B * p`. Parsing left-to-right
    /// we start from the identity and multiply each new step on the right,
    /// so `result = running.then(step)` must yield `running * step`.
    pub fn then(self, other: Affine) -> Affine {
        let a = self.a * other.a + self.c * other.b;
        let b = self.b * other.a + self.d * other.b;
        let c = self.a * other.c + self.c * other.d;
        let d = self.b * other.c + self.d * other.d;
        let e = self.a * other.e + self.c * other.f + self.e;
        let f = self.b * other.e + self.d * other.f + self.f;
        Affine { a, b, c, d, e, f }
    }

    pub fn transform_point(&self, x: f64, y: f64) -> (f64, f64) {
        (self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f)
    }

    pub fn is_identity(&self) -> bool {
        self.a == 1.0
            && self.b == 0.0
            && self.c == 0.0
            && self.d == 1.0
            && self.e == 0.0
            && self.f == 0.0
    }

    /// True if this affine is translate+uniform-scale only (no rotation,
    /// no skew). Useful when emitters prefer `prstGeom` over custom paths.
    pub fn is_axis_aligned_scale(&self) -> bool {
        self.b == 0.0 && self.c == 0.0
    }
}

impl Default for Affine {
    fn default() -> Self {
        Affine::IDENTITY
    }
}

/// Parse an SVG `transform` attribute value.
pub fn parse_transform(s: &str) -> Affine {
    let mut result = Affine::IDENTITY;
    let s = s.trim();
    if s.is_empty() {
        return result;
    }

    let bytes = s.as_bytes();
    let mut i = 0;

    while i < bytes.len() {
        // Skip whitespace and commas
        while i < bytes.len() && (bytes[i].is_ascii_whitespace() || bytes[i] == b',') {
            i += 1;
        }
        if i >= bytes.len() {
            break;
        }

        // Read function name
        let name_start = i;
        while i < bytes.len() && (bytes[i].is_ascii_alphabetic()) {
            i += 1;
        }
        if name_start == i {
            break;
        }
        let name = &s[name_start..i];

        // Skip whitespace before '('
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        if i >= bytes.len() || bytes[i] != b'(' {
            break;
        }
        i += 1; // skip '('

        // Find matching ')'
        let arg_start = i;
        while i < bytes.len() && bytes[i] != b')' {
            i += 1;
        }
        if i >= bytes.len() {
            break;
        }
        let arg_str = &s[arg_start..i];
        i += 1; // skip ')'

        let args: Vec<f64> = arg_str
            .split(|c: char| c.is_whitespace() || c == ',')
            .filter(|t| !t.is_empty())
            .filter_map(|t| t.parse::<f64>().ok())
            .collect();

        let step = match name {
            "translate" => match args.as_slice() {
                [tx] => Affine::translate(*tx, 0.0),
                [tx, ty, ..] => Affine::translate(*tx, *ty),
                _ => Affine::IDENTITY,
            },
            "scale" => match args.as_slice() {
                [s] => Affine::scale(*s, *s),
                [sx, sy, ..] => Affine::scale(*sx, *sy),
                _ => Affine::IDENTITY,
            },
            "rotate" => match args.as_slice() {
                [deg] => Affine::rotate_deg(*deg),
                // rotate(a, cx, cy) ≡ translate(cx,cy) rotate(a) translate(-cx,-cy)
                [deg, cx, cy, ..] => Affine::translate(*cx, *cy)
                    .then(Affine::rotate_deg(*deg))
                    .then(Affine::translate(-cx, -cy)),
                _ => Affine::IDENTITY,
            },
            "skewX" => args.first().copied().map(Affine::skew_x_deg).unwrap_or(Affine::IDENTITY),
            "skewY" => args.first().copied().map(Affine::skew_y_deg).unwrap_or(Affine::IDENTITY),
            "matrix" => {
                if args.len() >= 6 {
                    Affine {
                        a: args[0],
                        b: args[1],
                        c: args[2],
                        d: args[3],
                        e: args[4],
                        f: args[5],
                    }
                } else {
                    Affine::IDENTITY
                }
            }
            _ => Affine::IDENTITY,
        };
        result = result.then(step);
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    fn approx(a: f64, b: f64) -> bool {
        (a - b).abs() < 1e-9
    }

    #[test]
    fn identity() {
        let t = parse_transform("");
        assert!(t.is_identity());
    }

    #[test]
    fn translate_two_args() {
        let t = parse_transform("translate(10, 20)");
        assert_eq!(t.e, 10.0);
        assert_eq!(t.f, 20.0);
    }

    #[test]
    fn translate_one_arg() {
        let t = parse_transform("translate(5)");
        assert_eq!(t.e, 5.0);
        assert_eq!(t.f, 0.0);
    }

    #[test]
    fn scale_uniform() {
        let t = parse_transform("scale(2)");
        assert_eq!(t.a, 2.0);
        assert_eq!(t.d, 2.0);
    }

    #[test]
    fn rotate_90() {
        let t = parse_transform("rotate(90)");
        let (x, y) = t.transform_point(1.0, 0.0);
        assert!(approx(x, 0.0));
        assert!(approx(y, 1.0));
    }

    #[test]
    fn rotate_about_point() {
        let t = parse_transform("rotate(90, 10, 10)");
        // (10, 0) rotated 90deg about (10, 10) -> (20, 10)
        let (x, y) = t.transform_point(10.0, 0.0);
        assert!(approx(x, 20.0));
        assert!(approx(y, 10.0));
    }

    #[test]
    fn matrix_raw() {
        let t = parse_transform("matrix(1,0,0,1,50,60)");
        assert_eq!(t.e, 50.0);
        assert_eq!(t.f, 60.0);
    }

    #[test]
    fn composition_translate_then_scale() {
        // SVG says: first translate, then scale => applied in order,
        // so a point (0,0) goes translate->(10,10)->scale(2,2)->(20,20)
        let t = parse_transform("translate(10, 10) scale(2, 2)");
        let (x, y) = t.transform_point(0.0, 0.0);
        assert!(approx(x, 10.0));
        assert!(approx(y, 10.0));
        let (x, y) = t.transform_point(1.0, 1.0);
        assert!(approx(x, 12.0));
        assert!(approx(y, 12.0));
    }
}
