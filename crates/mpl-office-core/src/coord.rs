//! Coordinate system helpers.
//!
//! DrawingML uses EMU (English Metric Units): 914_400 EMU = 1 inch.
//! SVG uses "user units" which default to pixels, typically at 96 DPI.

/// 1 inch in EMU.
pub const EMU_PER_INCH: i64 = 914_400;

/// 1 CSS pixel in EMU (at the conventional 96 DPI of the DrawingML world).
pub const EMU_PER_PX: i64 = EMU_PER_INCH / 96; // 9_525

/// matplotlib's SVG backend writes at 72 DPI by default.
pub const MATPLOTLIB_SVG_DPI: f64 = 72.0;

/// Convert SVG pixels to EMU (rounded).
#[inline]
pub fn px_to_emu(px: f64) -> i64 {
    (px * EMU_PER_PX as f64).round() as i64
}

/// Convert pixels at an arbitrary DPI to EMU (rounded).
#[inline]
pub fn px_to_emu_at_dpi(px: f64, dpi: f64) -> i64 {
    (px * (EMU_PER_INCH as f64) / dpi).round() as i64
}

/// Convert inches to EMU.
#[inline]
pub fn inches_to_emu(inches: f64) -> i64 {
    (inches * EMU_PER_INCH as f64).round() as i64
}

/// DrawingML font size unit: 1/100 of a point. 1 px at 96 DPI = 0.75 pt.
pub const FONT_PX_TO_HUNDREDTHS_PT: f64 = 75.0;

/// DrawingML angle unit: 60_000ths of a degree.
pub const ANGLE_UNIT: i64 = 60_000;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn px_to_emu_one_inch() {
        assert_eq!(px_to_emu(96.0), EMU_PER_INCH);
    }

    #[test]
    fn px_to_emu_zero() {
        assert_eq!(px_to_emu(0.0), 0);
    }

    #[test]
    fn inches_to_emu_six_inches() {
        assert_eq!(inches_to_emu(6.0), 5_486_400);
    }

    #[test]
    fn px_to_emu_at_72dpi() {
        // 72 px @ 72 DPI = 1 inch = 914_400 EMU
        assert_eq!(px_to_emu_at_dpi(72.0, 72.0), EMU_PER_INCH);
    }
}
