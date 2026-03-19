//! Measurement unit types for OOXML
//!
//! OOXML uses various measurement units:
//! - Twips (1/20 of a point, 1440 per inch)
//! - Half-points (for font sizes, 2 half-points = 1 point)
//! - EMU (English Metric Units, 914400 per inch)
//! - Points, centimeters, millimeters, inches for convenience

/// Points (1/72 inch)
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Pt(pub f64);

/// EMU - English Metric Units (914400 per inch)
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Emu(pub i64);

/// Twips (1/20 point = 1/1440 inch)
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Twip(pub i32);

/// Centimeters
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Cm(pub f64);

/// Millimeters
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Mm(pub f64);

/// Inches
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct Inch(pub f64);

/// Half-points (used for font sizes in OOXML, 1 half-point = 0.5pt)
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct HalfPt(pub u32);

// === Constructors ===

impl Pt {
    pub fn new(value: f64) -> Self {
        Self(value)
    }
}

impl Emu {
    pub fn new(value: i64) -> Self {
        Self(value)
    }
}

impl Twip {
    pub fn new(value: i32) -> Self {
        Self(value)
    }
}

impl Cm {
    pub fn new(value: f64) -> Self {
        Self(value)
    }
}

impl Mm {
    pub fn new(value: f64) -> Self {
        Self(value)
    }
}

impl Inch {
    pub fn new(value: f64) -> Self {
        Self(value)
    }
}

impl HalfPt {
    pub fn new(value: u32) -> Self {
        Self(value)
    }
}

// === Constants ===
// 1 inch = 72 pt = 1440 twips = 914400 EMU = 2.54 cm = 25.4 mm

const TWIPS_PER_PT: f64 = 20.0;
const EMU_PER_PT: f64 = 12700.0;
const PT_PER_INCH: f64 = 72.0;
const CM_PER_INCH: f64 = 2.54;
const MM_PER_INCH: f64 = 25.4;

// === From Pt ===

impl From<Pt> for Twip {
    fn from(pt: Pt) -> Self {
        Twip((pt.0 * TWIPS_PER_PT) as i32)
    }
}

impl From<Pt> for Emu {
    fn from(pt: Pt) -> Self {
        Emu((pt.0 * EMU_PER_PT) as i64)
    }
}

impl From<Pt> for Inch {
    fn from(pt: Pt) -> Self {
        Inch(pt.0 / PT_PER_INCH)
    }
}

impl From<Pt> for Cm {
    fn from(pt: Pt) -> Self {
        Cm(pt.0 / PT_PER_INCH * CM_PER_INCH)
    }
}

impl From<Pt> for Mm {
    fn from(pt: Pt) -> Self {
        Mm(pt.0 / PT_PER_INCH * MM_PER_INCH)
    }
}

impl From<Pt> for HalfPt {
    fn from(pt: Pt) -> Self {
        HalfPt((pt.0 * 2.0) as u32)
    }
}

// === From Twip ===

impl From<Twip> for Pt {
    fn from(twip: Twip) -> Self {
        Pt(twip.0 as f64 / TWIPS_PER_PT)
    }
}

impl From<Twip> for Emu {
    fn from(twip: Twip) -> Self {
        Emu::from(Pt::from(twip))
    }
}

impl From<Twip> for Inch {
    fn from(twip: Twip) -> Self {
        Inch::from(Pt::from(twip))
    }
}

impl From<Twip> for Cm {
    fn from(twip: Twip) -> Self {
        Cm::from(Pt::from(twip))
    }
}

impl From<Twip> for Mm {
    fn from(twip: Twip) -> Self {
        Mm::from(Pt::from(twip))
    }
}

// === From Emu ===

impl From<Emu> for Pt {
    fn from(emu: Emu) -> Self {
        Pt(emu.0 as f64 / EMU_PER_PT)
    }
}

impl From<Emu> for Twip {
    fn from(emu: Emu) -> Self {
        Twip::from(Pt::from(emu))
    }
}

impl From<Emu> for Inch {
    fn from(emu: Emu) -> Self {
        Inch::from(Pt::from(emu))
    }
}

impl From<Emu> for Cm {
    fn from(emu: Emu) -> Self {
        Cm::from(Pt::from(emu))
    }
}

impl From<Emu> for Mm {
    fn from(emu: Emu) -> Self {
        Mm::from(Pt::from(emu))
    }
}

// === From Inch ===

impl From<Inch> for Pt {
    fn from(inch: Inch) -> Self {
        Pt(inch.0 * PT_PER_INCH)
    }
}

impl From<Inch> for Twip {
    fn from(inch: Inch) -> Self {
        Twip::from(Pt::from(inch))
    }
}

impl From<Inch> for Emu {
    fn from(inch: Inch) -> Self {
        Emu::from(Pt::from(inch))
    }
}

impl From<Inch> for Cm {
    fn from(inch: Inch) -> Self {
        Cm(inch.0 * CM_PER_INCH)
    }
}

impl From<Inch> for Mm {
    fn from(inch: Inch) -> Self {
        Mm(inch.0 * MM_PER_INCH)
    }
}

// === From Cm ===

impl From<Cm> for Pt {
    fn from(cm: Cm) -> Self {
        Pt(cm.0 / CM_PER_INCH * PT_PER_INCH)
    }
}

impl From<Cm> for Twip {
    fn from(cm: Cm) -> Self {
        Twip::from(Pt::from(cm))
    }
}

impl From<Cm> for Emu {
    fn from(cm: Cm) -> Self {
        Emu::from(Pt::from(cm))
    }
}

impl From<Cm> for Inch {
    fn from(cm: Cm) -> Self {
        Inch(cm.0 / CM_PER_INCH)
    }
}

impl From<Cm> for Mm {
    fn from(cm: Cm) -> Self {
        Mm(cm.0 * 10.0)
    }
}

// === From Mm ===

impl From<Mm> for Pt {
    fn from(mm: Mm) -> Self {
        Pt(mm.0 / MM_PER_INCH * PT_PER_INCH)
    }
}

impl From<Mm> for Twip {
    fn from(mm: Mm) -> Self {
        Twip::from(Pt::from(mm))
    }
}

impl From<Mm> for Emu {
    fn from(mm: Mm) -> Self {
        Emu::from(Pt::from(mm))
    }
}

impl From<Mm> for Inch {
    fn from(mm: Mm) -> Self {
        Inch(mm.0 / MM_PER_INCH)
    }
}

impl From<Mm> for Cm {
    fn from(mm: Mm) -> Self {
        Cm(mm.0 / 10.0)
    }
}

// === From HalfPt ===

impl From<HalfPt> for Pt {
    fn from(hp: HalfPt) -> Self {
        Pt(hp.0 as f64 / 2.0)
    }
}

// === Display ===

impl std::fmt::Display for Pt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}pt", self.0)
    }
}

impl std::fmt::Display for Emu {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}emu", self.0)
    }
}

impl std::fmt::Display for Twip {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}twip", self.0)
    }
}

impl std::fmt::Display for Cm {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}cm", self.0)
    }
}

impl std::fmt::Display for Mm {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}mm", self.0)
    }
}

impl std::fmt::Display for Inch {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}in", self.0)
    }
}

impl std::fmt::Display for HalfPt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}hp", self.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pt_to_twip() {
        let pt = Pt(12.0);
        let twip = Twip::from(pt);
        assert_eq!(twip, Twip(240));
    }

    #[test]
    fn test_inch_conversions() {
        let inch = Inch(1.0);
        assert_eq!(Pt::from(inch), Pt(72.0));
        assert_eq!(Twip::from(inch), Twip(1440));

        let cm = Cm::from(inch);
        assert!((cm.0 - 2.54).abs() < 0.001);
    }

    #[test]
    fn test_cm_to_inch() {
        let cm = Cm(2.54);
        let inch = Inch::from(cm);
        assert!((inch.0 - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_halfpt_to_pt() {
        let hp = HalfPt(24);
        assert_eq!(Pt::from(hp), Pt(12.0));
    }

    #[test]
    fn test_pt_to_halfpt() {
        let pt = Pt(14.0);
        assert_eq!(HalfPt::from(pt), HalfPt(28));
    }

    #[test]
    fn test_roundtrip_pt_twip() {
        let pt = Pt(12.5);
        let twip = Twip::from(pt);
        let pt2 = Pt::from(twip);
        assert!((pt.0 - pt2.0).abs() < 0.1);
    }

    #[test]
    fn test_mm_cm() {
        let cm = Cm(2.5);
        let mm = Mm::from(cm);
        assert!((mm.0 - 25.0).abs() < 0.001);

        let cm2 = Cm::from(mm);
        assert!((cm2.0 - 2.5).abs() < 0.001);
    }

    #[test]
    fn test_display() {
        assert_eq!(format!("{}", Pt(12.0)), "12pt");
        assert_eq!(format!("{}", Twip(240)), "240twip");
        assert_eq!(format!("{}", Cm(2.54)), "2.54cm");
        assert_eq!(format!("{}", Inch(1.0)), "1in");
        assert_eq!(format!("{}", HalfPt(24)), "24hp");
    }
}
