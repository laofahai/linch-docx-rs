//! Numbering-related types and enums

/// Number format
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum NumberFormat {
    /// 1, 2, 3
    Decimal,
    /// I, II, III
    UpperRoman,
    /// i, ii, iii
    LowerRoman,
    /// A, B, C
    UpperLetter,
    /// a, b, c
    LowerLetter,
    /// •
    Bullet,
    /// 一, 二, 三 (chineseCounting)
    ChineseCounting,
    /// 一, 二, 三 (chineseCountingThousand)
    ChineseCountingThousand,
    /// 壹, 贰, 叁 (ideographLegalTraditional)
    ChineseLegalTraditional,
    /// 甲, 乙, 丙 (ideographTraditional)
    IdeographTraditional,
    /// ㈠, ㈡, ㈢ (ideographEnclosedCircle)
    IdeographEnclosedCircle,
    /// 01, 02, 03 (decimalZero)
    DecimalZero,
    /// (一), (二), (三) (taiwaneseCounting)
    TaiwaneseCounting,
    /// None (no number)
    None,
    /// Other format (preserved as string)
    Other(String),
}

impl std::str::FromStr for NumberFormat {
    type Err = std::convert::Infallible;

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        Ok(match s {
            "decimal" => NumberFormat::Decimal,
            "upperRoman" => NumberFormat::UpperRoman,
            "lowerRoman" => NumberFormat::LowerRoman,
            "upperLetter" => NumberFormat::UpperLetter,
            "lowerLetter" => NumberFormat::LowerLetter,
            "bullet" => NumberFormat::Bullet,
            "chineseCounting" => NumberFormat::ChineseCounting,
            "chineseCountingThousand" => NumberFormat::ChineseCountingThousand,
            "ideographLegalTraditional" => NumberFormat::ChineseLegalTraditional,
            "ideographTraditional" => NumberFormat::IdeographTraditional,
            "ideographEnclosedCircle" => NumberFormat::IdeographEnclosedCircle,
            "decimalZero" => NumberFormat::DecimalZero,
            "taiwaneseCounting" => NumberFormat::TaiwaneseCounting,
            "none" => NumberFormat::None,
            other => NumberFormat::Other(other.to_string()),
        })
    }
}

impl NumberFormat {
    /// Convert to string
    pub fn as_str(&self) -> &str {
        match self {
            NumberFormat::Decimal => "decimal",
            NumberFormat::UpperRoman => "upperRoman",
            NumberFormat::LowerRoman => "lowerRoman",
            NumberFormat::UpperLetter => "upperLetter",
            NumberFormat::LowerLetter => "lowerLetter",
            NumberFormat::Bullet => "bullet",
            NumberFormat::ChineseCounting => "chineseCounting",
            NumberFormat::ChineseCountingThousand => "chineseCountingThousand",
            NumberFormat::ChineseLegalTraditional => "ideographLegalTraditional",
            NumberFormat::IdeographTraditional => "ideographTraditional",
            NumberFormat::IdeographEnclosedCircle => "ideographEnclosedCircle",
            NumberFormat::DecimalZero => "decimalZero",
            NumberFormat::TaiwaneseCounting => "taiwaneseCounting",
            NumberFormat::None => "none",
            NumberFormat::Other(s) => s,
        }
    }

    /// Check if this is a bullet format
    pub fn is_bullet(&self) -> bool {
        matches!(self, NumberFormat::Bullet)
    }

    /// Check if this is a numbered format (not bullet)
    pub fn is_numbered(&self) -> bool {
        !matches!(self, NumberFormat::Bullet | NumberFormat::None)
    }
}

/// Numbering information for a paragraph
#[derive(Clone, Debug)]
pub struct NumberingInfo {
    /// The numbering ID (references a Num definition)
    pub num_id: u32,
    /// The level (0-8)
    pub level: u32,
}

impl NumberingInfo {
    /// Create a new NumberingInfo
    pub fn new(num_id: u32, level: u32) -> Self {
        Self { num_id, level }
    }
}
