//! Table-related types and enums

/// Grid column definition
#[derive(Clone, Debug, Default)]
pub struct GridColumn {
    /// Width in twips
    pub width: Option<i32>,
}

/// Vertical merge type
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum VMerge {
    /// Start of a new vertical merge group
    Restart,
    /// Continuation of a vertical merge
    Continue,
}

/// Table width specification
#[derive(Clone, Debug, Default, PartialEq)]
pub enum TableWidth {
    /// Automatic width
    #[default]
    Auto,
    /// Width as percentage (0.0 - 100.0)
    Percent(f64),
    /// Width in twips (1/20 of a point)
    Twips(i32),
}

/// Table alignment
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TableAlignment {
    /// Left aligned (default)
    #[default]
    Left,
    /// Center aligned
    Center,
    /// Right aligned
    Right,
}

impl TableAlignment {
    /// Parse from OOXML string value
    pub fn parse(s: &str) -> Self {
        match s {
            "center" => TableAlignment::Center,
            "right" | "end" => TableAlignment::Right,
            _ => TableAlignment::Left,
        }
    }

    /// Convert to OOXML string value
    pub fn as_str(&self) -> &'static str {
        match self {
            TableAlignment::Left => "left",
            TableAlignment::Center => "center",
            TableAlignment::Right => "right",
        }
    }
}

/// Vertical alignment for table cells
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum VerticalAlignment {
    /// Top aligned (default)
    #[default]
    Top,
    /// Center aligned
    Center,
    /// Bottom aligned
    Bottom,
}

impl VerticalAlignment {
    /// Parse from OOXML string value
    pub fn parse(s: &str) -> Self {
        match s {
            "center" => VerticalAlignment::Center,
            "bottom" => VerticalAlignment::Bottom,
            _ => VerticalAlignment::Top,
        }
    }

    /// Convert to OOXML string value
    pub fn as_str(&self) -> &'static str {
        match self {
            VerticalAlignment::Top => "top",
            VerticalAlignment::Center => "center",
            VerticalAlignment::Bottom => "bottom",
        }
    }
}
