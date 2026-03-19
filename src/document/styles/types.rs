//! Style type definitions

use crate::document::{ParagraphProperties, RunProperties};
use crate::xml::RawXmlNode;

/// Style type
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl StyleType {
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "paragraph" => Some(Self::Paragraph),
            "character" => Some(Self::Character),
            "table" => Some(Self::Table),
            "numbering" => Some(Self::Numbering),
            _ => None,
        }
    }

    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Paragraph => "paragraph",
            Self::Character => "character",
            Self::Table => "table",
            Self::Numbering => "numbering",
        }
    }
}

/// A single style definition
#[derive(Clone, Debug, Default)]
pub struct Style {
    /// Style type (paragraph, character, table, numbering)
    pub style_type: Option<StyleType>,
    /// Style ID
    pub style_id: String,
    /// Display name
    pub name: Option<String>,
    /// Based-on style ID
    pub based_on: Option<String>,
    /// Next paragraph style ID
    pub next_style: Option<String>,
    /// Link to paired style
    pub link: Option<String>,
    /// Is this the default style for its type
    pub is_default: bool,
    /// Is this a custom style
    pub is_custom: bool,
    /// UI priority for sorting
    pub ui_priority: Option<u32>,
    /// Semi-hidden (not in style gallery by default)
    pub semi_hidden: bool,
    /// Unhide when used
    pub unhide_when_used: bool,
    /// Show in quick styles gallery
    pub qformat: bool,
    /// Paragraph properties
    pub paragraph_properties: Option<ParagraphProperties>,
    /// Run properties
    pub run_properties: Option<RunProperties>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
    /// Unknown attributes (preserved for round-trip)
    pub unknown_attrs: Vec<(String, String)>,
}

/// Document defaults (w:docDefaults)
#[derive(Clone, Debug, Default)]
pub struct DocDefaults {
    /// Default run properties
    pub run_properties: Option<RunProperties>,
    /// Default paragraph properties
    pub paragraph_properties: Option<ParagraphProperties>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}
