//! Run element (w:r) - a contiguous run of text with uniform formatting

use crate::error::Result;
use crate::xml::{get_w_val, parse_bool, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Run element (w:r)
#[derive(Clone, Debug, Default)]
pub struct Run {
    /// Run properties
    pub properties: Option<RunProperties>,
    /// Run content
    pub content: Vec<RunContent>,
    /// Unknown attributes (preserved)
    pub unknown_attrs: Vec<(String, String)>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Content within a run
#[derive(Clone, Debug)]
pub enum RunContent {
    /// Text (w:t)
    Text(String),
    /// Tab (w:tab)
    Tab,
    /// Break (w:br)
    Break(BreakType),
    /// Carriage return (w:cr)
    CarriageReturn,
    /// Soft hyphen
    SoftHyphen,
    /// Non-breaking hyphen
    NoBreakHyphen,
    /// Unknown (preserved)
    Unknown(RawXmlNode),
}

/// Break type
#[derive(Clone, Debug, Default)]
pub enum BreakType {
    #[default]
    TextWrapping,
    Page,
    Column,
}

/// Run properties (w:rPr)
#[derive(Clone, Debug, Default)]
pub struct RunProperties {
    /// Style ID
    pub style: Option<String>,
    /// Bold
    pub bold: Option<bool>,
    /// Italic
    pub italic: Option<bool>,
    /// Underline type
    pub underline: Option<String>,
    /// Strike-through
    pub strike: Option<bool>,
    /// Double strike-through
    pub double_strike: Option<bool>,
    /// Font size (in half-points, e.g., 24 = 12pt)
    pub size: Option<u32>,
    /// Color (RGB hex)
    pub color: Option<String>,
    /// Highlight color
    pub highlight: Option<String>,
    /// Font (ASCII)
    pub font_ascii: Option<String>,
    /// Font (East Asia)
    pub font_east_asia: Option<String>,
    /// Vertical alignment (superscript/subscript)
    pub vertical_align: Option<String>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Run {
    /// Parse from reader (after w:r start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut run = Run::default();

        // Parse attributes
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            run.unknown_attrs.push((key, value));
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"rPr" => {
                            run.properties = Some(RunProperties::from_reader(reader)?);
                        }
                        b"t" => {
                            // Read text content
                            let text = read_text_content(reader)?;
                            run.content.push(RunContent::Text(text));
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            run.content.push(RunContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::Empty(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"t" => {
                            // Empty text element
                            run.content.push(RunContent::Text(String::new()));
                        }
                        b"tab" => {
                            run.content.push(RunContent::Tab);
                        }
                        b"br" => {
                            let break_type = match crate::xml::get_attr(&e, "w:type")
                                .or_else(|| crate::xml::get_attr(&e, "type"))
                                .as_deref()
                            {
                                Some("page") => BreakType::Page,
                                Some("column") => BreakType::Column,
                                _ => BreakType::TextWrapping,
                            };
                            run.content.push(RunContent::Break(break_type));
                        }
                        b"cr" => {
                            run.content.push(RunContent::CarriageReturn);
                        }
                        b"softHyphen" => {
                            run.content.push(RunContent::SoftHyphen);
                        }
                        b"noBreakHyphen" => {
                            run.content.push(RunContent::NoBreakHyphen);
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement {
                                name: String::from_utf8_lossy(e.name().as_ref()).to_string(),
                                attributes: e
                                    .attributes()
                                    .filter_map(|a| a.ok())
                                    .map(|a| {
                                        (
                                            String::from_utf8_lossy(a.key.as_ref()).to_string(),
                                            String::from_utf8_lossy(&a.value).to_string(),
                                        )
                                    })
                                    .collect(),
                                children: Vec::new(),
                                self_closing: true,
                            };
                            run.content.push(RunContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"r" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(run)
    }

    /// Create from empty element
    pub fn from_empty(start: &BytesStart) -> Result<Self> {
        let mut run = Run::default();

        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            run.unknown_attrs.push((key, value));
        }

        Ok(run)
    }

    /// Get all text in this run
    pub fn text(&self) -> String {
        let mut result = String::new();
        for content in &self.content {
            match content {
                RunContent::Text(t) => result.push_str(t),
                RunContent::Tab => result.push('\t'),
                RunContent::Break(BreakType::TextWrapping) => result.push('\n'),
                RunContent::CarriageReturn => result.push('\n'),
                _ => {}
            }
        }
        result
    }

    /// Check if bold
    pub fn bold(&self) -> bool {
        self.properties.as_ref().and_then(|p| p.bold).unwrap_or(false)
    }

    /// Check if italic
    pub fn italic(&self) -> bool {
        self.properties.as_ref().and_then(|p| p.italic).unwrap_or(false)
    }

    /// Get font size in points (None if not specified)
    pub fn font_size_pt(&self) -> Option<f32> {
        self.properties.as_ref()?.size.map(|s| s as f32 / 2.0)
    }

    /// Get color (RGB hex string)
    pub fn color(&self) -> Option<&str> {
        self.properties.as_ref()?.color.as_deref()
    }

    /// Get underline type
    pub fn underline(&self) -> Option<&str> {
        self.properties.as_ref()?.underline.as_deref()
    }

    /// Check if has strike-through
    pub fn strike(&self) -> bool {
        self.properties.as_ref().and_then(|p| p.strike).unwrap_or(false)
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:r");
        for (key, value) in &self.unknown_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        // Check if run is empty (no properties, no content, no unknown children)
        let is_empty = self.properties.is_none()
            && self.content.is_empty()
            && self.unknown_children.is_empty();

        if is_empty {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;

            // Write properties
            if let Some(props) = &self.properties {
                props.write_to(writer)?;
            }

            // Write content
            for content in &self.content {
                content.write_to(writer)?;
            }

            // Write unknown children
            for child in &self.unknown_children {
                child.write_to(writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        }

        Ok(())
    }

    /// Create a new run with text
    pub fn new(text: impl Into<String>) -> Self {
        Run {
            content: vec![RunContent::Text(text.into())],
            ..Default::default()
        }
    }

    /// Set bold
    pub fn set_bold(&mut self, bold: bool) {
        self.properties.get_or_insert_with(Default::default).bold = Some(bold);
    }

    /// Set italic
    pub fn set_italic(&mut self, italic: bool) {
        self.properties.get_or_insert_with(Default::default).italic = Some(italic);
    }

    /// Set font size in points
    pub fn set_font_size_pt(&mut self, size: f32) {
        self.properties.get_or_insert_with(Default::default).size = Some((size * 2.0) as u32);
    }

    /// Set color (RGB hex string)
    pub fn set_color(&mut self, color: impl Into<String>) {
        self.properties.get_or_insert_with(Default::default).color = Some(color.into());
    }
}

impl RunContent {
    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            RunContent::Text(text) => {
                let mut start = BytesStart::new("w:t");
                // Preserve space if text has leading/trailing whitespace
                if text.starts_with(' ') || text.ends_with(' ') || text.contains("  ") {
                    start.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(text)))?;
                writer.write_event(Event::End(BytesEnd::new("w:t")))?;
            }
            RunContent::Tab => {
                writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
            }
            RunContent::Break(break_type) => {
                let mut start = BytesStart::new("w:br");
                match break_type {
                    BreakType::Page => start.push_attribute(("w:type", "page")),
                    BreakType::Column => start.push_attribute(("w:type", "column")),
                    BreakType::TextWrapping => {}
                }
                writer.write_event(Event::Empty(start))?;
            }
            RunContent::CarriageReturn => {
                writer.write_event(Event::Empty(BytesStart::new("w:cr")))?;
            }
            RunContent::SoftHyphen => {
                writer.write_event(Event::Empty(BytesStart::new("w:softHyphen")))?;
            }
            RunContent::NoBreakHyphen => {
                writer.write_event(Event::Empty(BytesStart::new("w:noBreakHyphen")))?;
            }
            RunContent::Unknown(node) => {
                node.write_to(writer)?;
            }
        }
        Ok(())
    }
}

impl RunProperties {
    /// Parse from reader (after w:rPr start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = RunProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"rFonts" => {
                            // Read font info then skip
                            props.font_ascii = crate::xml::get_attr(&e, "w:ascii")
                                .or_else(|| crate::xml::get_attr(&e, "ascii"));
                            props.font_east_asia = crate::xml::get_attr(&e, "w:eastAsia")
                                .or_else(|| crate::xml::get_attr(&e, "eastAsia"));
                            // Skip to end
                            skip_element(reader, &e)?;
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            props.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"rStyle" => {
                            props.style = get_w_val(&e);
                        }
                        b"b" => {
                            props.bold = Some(parse_bool(&e));
                        }
                        b"bCs" => {
                            // Complex script bold - ignore for now
                        }
                        b"i" => {
                            props.italic = Some(parse_bool(&e));
                        }
                        b"iCs" => {
                            // Complex script italic - ignore for now
                        }
                        b"u" => {
                            props.underline = get_w_val(&e).or(Some("single".into()));
                        }
                        b"strike" => {
                            props.strike = Some(parse_bool(&e));
                        }
                        b"dstrike" => {
                            props.double_strike = Some(parse_bool(&e));
                        }
                        b"sz" => {
                            props.size = get_w_val(&e).and_then(|v| v.parse().ok());
                        }
                        b"szCs" => {
                            // Complex script size - ignore for now
                        }
                        b"color" => {
                            props.color = get_w_val(&e);
                        }
                        b"highlight" => {
                            props.highlight = get_w_val(&e);
                        }
                        b"vertAlign" => {
                            props.vertical_align = get_w_val(&e);
                        }
                        b"rFonts" => {
                            props.font_ascii = crate::xml::get_attr(&e, "w:ascii")
                                .or_else(|| crate::xml::get_attr(&e, "ascii"));
                            props.font_east_asia = crate::xml::get_attr(&e, "w:eastAsia")
                                .or_else(|| crate::xml::get_attr(&e, "eastAsia"));
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement {
                                name: String::from_utf8_lossy(e.name().as_ref()).to_string(),
                                attributes: e
                                    .attributes()
                                    .filter_map(|a| a.ok())
                                    .map(|a| {
                                        (
                                            String::from_utf8_lossy(a.key.as_ref()).to_string(),
                                            String::from_utf8_lossy(&a.value).to_string(),
                                        )
                                    })
                                    .collect(),
                                children: Vec::new(),
                                self_closing: true,
                            };
                            props.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"rPr" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(props)
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // Check if there are any properties to write
        let has_content = self.style.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.strike.is_some()
            || self.double_strike.is_some()
            || self.size.is_some()
            || self.color.is_some()
            || self.highlight.is_some()
            || self.font_ascii.is_some()
            || self.vertical_align.is_some()
            || !self.unknown_children.is_empty();

        if !has_content {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        // Style
        if let Some(style) = &self.style {
            let mut elem = BytesStart::new("w:rStyle");
            elem.push_attribute(("w:val", style.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Fonts
        if self.font_ascii.is_some() || self.font_east_asia.is_some() {
            let mut elem = BytesStart::new("w:rFonts");
            if let Some(font) = &self.font_ascii {
                elem.push_attribute(("w:ascii", font.as_str()));
            }
            if let Some(font) = &self.font_east_asia {
                elem.push_attribute(("w:eastAsia", font.as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Bold
        if let Some(bold) = self.bold {
            let mut elem = BytesStart::new("w:b");
            if !bold {
                elem.push_attribute(("w:val", "0"));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Italic
        if let Some(italic) = self.italic {
            let mut elem = BytesStart::new("w:i");
            if !italic {
                elem.push_attribute(("w:val", "0"));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Strike
        if let Some(strike) = self.strike {
            let mut elem = BytesStart::new("w:strike");
            if !strike {
                elem.push_attribute(("w:val", "0"));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Double strike
        if let Some(dstrike) = self.double_strike {
            let mut elem = BytesStart::new("w:dstrike");
            if !dstrike {
                elem.push_attribute(("w:val", "0"));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Underline
        if let Some(underline) = &self.underline {
            let mut elem = BytesStart::new("w:u");
            elem.push_attribute(("w:val", underline.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Color
        if let Some(color) = &self.color {
            let mut elem = BytesStart::new("w:color");
            elem.push_attribute(("w:val", color.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Size
        if let Some(size) = self.size {
            let mut elem = BytesStart::new("w:sz");
            elem.push_attribute(("w:val", size.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Highlight
        if let Some(highlight) = &self.highlight {
            let mut elem = BytesStart::new("w:highlight");
            elem.push_attribute(("w:val", highlight.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Vertical align
        if let Some(valign) = &self.vertical_align {
            let mut elem = BytesStart::new("w:vertAlign");
            elem.push_attribute(("w:val", valign.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        Ok(())
    }
}

/// Read text content from w:t element
fn read_text_content<R: BufRead>(reader: &mut Reader<R>) -> Result<String> {
    let mut text = String::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(t) => {
                text.push_str(&t.unescape()?);
            }
            Event::End(e) => {
                if e.name().local_name().as_ref() == b"t" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(text)
}

/// Skip to end of element
fn skip_element<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<()> {
    let target_name = start.name().as_ref().to_vec();
    let mut depth = 1;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == target_name => depth += 1,
            Event::End(e) if e.name().as_ref() == target_name => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}
