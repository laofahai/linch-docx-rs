//! Paragraph properties and related types

use crate::error::Result;
use crate::xml::{get_w_val, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Paragraph alignment
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

impl Alignment {
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "left" | "start" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" | "end" => Some(Self::Right),
            "both" | "justify" => Some(Self::Justify),
            "distribute" => Some(Self::Distribute),
            _ => None,
        }
    }

    pub fn as_ooxml(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "both",
            Self::Distribute => "distribute",
        }
    }
}

/// Indentation settings (in twips)
#[derive(Clone, Debug, Default)]
pub struct Indentation {
    pub left: Option<i32>,
    pub right: Option<i32>,
    pub first_line: Option<i32>,
    pub hanging: Option<i32>,
}

/// Line spacing settings
#[derive(Clone, Debug, Default)]
pub struct LineSpacing {
    /// Space before paragraph (in twips)
    pub before: Option<u32>,
    /// Space after paragraph (in twips)
    pub after: Option<u32>,
    /// Line spacing value (in 240ths of a line for "auto", twips for exact/atLeast)
    pub line: Option<u32>,
    /// Line spacing rule
    pub line_rule: Option<String>,
}

/// Paragraph properties (w:pPr)
#[derive(Clone, Debug, Default)]
pub struct ParagraphProperties {
    /// Style ID
    pub style: Option<String>,
    /// Justification/alignment
    pub justification: Option<String>,
    /// Numbering properties
    pub num_id: Option<u32>,
    pub num_level: Option<u32>,
    /// Outline level (for headings)
    pub outline_level: Option<u8>,
    /// Indentation
    pub indentation: Option<Indentation>,
    /// Line spacing
    pub spacing: Option<LineSpacing>,
    /// Keep with next paragraph
    pub keep_next: Option<bool>,
    /// Keep lines together
    pub keep_lines: Option<bool>,
    /// Page break before
    pub page_break_before: Option<bool>,
    /// Run properties for paragraph mark
    pub run_properties: Option<crate::document::RunProperties>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl ParagraphProperties {
    /// Parse from reader (after w:pPr start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = ParagraphProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"numPr" => {
                            parse_num_pr(reader, &mut props)?;
                        }
                        b"rPr" => {
                            props.run_properties =
                                Some(crate::document::RunProperties::from_reader(reader)?);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            props.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"pStyle" => props.style = get_w_val(&e),
                        b"jc" => props.justification = get_w_val(&e),
                        b"outlineLvl" => {
                            props.outline_level = get_w_val(&e).and_then(|v| v.parse().ok());
                        }
                        b"ind" => {
                            props.indentation = Some(Indentation {
                                left: crate::xml::get_attr(&e, "w:left")
                                    .and_then(|v| v.parse().ok()),
                                right: crate::xml::get_attr(&e, "w:right")
                                    .and_then(|v| v.parse().ok()),
                                first_line: crate::xml::get_attr(&e, "w:firstLine")
                                    .and_then(|v| v.parse().ok()),
                                hanging: crate::xml::get_attr(&e, "w:hanging")
                                    .and_then(|v| v.parse().ok()),
                            });
                        }
                        b"spacing" => {
                            props.spacing = Some(LineSpacing {
                                before: crate::xml::get_attr(&e, "w:before")
                                    .and_then(|v| v.parse().ok()),
                                after: crate::xml::get_attr(&e, "w:after")
                                    .and_then(|v| v.parse().ok()),
                                line: crate::xml::get_attr(&e, "w:line")
                                    .and_then(|v| v.parse().ok()),
                                line_rule: crate::xml::get_attr(&e, "w:lineRule"),
                            });
                        }
                        b"keepNext" => props.keep_next = Some(crate::xml::parse_bool(&e)),
                        b"keepLines" => props.keep_lines = Some(crate::xml::parse_bool(&e)),
                        b"pageBreakBefore" => {
                            props.page_break_before = Some(crate::xml::parse_bool(&e));
                        }
                        _ => {
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
                Event::End(e) if e.name().local_name().as_ref() == b"pPr" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(props)
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let has_content = self.style.is_some()
            || self.justification.is_some()
            || self.num_id.is_some()
            || self.outline_level.is_some()
            || self.indentation.is_some()
            || self.spacing.is_some()
            || self.keep_next.is_some()
            || self.keep_lines.is_some()
            || self.page_break_before.is_some()
            || self.run_properties.is_some()
            || !self.unknown_children.is_empty();

        if !has_content {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        if let Some(style) = &self.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        if let Some(true) = self.keep_next {
            writer.write_event(Event::Empty(BytesStart::new("w:keepNext")))?;
        }
        if let Some(true) = self.keep_lines {
            writer.write_event(Event::Empty(BytesStart::new("w:keepLines")))?;
        }
        if let Some(true) = self.page_break_before {
            writer.write_event(Event::Empty(BytesStart::new("w:pageBreakBefore")))?;
        }

        if self.num_id.is_some() || self.num_level.is_some() {
            writer.write_event(Event::Start(BytesStart::new("w:numPr")))?;
            if let Some(level) = self.num_level {
                let mut elem = BytesStart::new("w:ilvl");
                elem.push_attribute(("w:val", level.to_string().as_str()));
                writer.write_event(Event::Empty(elem))?;
            }
            if let Some(num_id) = self.num_id {
                let mut elem = BytesStart::new("w:numId");
                elem.push_attribute(("w:val", num_id.to_string().as_str()));
                writer.write_event(Event::Empty(elem))?;
            }
            writer.write_event(Event::End(BytesEnd::new("w:numPr")))?;
        }

        if let Some(ref sp) = self.spacing {
            let mut elem = BytesStart::new("w:spacing");
            if let Some(v) = sp.before {
                elem.push_attribute(("w:before", v.to_string().as_str()));
            }
            if let Some(v) = sp.after {
                elem.push_attribute(("w:after", v.to_string().as_str()));
            }
            if let Some(v) = sp.line {
                elem.push_attribute(("w:line", v.to_string().as_str()));
            }
            if let Some(ref rule) = sp.line_rule {
                elem.push_attribute(("w:lineRule", rule.as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        if let Some(ref ind) = self.indentation {
            let mut elem = BytesStart::new("w:ind");
            if let Some(v) = ind.left {
                elem.push_attribute(("w:left", v.to_string().as_str()));
            }
            if let Some(v) = ind.right {
                elem.push_attribute(("w:right", v.to_string().as_str()));
            }
            if let Some(v) = ind.first_line {
                elem.push_attribute(("w:firstLine", v.to_string().as_str()));
            }
            if let Some(v) = ind.hanging {
                elem.push_attribute(("w:hanging", v.to_string().as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        if let Some(jc) = &self.justification {
            let mut elem = BytesStart::new("w:jc");
            elem.push_attribute(("w:val", jc.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        if let Some(level) = self.outline_level {
            let mut elem = BytesStart::new("w:outlineLvl");
            elem.push_attribute(("w:val", level.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        if let Some(ref rpr) = self.run_properties {
            rpr.write_to(writer)?;
        }

        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        Ok(())
    }
}

/// Parse numbering properties
fn parse_num_pr<R: BufRead>(reader: &mut Reader<R>, props: &mut ParagraphProperties) -> Result<()> {
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) => {
                let local = e.name().local_name();
                match local.as_ref() {
                    b"numId" => props.num_id = get_w_val(&e).and_then(|v| v.parse().ok()),
                    b"ilvl" => props.num_level = get_w_val(&e).and_then(|v| v.parse().ok()),
                    _ => {}
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"numPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}
