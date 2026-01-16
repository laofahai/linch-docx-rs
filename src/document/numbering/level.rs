//! Level definitions for numbering

use crate::error::Result;
use crate::xml::{get_w_val, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

use super::types::NumberFormat;

/// Level definition (w:lvl)
#[derive(Clone, Debug, Default)]
pub struct Level {
    /// Level index (0-8)
    pub ilvl: u8,
    /// Start value
    pub start: Option<u32>,
    /// Number format
    pub num_fmt: Option<NumberFormat>,
    /// Level text (e.g., "%1.", "%1.%2.")
    pub level_text: Option<String>,
    /// Level justification
    pub lvl_jc: Option<String>,
    /// Paragraph properties for this level
    pub p_pr: Option<LevelParagraphProperties>,
    /// Run properties for this level
    pub r_pr: Option<LevelRunProperties>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Level override (w:lvlOverride)
#[derive(Clone, Debug)]
pub struct LevelOverride {
    /// Level index
    pub ilvl: u8,
    /// Start override
    pub start_override: Option<u32>,
    /// Level definition override
    pub lvl: Option<Level>,
}

/// Simplified paragraph properties for numbering levels
#[derive(Clone, Debug, Default)]
pub struct LevelParagraphProperties {
    /// Left indentation (twips)
    pub ind_left: Option<i32>,
    /// Hanging indentation (twips)
    pub ind_hanging: Option<i32>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Simplified run properties for numbering levels
#[derive(Clone, Debug, Default)]
pub struct LevelRunProperties {
    /// Unknown children (preserved - we don't parse run props deeply here)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Level {
    /// Create a new level with the given index
    pub fn new(ilvl: u8) -> Self {
        Level {
            ilvl,
            start: Some(1),
            ..Default::default()
        }
    }

    /// Set the number format
    pub fn with_format(mut self, fmt: NumberFormat) -> Self {
        self.num_fmt = Some(fmt);
        self
    }

    /// Set the level text
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.level_text = Some(text.into());
        self
    }

    /// Set the start value
    pub fn with_start(mut self, start: u32) -> Self {
        self.start = Some(start);
        self
    }

    /// Set the justification
    pub fn with_justification(mut self, jc: impl Into<String>) -> Self {
        self.lvl_jc = Some(jc.into());
        self
    }

    pub(crate) fn from_reader<R: BufRead>(
        reader: &mut Reader<R>,
        start: &BytesStart,
    ) -> Result<Self> {
        let mut level = Level::default();

        // Get ilvl attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:ilvl" || key == b"ilvl" {
                level.ilvl = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"pPr" => {
                            level.p_pr = Some(LevelParagraphProperties::from_reader(reader)?);
                        }
                        b"rPr" => {
                            level.r_pr = Some(LevelRunProperties::from_reader(reader)?);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            level.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"start" => {
                            level.start = get_w_val(&e).and_then(|v| v.parse().ok());
                        }
                        b"numFmt" => {
                            level.num_fmt = get_w_val(&e).map(|v| v.parse().unwrap());
                        }
                        b"lvlText" => {
                            level.level_text = get_w_val(&e);
                        }
                        b"lvlJc" => {
                            level.lvl_jc = get_w_val(&e);
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
                            level.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"lvl" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(level)
    }

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:lvl");
        start.push_attribute(("w:ilvl", self.ilvl.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Start value
        if let Some(s) = self.start {
            let mut elem = BytesStart::new("w:start");
            elem.push_attribute(("w:val", s.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Number format
        if let Some(fmt) = &self.num_fmt {
            let mut elem = BytesStart::new("w:numFmt");
            elem.push_attribute(("w:val", fmt.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Level text
        if let Some(txt) = &self.level_text {
            let mut elem = BytesStart::new("w:lvlText");
            elem.push_attribute(("w:val", txt.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Level justification
        if let Some(jc) = &self.lvl_jc {
            let mut elem = BytesStart::new("w:lvlJc");
            elem.push_attribute(("w:val", jc.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Paragraph properties
        if let Some(p_pr) = &self.p_pr {
            p_pr.write_to(writer)?;
        }

        // Run properties
        if let Some(r_pr) = &self.r_pr {
            r_pr.write_to(writer)?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:lvl")))?;
        Ok(())
    }
}

impl LevelOverride {
    pub(crate) fn from_reader<R: BufRead>(
        reader: &mut Reader<R>,
        start: &BytesStart,
    ) -> Result<Self> {
        let mut ilvl = 0u8;
        let mut start_override = None;
        let mut lvl = None;

        // Get ilvl attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:ilvl" || key == b"ilvl" {
                ilvl = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"lvl" {
                        lvl = Some(Level::from_reader(reader, &e)?);
                    } else {
                        super::skip_element(reader, &e)?;
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"startOverride" {
                        start_override = get_w_val(&e).and_then(|v| v.parse().ok());
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"lvlOverride" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(LevelOverride {
            ilvl,
            start_override,
            lvl,
        })
    }

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:lvlOverride");
        start.push_attribute(("w:ilvl", self.ilvl.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Start override
        if let Some(s) = self.start_override {
            let mut elem = BytesStart::new("w:startOverride");
            elem.push_attribute(("w:val", s.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Level override
        if let Some(lvl) = &self.lvl {
            lvl.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:lvlOverride")))?;
        Ok(())
    }
}

impl LevelParagraphProperties {
    pub(crate) fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = LevelParagraphProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let raw = RawXmlElement::from_reader(reader, &e)?;
                    props.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"ind" {
                        // Parse indentation
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                b"w:left" | b"left" => {
                                    props.ind_left = val.parse().ok();
                                }
                                b"w:hanging" | b"hanging" => {
                                    props.ind_hanging = val.parse().ok();
                                }
                                _ => {}
                            }
                        }
                    }
                    // Also preserve as unknown for complete round-trip
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
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"pPr" {
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

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        // Write unknown children (which includes ind if it was preserved)
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        Ok(())
    }
}

impl LevelRunProperties {
    pub(crate) fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = LevelRunProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let raw = RawXmlElement::from_reader(reader, &e)?;
                    props.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::Empty(e) => {
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

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        Ok(())
    }
}
