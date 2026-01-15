//! Numbering definitions (numbering.xml)
//!
//! This module handles list numbering in DOCX documents.

use crate::error::Result;
use crate::xml::{get_w_val, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::BufRead;

/// Numbering definitions from numbering.xml
#[derive(Clone, Debug, Default)]
pub struct Numbering {
    /// Abstract numbering definitions
    pub abstract_nums: HashMap<u32, AbstractNum>,
    /// Numbering instances
    pub nums: HashMap<u32, Num>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Abstract numbering definition (w:abstractNum)
#[derive(Clone, Debug, Default)]
pub struct AbstractNum {
    /// Abstract numbering ID
    pub abstract_num_id: u32,
    /// Multi-level type
    pub multi_level_type: Option<String>,
    /// Level definitions
    pub levels: HashMap<u8, Level>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Numbering instance (w:num)
#[derive(Clone, Debug)]
pub struct Num {
    /// Numbering ID (referenced by paragraphs)
    pub num_id: u32,
    /// Referenced abstract numbering ID
    pub abstract_num_id: u32,
    /// Level overrides
    pub level_overrides: Vec<LevelOverride>,
}

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
    /// 一, 二, 三
    ChineseCountingThousand,
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
            "chineseCountingThousand" => NumberFormat::ChineseCountingThousand,
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
            NumberFormat::ChineseCountingThousand => "chineseCountingThousand",
            NumberFormat::None => "none",
            NumberFormat::Other(s) => s,
        }
    }

    /// Check if this is a bullet format
    pub fn is_bullet(&self) -> bool {
        matches!(self, NumberFormat::Bullet)
    }
}

impl Numbering {
    /// Parse numbering.xml content
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut numbering = Numbering::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"abstractNum" => {
                            let abs_num = AbstractNum::from_reader(&mut reader, &e)?;
                            numbering
                                .abstract_nums
                                .insert(abs_num.abstract_num_id, abs_num);
                        }
                        b"num" => {
                            let num = Num::from_reader(&mut reader, &e)?;
                            numbering.nums.insert(num.num_id, num);
                        }
                        b"numbering" => {
                            // Root element, continue
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            numbering.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    // Empty elements at root level - preserve
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
                    numbering.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(numbering)
    }

    /// Serialize to XML
    pub fn to_xml(&self) -> Result<String> {
        let mut buffer = Vec::new();
        let mut writer = Writer::new(&mut buffer);

        // XML declaration
        writer.write_event(Event::Decl(quick_xml::events::BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with namespaces
        let mut start = BytesStart::new("w:numbering");
        start.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(start))?;

        // Write abstract nums (sorted by ID for consistency)
        let mut abs_ids: Vec<_> = self.abstract_nums.keys().collect();
        abs_ids.sort();
        for id in abs_ids {
            self.abstract_nums[id].write_to(&mut writer)?;
        }

        // Write nums (sorted by ID)
        let mut num_ids: Vec<_> = self.nums.keys().collect();
        num_ids.sort();
        for id in num_ids {
            self.nums[id].write_to(&mut writer)?;
        }

        // Write unknown children
        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;

        String::from_utf8(buffer).map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get the format for a specific numId and level
    pub fn get_format(&self, num_id: u32, level: u8) -> Option<&NumberFormat> {
        let num = self.nums.get(&num_id)?;
        let abs_num = self.abstract_nums.get(&num.abstract_num_id)?;
        let lvl = abs_num.levels.get(&level)?;
        lvl.num_fmt.as_ref()
    }

    /// Check if a numId represents a bullet list
    pub fn is_bullet_list(&self, num_id: u32) -> bool {
        if let Some(fmt) = self.get_format(num_id, 0) {
            fmt.is_bullet()
        } else {
            false
        }
    }

    /// Get level text for a specific numId and level
    pub fn get_level_text(&self, num_id: u32, level: u8) -> Option<&str> {
        let num = self.nums.get(&num_id)?;
        let abs_num = self.abstract_nums.get(&num.abstract_num_id)?;
        let lvl = abs_num.levels.get(&level)?;
        lvl.level_text.as_deref()
    }
}

impl AbstractNum {
    fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut abs_num = AbstractNum::default();

        // Get abstractNumId attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:abstractNumId" || key == b"abstractNumId" {
                abs_num.abstract_num_id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"lvl" => {
                            let lvl = Level::from_reader(reader, &e)?;
                            abs_num.levels.insert(lvl.ilvl, lvl);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            abs_num.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"multiLevelType" => {
                            abs_num.multi_level_type = get_w_val(&e);
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
                            abs_num.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"abstractNum" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(abs_num)
    }

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:abstractNum");
        start.push_attribute(("w:abstractNumId", self.abstract_num_id.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Multi-level type
        if let Some(mlt) = &self.multi_level_type {
            let mut elem = BytesStart::new("w:multiLevelType");
            elem.push_attribute(("w:val", mlt.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Write levels (sorted)
        let mut level_ids: Vec<_> = self.levels.keys().collect();
        level_ids.sort();
        for id in level_ids {
            self.levels[id].write_to(writer)?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
        Ok(())
    }
}

impl Num {
    fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut num_id = 0u32;
        let mut abstract_num_id = 0u32;
        let mut level_overrides = Vec::new();

        // Get numId attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:numId" || key == b"numId" {
                num_id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"lvlOverride" {
                        let lo = LevelOverride::from_reader(reader, &e)?;
                        level_overrides.push(lo);
                    } else {
                        skip_element(reader, &e)?;
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"abstractNumId" {
                        abstract_num_id = get_w_val(&e).and_then(|v| v.parse().ok()).unwrap_or(0);
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"num" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(Num {
            num_id,
            abstract_num_id,
            level_overrides,
        })
    }

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:num");
        start.push_attribute(("w:numId", self.num_id.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Abstract num ID
        let mut elem = BytesStart::new("w:abstractNumId");
        elem.push_attribute(("w:val", self.abstract_num_id.to_string().as_str()));
        writer.write_event(Event::Empty(elem))?;

        // Level overrides
        for lo in &self.level_overrides {
            lo.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:num")))?;
        Ok(())
    }
}

impl Level {
    fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
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

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
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
    fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
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
                        skip_element(reader, &e)?;
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

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
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
    fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
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
                    } else {
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

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
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
    fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
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

    fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        Ok(())
    }
}

/// Skip an element and all its children
fn skip_element<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<()> {
    let name = start.name();
    let mut depth = 1;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name() == name => depth += 1,
            Event::End(e) if e.name() == name => {
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

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_NUMBERING: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>"#;

    #[test]
    fn test_parse_numbering() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();

        // Should have 2 abstract nums
        assert_eq!(numbering.abstract_nums.len(), 2);

        // Should have 2 nums
        assert_eq!(numbering.nums.len(), 2);

        // Check abstract num 0
        let abs0 = numbering.abstract_nums.get(&0).unwrap();
        assert_eq!(abs0.multi_level_type, Some("hybridMultilevel".to_string()));
        assert_eq!(abs0.levels.len(), 2);

        // Check level 0
        let lvl0 = abs0.levels.get(&0).unwrap();
        assert_eq!(lvl0.start, Some(1));
        assert_eq!(lvl0.num_fmt, Some(NumberFormat::Decimal));
        assert_eq!(lvl0.level_text, Some("%1.".to_string()));

        // Check num 1 references abstract num 0
        let num1 = numbering.nums.get(&1).unwrap();
        assert_eq!(num1.abstract_num_id, 0);
    }

    #[test]
    fn test_is_bullet_list() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();

        // numId 1 references abstractNum 0 which is decimal
        assert!(!numbering.is_bullet_list(1));

        // numId 2 references abstractNum 1 which is bullet
        assert!(numbering.is_bullet_list(2));
    }

    #[test]
    fn test_roundtrip() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();
        let xml = numbering.to_xml().unwrap();

        // Parse again
        let numbering2 = Numbering::from_xml(&xml).unwrap();

        // Should have same structure
        assert_eq!(
            numbering.abstract_nums.len(),
            numbering2.abstract_nums.len()
        );
        assert_eq!(numbering.nums.len(), numbering2.nums.len());

        // Check a specific value
        assert_eq!(numbering.get_format(1, 0), numbering2.get_format(1, 0));
    }
}
