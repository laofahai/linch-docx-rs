//! Paragraph element (w:p)

use crate::document::Run;
use crate::error::Result;
use crate::xml::{get_w_val, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Paragraph element (w:p)
#[derive(Clone, Debug, Default)]
pub struct Paragraph {
    /// Paragraph properties
    pub properties: Option<ParagraphProperties>,
    /// Paragraph content (runs, hyperlinks, etc.)
    pub content: Vec<ParagraphContent>,
    /// Unknown attributes (preserved for round-trip)
    pub unknown_attrs: Vec<(String, String)>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Content within a paragraph
#[derive(Clone, Debug)]
pub enum ParagraphContent {
    /// Text run
    Run(Run),
    /// Hyperlink
    Hyperlink(Hyperlink),
    /// Bookmark start
    BookmarkStart { id: String, name: String },
    /// Bookmark end
    BookmarkEnd { id: String },
    /// Unknown element (preserved)
    Unknown(RawXmlNode),
}

/// Hyperlink element
#[derive(Clone, Debug, Default)]
pub struct Hyperlink {
    /// Relationship ID (for external links)
    pub r_id: Option<String>,
    /// Anchor (for internal links)
    pub anchor: Option<String>,
    /// Content runs
    pub runs: Vec<Run>,
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
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Paragraph {
    /// Parse paragraph from reader (after w:p start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut para = Paragraph::default();

        // Parse attributes
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            // Store all attributes (including rsid* for round-trip)
            para.unknown_attrs.push((key, value));
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"pPr" => {
                            para.properties = Some(ParagraphProperties::from_reader(reader)?);
                        }
                        b"r" => {
                            let run = Run::from_reader(reader, &e)?;
                            para.content.push(ParagraphContent::Run(run));
                        }
                        b"hyperlink" => {
                            let link = Hyperlink::from_reader(reader, &e)?;
                            para.content.push(ParagraphContent::Hyperlink(link));
                        }
                        b"bookmarkStart" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            let name = crate::xml::get_attr(&e, "w:name")
                                .or_else(|| crate::xml::get_attr(&e, "name"))
                                .unwrap_or_default();
                            para.content.push(ParagraphContent::BookmarkStart { id, name });
                            // bookmarkStart is typically empty, but read until end just in case
                            skip_to_end(reader, &e)?;
                        }
                        b"bookmarkEnd" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            para.content.push(ParagraphContent::BookmarkEnd { id });
                            skip_to_end(reader, &e)?;
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            para.content.push(ParagraphContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::Empty(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"r" => {
                            let run = Run::from_empty(&e)?;
                            para.content.push(ParagraphContent::Run(run));
                        }
                        b"bookmarkStart" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            let name = crate::xml::get_attr(&e, "w:name")
                                .or_else(|| crate::xml::get_attr(&e, "name"))
                                .unwrap_or_default();
                            para.content.push(ParagraphContent::BookmarkStart { id, name });
                        }
                        b"bookmarkEnd" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            para.content.push(ParagraphContent::BookmarkEnd { id });
                        }
                        _ => {
                            // Unknown empty element - preserve
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
                            para.content.push(ParagraphContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"p" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(para)
    }

    /// Create from empty element
    pub fn from_empty(start: &BytesStart) -> Result<Self> {
        let mut para = Paragraph::default();

        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            para.unknown_attrs.push((key, value));
        }

        Ok(para)
    }

    /// Get all text in this paragraph
    pub fn text(&self) -> String {
        let mut result = String::new();
        for content in &self.content {
            match content {
                ParagraphContent::Run(run) => {
                    result.push_str(&run.text());
                }
                ParagraphContent::Hyperlink(link) => {
                    for run in &link.runs {
                        result.push_str(&run.text());
                    }
                }
                _ => {}
            }
        }
        result
    }

    /// Get style ID
    pub fn style(&self) -> Option<&str> {
        self.properties.as_ref()?.style.as_deref()
    }

    /// Get all runs
    pub fn runs(&self) -> impl Iterator<Item = &Run> {
        self.content.iter().filter_map(|c| {
            if let ParagraphContent::Run(r) = c {
                Some(r)
            } else {
                None
            }
        })
    }

    /// Check if this is a heading (has outline level or heading style)
    pub fn is_heading(&self) -> bool {
        if let Some(ref props) = self.properties {
            if props.outline_level.is_some() {
                return true;
            }
            if let Some(ref style) = props.style {
                return style.starts_with("Heading") || style.starts_with("heading");
            }
        }
        false
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:p");
        for (key, value) in &self.unknown_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        // Check if paragraph is completely empty
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

            writer.write_event(Event::End(BytesEnd::new("w:p")))?;
        }

        Ok(())
    }

    /// Create a new paragraph with text
    pub fn new(text: impl Into<String>) -> Self {
        Paragraph {
            content: vec![ParagraphContent::Run(Run::new(text))],
            ..Default::default()
        }
    }

    /// Add a run to this paragraph
    pub fn add_run(&mut self, run: Run) {
        self.content.push(ParagraphContent::Run(run));
    }

    /// Set style
    pub fn set_style(&mut self, style: impl Into<String>) {
        self.properties
            .get_or_insert_with(Default::default)
            .style = Some(style.into());
    }
}

impl ParagraphContent {
    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            ParagraphContent::Run(run) => run.write_to(writer),
            ParagraphContent::Hyperlink(link) => link.write_to(writer),
            ParagraphContent::BookmarkStart { id, name } => {
                let mut elem = BytesStart::new("w:bookmarkStart");
                elem.push_attribute(("w:id", id.as_str()));
                elem.push_attribute(("w:name", name.as_str()));
                writer.write_event(Event::Empty(elem))?;
                Ok(())
            }
            ParagraphContent::BookmarkEnd { id } => {
                let mut elem = BytesStart::new("w:bookmarkEnd");
                elem.push_attribute(("w:id", id.as_str()));
                writer.write_event(Event::Empty(elem))?;
                Ok(())
            }
            ParagraphContent::Unknown(node) => node.write_to(writer),
        }
    }
}

impl ParagraphProperties {
    /// Parse from reader (after w:pPr start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = ParagraphProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"numPr" => {
                            // Parse numbering properties
                            parse_num_pr(reader, &mut props)?;
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
                        b"pStyle" => {
                            props.style = get_w_val(&e);
                        }
                        b"jc" => {
                            props.justification = get_w_val(&e);
                        }
                        b"outlineLvl" => {
                            props.outline_level = get_w_val(&e).and_then(|v| v.parse().ok());
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

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // Check if there are any properties to write
        let has_content = self.style.is_some()
            || self.justification.is_some()
            || self.num_id.is_some()
            || self.outline_level.is_some()
            || !self.unknown_children.is_empty();

        if !has_content {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        // Style
        if let Some(style) = &self.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Numbering
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

        // Justification
        if let Some(jc) = &self.justification {
            let mut elem = BytesStart::new("w:jc");
            elem.push_attribute(("w:val", jc.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Outline level
        if let Some(level) = self.outline_level {
            let mut elem = BytesStart::new("w:outlineLvl");
            elem.push_attribute(("w:val", level.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        Ok(())
    }
}

impl Hyperlink {
    /// Parse from reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut link = Hyperlink::default();

        // Get r:id or anchor
        link.r_id = crate::xml::get_attr(start, "r:id");
        link.anchor = crate::xml::get_attr(start, "w:anchor")
            .or_else(|| crate::xml::get_attr(start, "anchor"));

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    if e.name().local_name().as_ref() == b"r" {
                        let run = Run::from_reader(reader, &e)?;
                        link.runs.push(run);
                    } else {
                        // Skip unknown
                        skip_to_end(reader, &e)?;
                    }
                }
                Event::Empty(e) => {
                    if e.name().local_name().as_ref() == b"r" {
                        let run = Run::from_empty(&e)?;
                        link.runs.push(run);
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"hyperlink" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(link)
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:hyperlink");
        if let Some(r_id) = &self.r_id {
            start.push_attribute(("r:id", r_id.as_str()));
        }
        if let Some(anchor) = &self.anchor {
            start.push_attribute(("w:anchor", anchor.as_str()));
        }

        if self.runs.is_empty() {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            for run in &self.runs {
                run.write_to(writer)?;
            }
            writer.write_event(Event::End(BytesEnd::new("w:hyperlink")))?;
        }

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
                    b"numId" => {
                        props.num_id = get_w_val(&e).and_then(|v| v.parse().ok());
                    }
                    b"ilvl" => {
                        props.num_level = get_w_val(&e).and_then(|v| v.parse().ok());
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                if e.name().local_name().as_ref() == b"numPr" {
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

/// Skip to end of current element
fn skip_to_end<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<()> {
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
