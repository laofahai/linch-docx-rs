//! Paragraph element (w:p)

mod properties;

pub use properties::{Alignment, Indentation, LineSpacing, ParagraphProperties};

use crate::document::numbering::NumberingInfo;
use crate::document::Run;
use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
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

impl Paragraph {
    /// Parse paragraph from reader (after w:p start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut para = Paragraph::default();

        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            para.unknown_attrs.push((key, value));
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"pPr" => {
                            para.properties = Some(ParagraphProperties::from_reader(reader)?);
                        }
                        b"r" => {
                            para.content
                                .push(ParagraphContent::Run(Run::from_reader(reader, &e)?));
                        }
                        b"hyperlink" => {
                            para.content
                                .push(ParagraphContent::Hyperlink(Hyperlink::from_reader(
                                    reader, &e,
                                )?));
                        }
                        b"bookmarkStart" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            let name = crate::xml::get_attr(&e, "w:name")
                                .or_else(|| crate::xml::get_attr(&e, "name"))
                                .unwrap_or_default();
                            para.content
                                .push(ParagraphContent::BookmarkStart { id, name });
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
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            para.content
                                .push(ParagraphContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"r" => {
                            para.content
                                .push(ParagraphContent::Run(Run::from_empty(&e)?));
                        }
                        b"bookmarkStart" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            let name = crate::xml::get_attr(&e, "w:name")
                                .or_else(|| crate::xml::get_attr(&e, "name"))
                                .unwrap_or_default();
                            para.content
                                .push(ParagraphContent::BookmarkStart { id, name });
                        }
                        b"bookmarkEnd" => {
                            let id = crate::xml::get_attr(&e, "w:id")
                                .or_else(|| crate::xml::get_attr(&e, "id"))
                                .unwrap_or_default();
                            para.content.push(ParagraphContent::BookmarkEnd { id });
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
                            para.content
                                .push(ParagraphContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::End(e) if e.name().local_name().as_ref() == b"p" => break,
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
                ParagraphContent::Run(run) => result.push_str(&run.text()),
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

    /// Check if this is a heading
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

    /// Get numbering information if this paragraph is a list item
    pub fn numbering(&self) -> Option<NumberingInfo> {
        let props = self.properties.as_ref()?;
        let num_id = props.num_id?;
        Some(NumberingInfo::new(num_id, props.num_level.unwrap_or(0)))
    }

    /// Check if this paragraph is a list item
    pub fn is_list_item(&self) -> bool {
        self.properties.as_ref().and_then(|p| p.num_id).is_some()
    }

    /// Get the list level (0-8) if this is a list item
    pub fn list_level(&self) -> Option<u32> {
        let props = self.properties.as_ref()?;
        if props.num_id.is_some() {
            Some(props.num_level.unwrap_or(0))
        } else {
            None
        }
    }

    /// Set numbering for this paragraph
    pub fn set_numbering(&mut self, num_id: u32, level: u32) {
        let props = self.properties.get_or_insert_with(Default::default);
        props.num_id = Some(num_id);
        props.num_level = Some(level);
    }

    /// Remove numbering from this paragraph
    pub fn clear_numbering(&mut self) {
        if let Some(ref mut props) = self.properties {
            props.num_id = None;
            props.num_level = None;
        }
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:p");
        for (key, value) in &self.unknown_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        let is_empty = self.properties.is_none()
            && self.content.is_empty()
            && self.unknown_children.is_empty();

        if is_empty {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            if let Some(props) = &self.properties {
                props.write_to(writer)?;
            }
            for content in &self.content {
                content.write_to(writer)?;
            }
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
        self.properties.get_or_insert_with(Default::default).style = Some(style.into());
    }

    /// Get alignment
    pub fn alignment(&self) -> Option<Alignment> {
        self.properties
            .as_ref()?
            .justification
            .as_deref()
            .and_then(Alignment::from_ooxml)
    }

    /// Set alignment
    pub fn set_alignment(&mut self, alignment: Alignment) {
        self.properties
            .get_or_insert_with(Default::default)
            .justification = Some(alignment.as_ooxml().to_string());
    }

    /// Set indentation
    pub fn set_indentation(&mut self, indentation: Indentation) {
        self.properties
            .get_or_insert_with(Default::default)
            .indentation = Some(indentation);
    }

    /// Set line spacing
    pub fn set_spacing(&mut self, spacing: LineSpacing) {
        self.properties.get_or_insert_with(Default::default).spacing = Some(spacing);
    }

    /// Set keep with next
    pub fn set_keep_next(&mut self, keep: bool) {
        self.properties
            .get_or_insert_with(Default::default)
            .keep_next = Some(keep);
    }

    /// Set page break before
    pub fn set_page_break_before(&mut self, page_break: bool) {
        self.properties
            .get_or_insert_with(Default::default)
            .page_break_before = Some(page_break);
    }

    /// Get mutable runs
    pub fn runs_mut(&mut self) -> impl Iterator<Item = &mut Run> {
        self.content.iter_mut().filter_map(|c| {
            if let ParagraphContent::Run(r) = c {
                Some(r)
            } else {
                None
            }
        })
    }

    /// Clear all content
    pub fn clear(&mut self) {
        self.content.clear();
    }

    /// Set text (replaces all content with a single run)
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.content.clear();
        self.content.push(ParagraphContent::Run(Run::new(text)));
    }

    /// Check if paragraph is empty
    pub fn is_empty(&self) -> bool {
        self.content.is_empty()
    }

    /// Get heading level (1-9) or None
    pub fn heading_level(&self) -> Option<u8> {
        if let Some(ref props) = self.properties {
            if let Some(level) = props.outline_level {
                return Some(level + 1);
            }
            if let Some(ref style) = props.style {
                if let Some(suffix) = style.strip_prefix("Heading") {
                    return suffix.parse().ok();
                }
            }
        }
        None
    }

    /// Get run count
    pub fn run_count(&self) -> usize {
        self.content
            .iter()
            .filter(|c| matches!(c, ParagraphContent::Run(_)))
            .count()
    }

    /// Add a hyperlink with a relationship ID (for external links)
    pub fn add_hyperlink(&mut self, r_id: impl Into<String>, text: impl Into<String>) {
        let link = Hyperlink {
            r_id: Some(r_id.into()),
            anchor: None,
            runs: vec![Run::new(text)],
        };
        self.content.push(ParagraphContent::Hyperlink(link));
    }

    /// Add an internal hyperlink (bookmark anchor)
    pub fn add_internal_link(&mut self, anchor: impl Into<String>, text: impl Into<String>) {
        let link = Hyperlink {
            r_id: None,
            anchor: Some(anchor.into()),
            runs: vec![Run::new(text)],
        };
        self.content.push(ParagraphContent::Hyperlink(link));
    }

    /// Add a bookmark
    pub fn add_bookmark(&mut self, id: impl Into<String>, name: impl Into<String>) {
        let id = id.into();
        self.content.push(ParagraphContent::BookmarkStart {
            id: id.clone(),
            name: name.into(),
        });
        self.content.push(ParagraphContent::BookmarkEnd { id });
    }

    /// Remove a run by index
    pub fn remove_run(&mut self, index: usize) -> bool {
        let mut run_idx = 0;
        for i in 0..self.content.len() {
            if matches!(self.content[i], ParagraphContent::Run(_)) {
                if run_idx == index {
                    self.content.remove(i);
                    return true;
                }
                run_idx += 1;
            }
        }
        false
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

impl Hyperlink {
    /// Parse from reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut link = Hyperlink {
            r_id: crate::xml::get_attr(start, "r:id"),
            anchor: crate::xml::get_attr(start, "w:anchor")
                .or_else(|| crate::xml::get_attr(start, "anchor")),
            ..Default::default()
        };

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    if e.name().local_name().as_ref() == b"r" {
                        link.runs.push(Run::from_reader(reader, &e)?);
                    } else {
                        skip_to_end(reader, &e)?;
                    }
                }
                Event::Empty(e) => {
                    if e.name().local_name().as_ref() == b"r" {
                        link.runs.push(Run::from_empty(&e)?);
                    }
                }
                Event::End(e) if e.name().local_name().as_ref() == b"hyperlink" => break,
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
