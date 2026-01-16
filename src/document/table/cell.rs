//! Table cell elements (w:tc, w:tcPr)

use crate::document::Paragraph;
use crate::error::Result;
use crate::xml::RawXmlElement;
use crate::xml::RawXmlNode;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

use super::types::{VMerge, VerticalAlignment};

/// Table cell (w:tc)
#[derive(Clone, Debug, Default)]
pub struct TableCell {
    /// Cell properties
    pub properties: Option<TableCellProperties>,
    /// Cell content (paragraphs)
    pub paragraphs: Vec<Paragraph>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

/// Table cell properties
#[derive(Clone, Debug, Default)]
pub struct TableCellProperties {
    /// Cell width
    pub width: Option<i32>,
    /// Grid span (horizontal merge)
    pub grid_span: Option<u32>,
    /// Vertical merge
    pub v_merge: Option<VMerge>,
    /// Vertical alignment
    pub v_align: Option<String>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl TableCell {
    /// Create a new cell with text
    pub fn new(text: impl Into<String>) -> Self {
        let text = text.into();
        let paragraphs = if text.is_empty() {
            vec![Paragraph::default()]
        } else {
            vec![Paragraph::new(text)]
        };
        TableCell {
            paragraphs,
            ..Default::default()
        }
    }

    /// Set the cell text (replaces all paragraphs with a single one)
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.paragraphs.clear();
        self.paragraphs.push(Paragraph::new(text));
    }

    /// Add a paragraph to the cell
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.paragraphs.push(para);
    }

    /// Set cell width (in twips)
    pub fn set_width(&mut self, width: i32) {
        self.properties.get_or_insert_with(Default::default).width = Some(width);
    }

    /// Set horizontal merge (grid span)
    pub fn set_grid_span(&mut self, span: u32) {
        self.properties
            .get_or_insert_with(Default::default)
            .grid_span = Some(span);
    }

    /// Set vertical merge
    pub fn set_v_merge(&mut self, v_merge: VMerge) {
        self.properties
            .get_or_insert_with(Default::default)
            .v_merge = Some(v_merge);
    }

    /// Set vertical alignment
    pub fn set_v_align(&mut self, align: impl Into<String>) {
        self.properties
            .get_or_insert_with(Default::default)
            .v_align = Some(align.into());
    }

    /// Parse from reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, _start: &BytesStart) -> Result<Self> {
        let mut cell = TableCell::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();

                    match local.as_ref() {
                        b"tcPr" => {
                            cell.properties = Some(TableCellProperties::from_reader(reader)?);
                        }
                        b"p" => {
                            let para = Paragraph::from_reader(reader, &e)?;
                            cell.paragraphs.push(para);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            cell.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"p" {
                        let para = Paragraph::from_empty(&e)?;
                        cell.paragraphs.push(para);
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
                        cell.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"tc" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(cell)
    }

    /// Get cell text (all paragraphs concatenated)
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Iterate over paragraphs
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.paragraphs.iter()
    }

    /// Get mutable paragraphs iterator
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut Paragraph> {
        self.paragraphs.iter_mut()
    }

    /// Get cell width in twips
    pub fn width(&self) -> Option<i32> {
        self.properties.as_ref()?.width
    }

    /// Get vertical alignment
    pub fn vertical_alignment(&self) -> Option<VerticalAlignment> {
        self.properties
            .as_ref()?
            .v_align
            .as_ref()
            .map(|s| VerticalAlignment::parse(s))
    }

    /// Get grid span (horizontal merge count)
    pub fn grid_span(&self) -> Option<u32> {
        self.properties.as_ref()?.grid_span
    }

    /// Get vertical merge status
    pub fn v_merge(&self) -> Option<&VMerge> {
        self.properties.as_ref()?.v_merge.as_ref()
    }

    /// Check if this cell is the start of a horizontal merge
    pub fn is_merge_start(&self) -> bool {
        self.grid_span().map(|s| s > 1).unwrap_or(false)
    }

    /// Check if this cell is the start of a vertical merge
    pub fn is_v_merge_start(&self) -> bool {
        matches!(self.v_merge(), Some(VMerge::Restart))
    }

    /// Check if this cell continues a vertical merge
    pub fn is_v_merge_continue(&self) -> bool {
        matches!(self.v_merge(), Some(VMerge::Continue))
    }

    /// Clear cell content
    pub fn clear(&mut self) {
        self.paragraphs.clear();
        self.paragraphs.push(Paragraph::default());
    }

    /// Set vertical alignment using VerticalAlignment enum
    pub fn set_vertical_alignment(&mut self, align: VerticalAlignment) {
        self.properties
            .get_or_insert_with(Default::default)
            .v_align = Some(align.as_str().to_string());
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tc")))?;

        // Cell properties
        if let Some(props) = &self.properties {
            props.write_to(writer)?;
        }

        // Paragraphs (at least one required in cell)
        if self.paragraphs.is_empty() {
            // Write empty paragraph if none exist
            writer.write_event(Event::Empty(BytesStart::new("w:p")))?;
        } else {
            for para in &self.paragraphs {
                para.write_to(writer)?;
            }
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tc")))?;
        Ok(())
    }
}

impl TableCellProperties {
    /// Parse from reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut props = TableCellProperties::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let raw = RawXmlElement::from_reader(reader, &e)?;
                    props.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();

                    match local.as_ref() {
                        b"tcW" => {
                            props.width = crate::xml::get_attr(&e, "w:w")
                                .or_else(|| crate::xml::get_attr(&e, "w"))
                                .and_then(|v| v.parse().ok());
                        }
                        b"gridSpan" => {
                            props.grid_span =
                                crate::xml::get_w_val(&e).and_then(|v| v.parse().ok());
                        }
                        b"vMerge" => {
                            let val = crate::xml::get_w_val(&e);
                            props.v_merge = Some(match val.as_deref() {
                                Some("restart") => VMerge::Restart,
                                _ => VMerge::Continue,
                            });
                        }
                        b"vAlign" => {
                            props.v_align = crate::xml::get_w_val(&e);
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
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"tcPr" {
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
        let has_content = self.width.is_some()
            || self.grid_span.is_some()
            || self.v_merge.is_some()
            || self.v_align.is_some()
            || !self.unknown_children.is_empty();

        if !has_content {
            return Ok(());
        }

        writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;

        // Width
        if let Some(width) = self.width {
            let mut elem = BytesStart::new("w:tcW");
            elem.push_attribute(("w:w", width.to_string().as_str()));
            elem.push_attribute(("w:type", "dxa"));
            writer.write_event(Event::Empty(elem))?;
        }

        // Grid span
        if let Some(span) = self.grid_span {
            let mut elem = BytesStart::new("w:gridSpan");
            elem.push_attribute(("w:val", span.to_string().as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Vertical merge
        if let Some(v_merge) = &self.v_merge {
            let mut elem = BytesStart::new("w:vMerge");
            match v_merge {
                VMerge::Restart => elem.push_attribute(("w:val", "restart")),
                VMerge::Continue => {}
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Vertical alignment
        if let Some(v_align) = &self.v_align {
            let mut elem = BytesStart::new("w:vAlign");
            elem.push_attribute(("w:val", v_align.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;
        Ok(())
    }
}
