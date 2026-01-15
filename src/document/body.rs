//! Document body and block-level content

use crate::document::{Paragraph, Table};
use crate::error::Result;
use crate::xml::RawXmlNode;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Block-level content in a document body
#[derive(Clone, Debug)]
pub enum BlockContent {
    /// Paragraph
    Paragraph(Paragraph),
    /// Table
    Table(Table),
    /// Unknown element (preserved for round-trip)
    Unknown(RawXmlNode),
}

/// Document body (w:body)
#[derive(Clone, Debug, Default)]
pub struct Body {
    /// Block-level content
    pub content: Vec<BlockContent>,
    /// Section properties (last sectPr in body)
    pub section_properties: Option<RawXmlNode>,
}

impl Body {
    /// Parse body from XML reader (after w:body start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut body = Body::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"p" => {
                            let para = Paragraph::from_reader(reader, &e)?;
                            body.content.push(BlockContent::Paragraph(para));
                        }
                        b"tbl" => {
                            let table = Table::from_reader(reader, &e)?;
                            body.content.push(BlockContent::Table(table));
                        }
                        b"sectPr" => {
                            // Section properties - preserve raw
                            let raw = crate::xml::RawXmlElement::from_reader(reader, &e)?;
                            body.section_properties = Some(RawXmlNode::Element(raw));
                        }
                        _ => {
                            // Unknown element - preserve for round-trip
                            let raw = crate::xml::RawXmlElement::from_reader(reader, &e)?;
                            body.content.push(BlockContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::Empty(e) => {
                    let name = e.name();
                    let local = name.local_name();

                    match local.as_ref() {
                        b"p" => {
                            // Empty paragraph
                            let para = Paragraph::from_empty(&e)?;
                            body.content.push(BlockContent::Paragraph(para));
                        }
                        _ => {
                            // Preserve unknown empty elements
                            let raw = crate::xml::RawXmlElement {
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
                            body.content.push(BlockContent::Unknown(RawXmlNode::Element(raw)));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"body" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(body)
    }

    /// Get all paragraphs
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.content.iter().filter_map(|c| {
            if let BlockContent::Paragraph(p) = c {
                Some(p)
            } else {
                None
            }
        })
    }

    /// Get all paragraphs mutably
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut Paragraph> {
        self.content.iter_mut().filter_map(|c| {
            if let BlockContent::Paragraph(p) = c {
                Some(p)
            } else {
                None
            }
        })
    }

    /// Get all tables
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.content.iter().filter_map(|c| {
            if let BlockContent::Table(t) = c {
                Some(t)
            } else {
                None
            }
        })
    }

    /// Write body to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:body")))?;

        // Write content
        for content in &self.content {
            content.write_to(writer)?;
        }

        // Write section properties
        if let Some(sect_pr) = &self.section_properties {
            sect_pr.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:body")))?;
        Ok(())
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.content.push(BlockContent::Paragraph(para));
    }

    /// Add a table
    pub fn add_table(&mut self, table: Table) {
        self.content.push(BlockContent::Table(table));
    }
}

impl BlockContent {
    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            BlockContent::Paragraph(para) => para.write_to(writer),
            BlockContent::Table(table) => table.write_to(writer),
            BlockContent::Unknown(node) => node.write_to(writer),
        }
    }
}
