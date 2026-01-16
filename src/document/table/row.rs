//! Table row elements (w:tr)

use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

use super::cell::TableCell;

/// Table row (w:tr)
#[derive(Clone, Debug, Default)]
pub struct TableRow {
    /// Row properties
    pub properties: Option<RawXmlNode>,
    /// Cells
    pub cells: Vec<TableCell>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl TableRow {
    /// Create a new row with empty cells
    pub fn new(cell_count: usize) -> Self {
        let cells = (0..cell_count).map(|_| TableCell::new("")).collect();
        TableRow {
            cells,
            ..Default::default()
        }
    }

    /// Create a row from cell texts
    pub fn from_texts<S: Into<String>>(texts: impl IntoIterator<Item = S>) -> Self {
        let cells = texts.into_iter().map(TableCell::new).collect();
        TableRow {
            cells,
            ..Default::default()
        }
    }

    /// Parse from reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, _start: &BytesStart) -> Result<Self> {
        let mut row = TableRow::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();

                    match local.as_ref() {
                        b"trPr" => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            row.properties = Some(RawXmlNode::Element(raw));
                        }
                        b"tc" => {
                            let cell = TableCell::from_reader(reader, &e)?;
                            row.cells.push(cell);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            row.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
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
                    row.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"tr" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(row)
    }

    /// Get cell count
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Iterate over cells
    pub fn cells(&self) -> impl Iterator<Item = &TableCell> {
        self.cells.iter()
    }

    /// Get cell at index
    pub fn cell(&self, index: usize) -> Option<&TableCell> {
        self.cells.get(index)
    }

    /// Get mutable cell at index
    pub fn cell_mut(&mut self, index: usize) -> Option<&mut TableCell> {
        self.cells.get_mut(index)
    }

    /// Add a cell to the row
    pub fn add_cell(&mut self, cell: TableCell) {
        self.cells.push(cell);
    }

    /// Insert a cell at the specified index
    pub fn insert_cell(&mut self, index: usize, cell: TableCell) {
        if index <= self.cells.len() {
            self.cells.insert(index, cell);
        }
    }

    /// Remove a cell at the specified index
    pub fn remove_cell(&mut self, index: usize) -> Option<TableCell> {
        if index < self.cells.len() {
            Some(self.cells.remove(index))
        } else {
            None
        }
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;

        // Row properties
        if let Some(props) = &self.properties {
            props.write_to(writer)?;
        }

        // Cells
        for cell in &self.cells {
            cell.write_to(writer)?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
        Ok(())
    }
}
