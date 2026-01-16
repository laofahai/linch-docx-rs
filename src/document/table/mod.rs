//! Table elements (w:tbl, w:tr, w:tc)

mod builder;
mod cell;
mod row;
mod types;

pub use builder::TableBuilder;
pub use cell::{TableCell, TableCellProperties};
pub use row::TableRow;
pub use types::{GridColumn, TableAlignment, TableWidth, VMerge, VerticalAlignment};

use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Table element (w:tbl)
#[derive(Clone, Debug, Default)]
pub struct Table {
    /// Table properties
    pub properties: Option<RawXmlNode>,
    /// Table grid
    pub grid: Vec<GridColumn>,
    /// Table rows
    pub rows: Vec<TableRow>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Table {
    /// Create a new table with the specified number of rows and columns
    pub fn new(rows: usize, cols: usize) -> Self {
        let mut table_rows = Vec::with_capacity(rows);
        for _ in 0..rows {
            let mut cells = Vec::with_capacity(cols);
            for _ in 0..cols {
                cells.push(TableCell::new(""));
            }
            table_rows.push(TableRow {
                cells,
                ..Default::default()
            });
        }

        let grid = (0..cols).map(|_| GridColumn { width: None }).collect();

        Table {
            grid,
            rows: table_rows,
            ..Default::default()
        }
    }

    /// Create a table from a 2D array of strings
    pub fn from_data<S: Into<String> + Clone>(data: &[&[S]]) -> Self {
        let rows: Vec<TableRow> = data
            .iter()
            .map(|row| {
                let cells: Vec<TableCell> = row.iter().map(|text| TableCell::new(text.clone())).collect();
                TableRow {
                    cells,
                    ..Default::default()
                }
            })
            .collect();

        let cols = data.first().map(|r| r.len()).unwrap_or(0);
        let grid = (0..cols).map(|_| GridColumn { width: None }).collect();

        Table {
            grid,
            rows,
            ..Default::default()
        }
    }

    /// Parse from reader (after w:tbl start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, _start: &BytesStart) -> Result<Self> {
        let mut table = Table::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();

                    match local.as_ref() {
                        b"tblPr" => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            table.properties = Some(RawXmlNode::Element(raw));
                        }
                        b"tblGrid" => {
                            table.grid = parse_table_grid(reader)?;
                        }
                        b"tr" => {
                            let row = TableRow::from_reader(reader, &e)?;
                            table.rows.push(row);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            table.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    // Handle empty elements
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
                    table.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"tbl" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(table)
    }

    /// Get row count
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get column count (based on first row)
    pub fn column_count(&self) -> usize {
        self.rows.first().map(|r| r.cells.len()).unwrap_or(0)
    }

    /// Get cell at position
    pub fn cell(&self, row: usize, col: usize) -> Option<&TableCell> {
        self.rows.get(row)?.cells.get(col)
    }

    /// Iterate over rows
    pub fn rows(&self) -> impl Iterator<Item = &TableRow> {
        self.rows.iter()
    }

    /// Get row by index
    pub fn row(&self, index: usize) -> Option<&TableRow> {
        self.rows.get(index)
    }

    /// Get mutable cell at position
    pub fn cell_mut(&mut self, row: usize, col: usize) -> Option<&mut TableCell> {
        self.rows.get_mut(row)?.cells.get_mut(col)
    }

    /// Get mutable row
    pub fn row_mut(&mut self, index: usize) -> Option<&mut TableRow> {
        self.rows.get_mut(index)
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.rows.push(row);
    }

    /// Add a new empty row with the same column count as the table
    pub fn add_empty_row(&mut self) -> &mut TableRow {
        let cols = self.column_count();
        self.rows.push(TableRow::new(cols));
        self.rows.last_mut().expect("just added")
    }

    /// Insert a row at the specified index
    pub fn insert_row(&mut self, index: usize, row: TableRow) {
        if index <= self.rows.len() {
            self.rows.insert(index, row);
        }
    }

    /// Remove a row at the specified index
    pub fn remove_row(&mut self, index: usize) -> Option<TableRow> {
        if index < self.rows.len() {
            Some(self.rows.remove(index))
        } else {
            None
        }
    }

    /// Add a column to the table (adds an empty cell to each row)
    pub fn add_column(&mut self) {
        self.grid.push(GridColumn { width: None });
        for row in &mut self.rows {
            row.add_cell(TableCell::new(""));
        }
    }

    /// Insert a column at the specified index
    pub fn insert_column(&mut self, index: usize) {
        if index <= self.grid.len() {
            self.grid.insert(index, GridColumn { width: None });
            for row in &mut self.rows {
                row.insert_cell(index, TableCell::new(""));
            }
        }
    }

    /// Remove a column at the specified index
    pub fn remove_column(&mut self, index: usize) -> bool {
        if index < self.grid.len() {
            self.grid.remove(index);
            for row in &mut self.rows {
                row.remove_cell(index);
            }
            true
        } else {
            false
        }
    }

    /// Set cell text at position
    pub fn set_cell_text(&mut self, row: usize, col: usize, text: impl Into<String>) {
        if let Some(cell) = self.cell_mut(row, col) {
            cell.set_text(text);
        }
    }

    /// Set column width (in twips, 1/20 of a point)
    pub fn set_column_width(&mut self, col: usize, width: i32) {
        if col < self.grid.len() {
            self.grid[col].width = Some(width);
        }
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;

        // Table properties
        if let Some(props) = &self.properties {
            props.write_to(writer)?;
        }

        // Table grid
        if !self.grid.is_empty() {
            writer.write_event(Event::Start(BytesStart::new("w:tblGrid")))?;
            for col in &self.grid {
                let mut elem = BytesStart::new("w:gridCol");
                if let Some(w) = col.width {
                    elem.push_attribute(("w:w", w.to_string().as_str()));
                }
                writer.write_event(Event::Empty(elem))?;
            }
            writer.write_event(Event::End(BytesEnd::new("w:tblGrid")))?;
        }

        // Rows
        for row in &self.rows {
            row.write_to(writer)?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
        Ok(())
    }
}

/// Parse table grid
fn parse_table_grid<R: BufRead>(reader: &mut Reader<R>) -> Result<Vec<GridColumn>> {
    let mut columns = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) => {
                if e.name().local_name().as_ref() == b"gridCol" {
                    let width = crate::xml::get_attr(&e, "w:w")
                        .or_else(|| crate::xml::get_attr(&e, "w"))
                        .and_then(|v| v.parse().ok());
                    columns.push(GridColumn { width });
                }
            }
            Event::End(e) => {
                if e.name().local_name().as_ref() == b"tblGrid" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(columns)
}
