//! Table elements (w:tbl, w:tr, w:tc)

use crate::document::Paragraph;
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

/// Grid column definition
#[derive(Clone, Debug, Default)]
pub struct GridColumn {
    /// Width in twips
    pub width: Option<i32>,
}

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

/// Vertical merge type
#[derive(Clone, Debug)]
pub enum VMerge {
    Restart,
    Continue,
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
                let cells: Vec<TableCell> = row
                    .iter()
                    .map(|text| TableCell::new(text.clone()))
                    .collect();
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

    /// Get mutable cell at index
    pub fn cell_mut(&mut self, index: usize) -> Option<&mut TableCell> {
        self.cells.get_mut(index)
    }

    /// Add a cell to the row
    pub fn add_cell(&mut self, cell: TableCell) {
        self.cells.push(cell);
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
        self.properties.get_or_insert_with(Default::default).v_merge = Some(v_merge);
    }

    /// Set vertical alignment
    pub fn set_v_align(&mut self, align: impl Into<String>) {
        self.properties.get_or_insert_with(Default::default).v_align = Some(align.into());
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
