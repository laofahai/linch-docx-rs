//! Table builder for fluent table construction

use super::cell::TableCell;
use super::row::TableRow;
use super::types::{GridColumn, TableAlignment, TableWidth};
use super::Table;

/// Builder for creating tables with a fluent API
pub struct TableBuilder {
    rows: usize,
    cols: usize,
    width: Option<TableWidth>,
    alignment: Option<TableAlignment>,
    data: Option<Vec<Vec<String>>>,
    column_widths: Vec<Option<i32>>,
}

impl TableBuilder {
    /// Create a new table builder with specified dimensions
    pub fn new(rows: usize, cols: usize) -> Self {
        TableBuilder {
            rows,
            cols,
            width: None,
            alignment: None,
            data: None,
            column_widths: vec![None; cols],
        }
    }

    /// Set table width
    pub fn width(mut self, width: TableWidth) -> Self {
        self.width = Some(width);
        self
    }

    /// Set table alignment
    pub fn alignment(mut self, alignment: TableAlignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    /// Set column widths (in twips)
    pub fn column_widths(mut self, widths: &[i32]) -> Self {
        for (i, &w) in widths.iter().enumerate() {
            if i < self.column_widths.len() {
                self.column_widths[i] = Some(w);
            }
        }
        self
    }

    /// Set data from 2D string slice
    pub fn data<S: Into<String> + Clone>(mut self, data: &[&[S]]) -> Self {
        self.data = Some(
            data.iter()
                .map(|row| row.iter().map(|s| s.clone().into()).collect())
                .collect(),
        );
        // Update rows/cols if data is provided
        if let Some(ref d) = self.data {
            self.rows = d.len();
            self.cols = d.first().map(|r| r.len()).unwrap_or(0);
            self.column_widths.resize(self.cols, None);
        }
        self
    }

    /// Build the table
    pub fn build(self) -> Table {
        let table = if let Some(data) = self.data {
            let rows: Vec<TableRow> = data
                .into_iter()
                .map(|row| {
                    let cells: Vec<TableCell> = row.into_iter().map(TableCell::new).collect();
                    TableRow {
                        cells,
                        ..Default::default()
                    }
                })
                .collect();

            let grid: Vec<GridColumn> = self
                .column_widths
                .into_iter()
                .map(|w| GridColumn { width: w })
                .collect();

            Table {
                grid,
                rows,
                ..Default::default()
            }
        } else {
            let mut t = Table::new(self.rows, self.cols);
            for (i, width) in self.column_widths.into_iter().enumerate() {
                if let Some(w) = width {
                    t.set_column_width(i, w);
                }
            }
            t
        };

        // Apply width and alignment properties would require modifying table properties
        // For now, we store them but don't apply to XML (would need tblPr element construction)
        let _ = self.width;
        let _ = self.alignment;

        table
    }
}

impl Table {
    /// Create a table builder with specified dimensions
    pub fn builder(rows: usize, cols: usize) -> TableBuilder {
        TableBuilder::new(rows, cols)
    }
}
