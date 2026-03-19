//! Template engine for DOCX documents
//!
//! Supports `{{placeholder}}` syntax for text replacement in paragraphs,
//! headers, footers, and table cells.

use crate::document::{BlockContent, Document, ParagraphContent, RunContent};
use std::collections::HashMap;

/// Template context: a map of placeholder names to replacement values
pub type TemplateContext = HashMap<String, String>;

impl Document {
    /// Fill template placeholders in the entire document.
    ///
    /// Replaces all occurrences of `{{key}}` with the corresponding value
    /// from the context map. Works across paragraphs, tables, headers,
    /// footers, footnotes, and endnotes.
    ///
    /// Returns the total number of replacements made.
    ///
    /// # Example
    /// ```rust,ignore
    /// use linch_docx_rs::Document;
    /// use std::collections::HashMap;
    ///
    /// let mut doc = Document::open("template.docx")?;
    /// let mut ctx = HashMap::new();
    /// ctx.insert("name".into(), "Alice".into());
    /// ctx.insert("date".into(), "2024-01-15".into());
    /// let count = doc.fill_template(&ctx);
    /// doc.save("filled.docx")?;
    /// ```
    pub fn fill_template(&mut self, context: &TemplateContext) -> usize {
        let mut count = 0;

        // Body content
        for block in &mut self.body.content {
            match block {
                BlockContent::Paragraph(para) => {
                    count += fill_paragraph_runs(para, context);
                }
                BlockContent::Table(table) => {
                    for row in &mut table.rows {
                        for cell in &mut row.cells {
                            for para in &mut cell.paragraphs {
                                count += fill_paragraph_runs(para, context);
                            }
                        }
                    }
                }
                BlockContent::Unknown(_) => {}
            }
        }

        // Headers
        for (_, hf) in &mut self.headers {
            for para in &mut hf.paragraphs {
                count += fill_paragraph_runs(para, context);
            }
        }

        // Footers
        for (_, hf) in &mut self.footers {
            for para in &mut hf.paragraphs {
                count += fill_paragraph_runs(para, context);
            }
        }

        // Footnotes
        if let Some(ref mut notes) = self.footnotes {
            for note in &mut notes.notes {
                for para in &mut note.paragraphs {
                    count += fill_paragraph_runs(para, context);
                }
            }
        }

        // Endnotes
        if let Some(ref mut notes) = self.endnotes {
            for note in &mut notes.notes {
                for para in &mut note.paragraphs {
                    count += fill_paragraph_runs(para, context);
                }
            }
        }

        count
    }

    /// Get all placeholder names found in the document.
    ///
    /// Scans all text content for `{{...}}` patterns and returns
    /// the unique placeholder names found.
    pub fn template_placeholders(&self) -> Vec<String> {
        let mut placeholders = Vec::new();
        let full_text = self.text();
        extract_placeholders(&full_text, &mut placeholders);

        // Also check tables
        for table in self.tables() {
            for row_idx in 0..table.row_count() {
                if let Some(row) = table.row(row_idx) {
                    for cell_idx in 0..row.cell_count() {
                        if let Some(cell) = row.cell(cell_idx) {
                            extract_placeholders(&cell.text(), &mut placeholders);
                        }
                    }
                }
            }
        }

        placeholders.sort();
        placeholders.dedup();
        placeholders
    }
}

/// Replace `{{key}}` placeholders in a paragraph's runs
fn fill_paragraph_runs(para: &mut crate::document::Paragraph, context: &TemplateContext) -> usize {
    let mut count = 0;

    // First, try simple per-run replacement
    for content in &mut para.content {
        if let ParagraphContent::Run(run) = content {
            for rc in &mut run.content {
                if let RunContent::Text(ref mut text) = rc {
                    for (key, value) in context {
                        let placeholder = format!("{{{{{}}}}}", key);
                        let matches = text.matches(&placeholder).count();
                        if matches > 0 {
                            *text = text.replace(&placeholder, value);
                            count += matches;
                        }
                    }
                }
            }
        }
    }

    // Handle cross-run placeholders: when {{ and }} span multiple runs.
    // Merge all text, do replacements, then check if anything changed.
    if count == 0 {
        let full_text = para.text();
        let mut has_placeholder = false;
        for key in context.keys() {
            let placeholder = format!("{{{{{}}}}}", key);
            if full_text.contains(&placeholder) {
                has_placeholder = true;
                break;
            }
        }

        if has_placeholder {
            // Merge all runs into one, replace, set back
            let mut merged = full_text;
            for (key, value) in context {
                let placeholder = format!("{{{{{}}}}}", key);
                let matches = merged.matches(&placeholder).count();
                if matches > 0 {
                    merged = merged.replace(&placeholder, value);
                    count += matches;
                }
            }
            if count > 0 {
                // Replace content with single run preserving first run's properties
                let first_props = para.content.iter().find_map(|c| {
                    if let ParagraphContent::Run(r) = c {
                        r.properties.clone()
                    } else {
                        None
                    }
                });

                // Remove all runs, keep non-run content
                para.content
                    .retain(|c| !matches!(c, ParagraphContent::Run(_)));

                let mut new_run = crate::document::Run::new(merged);
                new_run.properties = first_props;
                para.content.insert(0, ParagraphContent::Run(new_run));
            }
        }
    }

    count
}

/// Extract placeholder names from text
fn extract_placeholders(text: &str, out: &mut Vec<String>) {
    let mut start = 0;
    while let Some(open) = text[start..].find("{{") {
        let abs_open = start + open + 2;
        if let Some(close) = text[abs_open..].find("}}") {
            let name = text[abs_open..abs_open + close].trim().to_string();
            if !name.is_empty() && !name.contains('{') && !name.contains('}') {
                out.push(name);
            }
            start = abs_open + close + 2;
        } else {
            break;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::{Document, Run, Table};

    #[test]
    fn test_fill_template_simple() {
        let mut doc = Document::new();
        doc.add_paragraph("Hello, {{name}}! Today is {{date}}.");

        let mut ctx = TemplateContext::new();
        ctx.insert("name".into(), "Alice".into());
        ctx.insert("date".into(), "Monday".into());

        let count = doc.fill_template(&ctx);
        assert_eq!(count, 2);
        assert_eq!(
            doc.paragraph(0).unwrap().text(),
            "Hello, Alice! Today is Monday."
        );
    }

    #[test]
    fn test_fill_template_multiple_paragraphs() {
        let mut doc = Document::new();
        doc.add_paragraph("Dear {{name}},");
        doc.add_paragraph("Your order #{{order_id}} is ready.");

        let mut ctx = TemplateContext::new();
        ctx.insert("name".into(), "Bob".into());
        ctx.insert("order_id".into(), "12345".into());

        let count = doc.fill_template(&ctx);
        assert_eq!(count, 2);
        assert!(doc.text().contains("Dear Bob,"));
        assert!(doc.text().contains("order #12345"));
    }

    #[test]
    fn test_fill_template_in_table() {
        let mut doc = Document::new();
        let table = Table::new(1, 2);
        let t = doc.add_table(table);
        t.set_cell_text(0, 0, "{{col1}}");
        t.set_cell_text(0, 1, "{{col2}}");

        let mut ctx = TemplateContext::new();
        ctx.insert("col1".into(), "Value A".into());
        ctx.insert("col2".into(), "Value B".into());

        let count = doc.fill_template(&ctx);
        assert_eq!(count, 2);
    }

    #[test]
    fn test_template_placeholders() {
        let mut doc = Document::new();
        doc.add_paragraph("{{name}} lives in {{city}}");
        doc.add_paragraph("Age: {{age}}");

        let placeholders = doc.template_placeholders();
        assert!(placeholders.contains(&"name".to_string()));
        assert!(placeholders.contains(&"city".to_string()));
        assert!(placeholders.contains(&"age".to_string()));
        assert_eq!(placeholders.len(), 3);
    }

    #[test]
    fn test_fill_template_no_match() {
        let mut doc = Document::new();
        doc.add_paragraph("No placeholders here.");

        let ctx = TemplateContext::new();
        let count = doc.fill_template(&ctx);
        assert_eq!(count, 0);
    }

    #[test]
    fn test_fill_template_preserves_formatting() {
        let mut doc = Document::new();
        let p = doc.add_empty_paragraph();
        let mut run = Run::new("Hello {{name}}!");
        run.set_bold(true);
        p.add_run(run);

        let mut ctx = TemplateContext::new();
        ctx.insert("name".into(), "World".into());

        doc.fill_template(&ctx);
        let r = doc.paragraph(0).unwrap().runs().next().unwrap();
        assert_eq!(r.text(), "Hello World!");
        assert!(r.bold());
    }

    #[test]
    fn test_fill_template_roundtrip() {
        let mut doc = Document::new();
        doc.add_paragraph("{{greeting}}, {{name}}!");

        let mut ctx = TemplateContext::new();
        ctx.insert("greeting".into(), "你好".into());
        ctx.insert("name".into(), "世界".into());

        doc.fill_template(&ctx);

        let bytes = doc.to_bytes().unwrap();
        let doc2 = Document::from_bytes(&bytes).unwrap();
        assert_eq!(doc2.paragraph(0).unwrap().text(), "你好, 世界!");
    }
}
