//! Document model - high-level API for DOCX documents

mod body;
mod paragraph;
mod run;
mod table;

pub use body::{BlockContent, Body};
pub use paragraph::{Hyperlink, Paragraph, ParagraphContent, ParagraphProperties};
pub use run::{BreakType, Run, RunContent, RunProperties};
pub use table::{GridColumn, Table, TableCell, TableCellProperties, TableRow, VMerge};

use crate::error::{Error, Result};
use crate::opc::{Package, Part, PartUri};
use crate::xml;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Cursor};
use std::path::Path;

/// A DOCX document
#[derive(Debug)]
pub struct Document {
    /// Underlying OPC package
    package: Package,
    /// Parsed document body
    body: Body,
}

impl Document {
    /// Open a document from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = Package::open(path)?;
        Self::from_package(package)
    }

    /// Open a document from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    /// Create document from an OPC package
    fn from_package(package: Package) -> Result<Self> {
        // Get main document part
        let doc_part = package
            .main_document_part()
            .ok_or_else(|| Error::MissingPart("Main document part not found".into()))?;

        // Parse document.xml
        let xml = doc_part.data_as_str()?;
        let body = parse_document_xml(xml)?;

        Ok(Self { package, body })
    }

    /// Create a new empty document
    pub fn new() -> Self {
        Self {
            package: Package::new(),
            body: Body::default(),
        }
    }

    /// Save the document to a file
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.update_package()?;
        self.package.save(path)
    }

    /// Save the document to bytes
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.update_package()?;
        self.package.to_bytes()
    }

    /// Update the package with current body content
    fn update_package(&mut self) -> Result<()> {
        let xml = serialize_document_xml(&self.body)?;
        let uri = PartUri::new("/word/document.xml")?;

        // Update or add the document part
        let part = Part::new(
            uri.clone(),
            crate::opc::MAIN_DOCUMENT.to_string(),
            xml.into_bytes(),
        );
        self.package.add_part(part);

        // Ensure the relationship exists for the main document
        if self.package.main_document_part().is_none() {
            use crate::opc::rel_types;
            self.package
                .add_relationship(rel_types::OFFICE_DOCUMENT, uri.as_str());
        }

        Ok(())
    }

    /// Get all paragraphs
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.body.paragraphs()
    }

    /// Get paragraph count
    pub fn paragraph_count(&self) -> usize {
        self.body
            .content
            .iter()
            .filter(|c| matches!(c, BlockContent::Paragraph(_)))
            .count()
    }

    /// Get paragraph by index
    pub fn paragraph(&self, index: usize) -> Option<&Paragraph> {
        self.body.paragraphs().nth(index)
    }

    /// Get all tables
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.body.tables()
    }

    /// Get table count
    pub fn table_count(&self) -> usize {
        self.body
            .content
            .iter()
            .filter(|c| matches!(c, BlockContent::Table(_)))
            .count()
    }

    /// Get table by index
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.body.tables().nth(index)
    }

    /// Get all text in the document
    pub fn text(&self) -> String {
        self.body
            .paragraphs()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get the underlying package
    pub fn package(&self) -> &Package {
        &self.package
    }

    /// Get mutable body
    pub fn body_mut(&mut self) -> &mut Body {
        &mut self.body
    }

    /// Add a paragraph with text
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        let para = Paragraph::new(text);
        self.body.add_paragraph(para);
        // Return mutable reference to the last paragraph
        self.body
            .content
            .iter_mut()
            .rev()
            .find_map(|c| {
                if let BlockContent::Paragraph(p) = c {
                    Some(p)
                } else {
                    None
                }
            })
            .expect("Just added paragraph")
    }

    /// Add an empty paragraph
    pub fn add_empty_paragraph(&mut self) -> &mut Paragraph {
        self.body.add_paragraph(Paragraph::default());
        self.body
            .content
            .iter_mut()
            .rev()
            .find_map(|c| {
                if let BlockContent::Paragraph(p) = c {
                    Some(p)
                } else {
                    None
                }
            })
            .expect("Just added paragraph")
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

/// Parse document.xml content
fn parse_document_xml(xml: &str) -> Result<Body> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut body = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let name = e.name();
                let local = name.local_name();

                match local.as_ref() {
                    b"body" => {
                        body = Some(Body::from_reader(&mut reader)?);
                    }
                    b"document" => {
                        // Continue to find body
                    }
                    _ => {
                        // Skip unknown elements at document level
                        skip_element(&mut reader, &e)?;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    body.ok_or_else(|| Error::InvalidDocument("Missing w:body element".into()))
}

/// Serialize body to document.xml content
fn serialize_document_xml(body: &Body) -> Result<String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut writer = Writer::new(&mut buffer);

    // XML declaration
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Document element with namespaces
    let mut doc_start = BytesStart::new("w:document");
    for (attr, value) in xml::document_namespaces() {
        doc_start.push_attribute((attr, value));
    }
    writer.write_event(Event::Start(doc_start))?;

    // Body
    body.write_to(&mut writer)?;

    // Close document
    writer.write_event(Event::End(BytesEnd::new("w:document")))?;

    let xml_bytes = buffer.into_inner();
    String::from_utf8(xml_bytes).map_err(|e| Error::InvalidDocument(e.to_string()))
}

/// Skip an element and all its children
fn skip_element<R: BufRead>(
    reader: &mut Reader<R>,
    start: &quick_xml::events::BytesStart,
) -> Result<()> {
    let target = start.name().as_ref().to_vec();
    let mut depth = 1;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == target => depth += 1,
            Event::End(e) if e.name().as_ref() == target => {
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

    const SIMPLE_DOC: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>This is a heading</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    #[test]
    fn test_parse_simple_document() {
        let body = parse_document_xml(SIMPLE_DOC).unwrap();

        // Should have 2 paragraphs
        let paras: Vec<_> = body.paragraphs().collect();
        assert_eq!(paras.len(), 2);

        // First paragraph
        assert_eq!(paras[0].text(), "Hello, World!");

        // Second paragraph
        assert_eq!(paras[1].text(), "This is a heading");
        assert_eq!(paras[1].style(), Some("Heading1"));

        // Check run properties
        let runs: Vec<_> = paras[1].runs().collect();
        assert_eq!(runs.len(), 1);
        assert!(runs[0].bold());
    }

    #[test]
    fn test_parse_with_formatting() {
        let xml = r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:sz w:val="28"/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>Formatted text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document_xml(xml).unwrap();
        let para = body.paragraphs().next().unwrap();
        let run = para.runs().next().unwrap();

        assert!(run.bold());
        assert!(run.italic());
        assert_eq!(run.font_size_pt(), Some(14.0)); // 28 half-points = 14pt
        assert_eq!(run.color(), Some("FF0000"));
    }
}
