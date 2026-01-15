//! Relationships handling for OPC packages
//!
//! Parses and generates `.rels` files

use crate::error::{Error, Result};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::{BufRead, Write};

/// Collection of relationships
#[derive(Clone, Debug)]
pub struct Relationships {
    /// Relationships by ID
    items: HashMap<String, Relationship>,
    /// Next auto-generated ID number
    next_id: u32,
}

impl Default for Relationships {
    fn default() -> Self {
        Self {
            items: HashMap::new(),
            next_id: 1, // Start from 1, not 0
        }
    }
}

/// A single relationship
#[derive(Clone, Debug)]
pub struct Relationship {
    /// Relationship ID (e.g., "rId1")
    pub id: String,
    /// Relationship type URI
    pub rel_type: String,
    /// Target path (relative or absolute)
    pub target: String,
    /// Target mode
    pub target_mode: TargetMode,
}

/// Target mode for relationships
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TargetMode {
    /// Internal target (part within the package)
    #[default]
    Internal,
    /// External target (hyperlink, etc.)
    External,
}

impl Relationships {
    /// Create empty relationships
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse from XML string
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        Self::from_reader(&mut reader)
    }

    /// Parse from a reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut rels = Self::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) | Event::Start(e) => {
                    let name = e.name();
                    if name.local_name().as_ref() == b"Relationship" {
                        let rel = parse_relationship(&e)?;
                        rels.items.insert(rel.id.clone(), rel);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        rels.update_next_id();
        Ok(rels)
    }

    /// Serialize to XML string
    pub fn to_xml(&self) -> String {
        let mut buf = Vec::new();
        self.write_to(&mut buf).expect("write to Vec should not fail");
        String::from_utf8(buf).expect("XML should be valid UTF-8")
    }

    /// Write to a writer
    pub fn write_to<W: Write>(&self, writer: W) -> Result<()> {
        let mut xml = Writer::new(writer);

        // XML declaration
        xml.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        // Relationships element
        let mut rels_elem = BytesStart::new("Relationships");
        rels_elem.push_attribute(("xmlns", NS_RELATIONSHIPS));
        xml.write_event(Event::Start(rels_elem))?;

        // Relationship elements
        for rel in self.items.values() {
            let mut rel_elem = BytesStart::new("Relationship");
            rel_elem.push_attribute(("Id", rel.id.as_str()));
            rel_elem.push_attribute(("Type", rel.rel_type.as_str()));
            rel_elem.push_attribute(("Target", rel.target.as_str()));

            if rel.target_mode == TargetMode::External {
                rel_elem.push_attribute(("TargetMode", "External"));
            }

            xml.write_event(Event::Empty(rel_elem))?;
        }

        xml.write_event(Event::End(BytesEnd::new("Relationships")))?;

        Ok(())
    }

    /// Get a relationship by ID
    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.items.get(id)
    }

    /// Get a relationship by type (returns first match)
    pub fn by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.items.values().find(|r| r.rel_type == rel_type)
    }

    /// Get all relationships of a given type
    pub fn all_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.items
            .values()
            .filter(|r| r.rel_type == rel_type)
            .collect()
    }

    /// Add a relationship (auto-generates ID)
    pub fn add(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.generate_id();
        self.add_with_id(&id, rel_type, target, TargetMode::Internal);
        id
    }

    /// Add an external relationship
    pub fn add_external(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.generate_id();
        self.add_with_id(&id, rel_type, target, TargetMode::External);
        id
    }

    /// Add a relationship with a specific ID
    pub fn add_with_id(&mut self, id: &str, rel_type: &str, target: &str, mode: TargetMode) {
        let rel = Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: mode,
        };
        self.items.insert(id.to_string(), rel);
    }

    /// Remove a relationship by ID
    pub fn remove(&mut self, id: &str) -> Option<Relationship> {
        self.items.remove(id)
    }

    /// Iterate over all relationships
    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.items.values()
    }

    /// Number of relationships
    pub fn len(&self) -> usize {
        self.items.len()
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// Generate a new unique ID
    fn generate_id(&mut self) -> String {
        let id = format!("rId{}", self.next_id);
        self.next_id += 1;
        id
    }

    /// Update next_id based on existing relationships
    fn update_next_id(&mut self) {
        let max_id = self
            .items
            .keys()
            .filter_map(|id| {
                if id.starts_with("rId") {
                    id[3..].parse::<u32>().ok()
                } else {
                    None
                }
            })
            .max()
            .unwrap_or(0);

        self.next_id = max_id + 1;
    }
}

/// Parse a single Relationship element
fn parse_relationship(element: &BytesStart) -> Result<Relationship> {
    let mut id = None;
    let mut rel_type = None;
    let mut target = None;
    let mut target_mode = TargetMode::Internal;

    for attr in element.attributes() {
        let attr = attr?;
        let key = attr.key.local_name();
        let value = String::from_utf8_lossy(&attr.value).to_string();

        match key.as_ref() {
            b"Id" => id = Some(value),
            b"Type" => rel_type = Some(value),
            b"Target" => target = Some(value),
            b"TargetMode" => {
                if value == "External" {
                    target_mode = TargetMode::External;
                }
            }
            _ => {}
        }
    }

    Ok(Relationship {
        id: id.ok_or_else(|| Error::MissingAttribute {
            element: "Relationship".into(),
            attr: "Id".into(),
        })?,
        rel_type: rel_type.ok_or_else(|| Error::MissingAttribute {
            element: "Relationship".into(),
            attr: "Type".into(),
        })?,
        target: target.ok_or_else(|| Error::MissingAttribute {
            element: "Relationship".into(),
            attr: "Target".into(),
        })?,
        target_mode,
    })
}

// Namespace
const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

// Well-known relationship types
pub mod rel_types {
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_relationships() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let rels = Relationships::from_xml(xml).unwrap();

        assert_eq!(rels.len(), 2);

        let r1 = rels.get("rId1").unwrap();
        assert_eq!(r1.target, "word/document.xml");
        assert_eq!(r1.target_mode, TargetMode::Internal);

        let r2 = rels.get("rId2").unwrap();
        assert_eq!(r2.target, "https://example.com");
        assert_eq!(r2.target_mode, TargetMode::External);
    }

    #[test]
    fn test_by_type() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

        let rels = Relationships::from_xml(xml).unwrap();
        let doc = rels.by_type(rel_types::OFFICE_DOCUMENT).unwrap();
        assert_eq!(doc.target, "word/document.xml");
    }

    #[test]
    fn test_roundtrip() {
        let mut rels = Relationships::new();
        rels.add(rel_types::STYLES, "styles.xml");
        rels.add_external(rel_types::HYPERLINK, "https://example.com");

        let xml = rels.to_xml();
        let rels2 = Relationships::from_xml(&xml).unwrap();

        assert_eq!(rels2.len(), 2);
        assert!(rels2.by_type(rel_types::STYLES).is_some());
        assert!(rels2.by_type(rel_types::HYPERLINK).is_some());
    }

    #[test]
    fn test_auto_id() {
        let mut rels = Relationships::new();
        let id1 = rels.add(rel_types::STYLES, "styles.xml");
        let id2 = rels.add(rel_types::NUMBERING, "numbering.xml");

        assert_eq!(id1, "rId1");
        assert_eq!(id2, "rId2");
    }
}
