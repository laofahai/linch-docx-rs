//! Content Types handling for OPC packages
//!
//! Parses and generates `[Content_Types].xml`

use crate::error::{Error, Result};
use crate::opc::PartUri;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::{BufRead, Write};

/// Content types definition for an OPC package
#[derive(Clone, Debug, Default)]
pub struct ContentTypes {
    /// Default extension mappings (extension -> content type)
    defaults: HashMap<String, String>,
    /// Override mappings (part URI -> content type)
    overrides: HashMap<PartUri, String>,
}

impl ContentTypes {
    /// Create a new ContentTypes with standard defaults
    pub fn new() -> Self {
        let mut ct = Self::default();

        // Standard defaults
        ct.add_default("rels", RELATIONSHIPS);
        ct.add_default("xml", XML);

        // Common image types
        ct.add_default("png", "image/png");
        ct.add_default("jpeg", "image/jpeg");
        ct.add_default("jpg", "image/jpeg");
        ct.add_default("gif", "image/gif");
        ct.add_default("bmp", "image/bmp");

        ct
    }

    /// Parse from XML string
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        Self::from_reader(&mut reader)
    }

    /// Parse from a reader
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Self> {
        let mut ct = Self::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) => {
                    let name = e.name();
                    let local_name = name.local_name();
                    let local_name_ref = local_name.as_ref();

                    match local_name_ref {
                        b"Default" => {
                            let ext = get_attr(&e, "Extension")?;
                            let content_type = get_attr(&e, "ContentType")?;
                            ct.defaults.insert(ext.to_lowercase(), content_type);
                        }
                        b"Override" => {
                            let part_name = get_attr(&e, "PartName")?;
                            let content_type = get_attr(&e, "ContentType")?;
                            let uri = PartUri::new(&part_name)?;
                            ct.overrides.insert(uri, content_type);
                        }
                        _ => {}
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(ct)
    }

    /// Serialize to XML string
    pub fn to_xml(&self) -> String {
        let mut buf = Vec::new();
        self.write_to(&mut buf)
            .expect("write to Vec should not fail");
        String::from_utf8(buf).expect("XML should be valid UTF-8")
    }

    /// Write to a writer
    pub fn write_to<W: Write>(&self, writer: W) -> Result<()> {
        let mut xml = Writer::new(writer);

        // XML declaration
        xml.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Types element
        let mut types = BytesStart::new("Types");
        types.push_attribute(("xmlns", NS_CONTENT_TYPES));
        xml.write_event(Event::Start(types))?;

        // Default elements
        for (ext, content_type) in &self.defaults {
            let mut default = BytesStart::new("Default");
            default.push_attribute(("Extension", ext.as_str()));
            default.push_attribute(("ContentType", content_type.as_str()));
            xml.write_event(Event::Empty(default))?;
        }

        // Override elements
        for (uri, content_type) in &self.overrides {
            let mut override_elem = BytesStart::new("Override");
            override_elem.push_attribute(("PartName", uri.as_str()));
            override_elem.push_attribute(("ContentType", content_type.as_str()));
            xml.write_event(Event::Empty(override_elem))?;
        }

        xml.write_event(Event::End(BytesEnd::new("Types")))?;

        Ok(())
    }

    /// Add a default extension mapping
    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .insert(extension.to_lowercase(), content_type.to_string());
    }

    /// Add an override for a specific part
    pub fn add_override(&mut self, uri: &PartUri, content_type: &str) {
        self.overrides.insert(uri.clone(), content_type.to_string());
    }

    /// Get the content type for a part
    pub fn get(&self, uri: &PartUri) -> Option<&str> {
        // Check overrides first
        if let Some(ct) = self.overrides.get(uri) {
            return Some(ct);
        }

        // Fall back to extension default
        uri.extension()
            .and_then(|ext| self.defaults.get(&ext.to_lowercase()))
            .map(|s| s.as_str())
    }

    /// Remove an override
    pub fn remove_override(&mut self, uri: &PartUri) -> Option<String> {
        self.overrides.remove(uri)
    }
}

/// Get an attribute value from an XML element
fn get_attr(element: &BytesStart, name: &str) -> Result<String> {
    for attr in element.attributes() {
        let attr = attr?;
        if attr.key.local_name().as_ref() == name.as_bytes() {
            return Ok(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    Err(Error::MissingAttribute {
        element: String::from_utf8_lossy(element.name().as_ref()).to_string(),
        attr: name.to_string(),
    })
}

// Namespace
const NS_CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

// Well-known content types
pub const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
pub const XML: &str = "application/xml";
pub const MAIN_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_content_types() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

        let ct = ContentTypes::from_xml(xml).unwrap();

        assert_eq!(ct.defaults.get("rels"), Some(&RELATIONSHIPS.to_string()));
        assert_eq!(ct.defaults.get("xml"), Some(&XML.to_string()));

        let doc_uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(ct.get(&doc_uri), Some(MAIN_DOCUMENT));
    }

    #[test]
    fn test_roundtrip() {
        let mut ct = ContentTypes::new();
        ct.add_override(&PartUri::new("/word/document.xml").unwrap(), MAIN_DOCUMENT);

        let xml = ct.to_xml();
        let ct2 = ContentTypes::from_xml(&xml).unwrap();

        let doc_uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(ct2.get(&doc_uri), Some(MAIN_DOCUMENT));
    }

    #[test]
    fn test_get_by_extension() {
        let ct = ContentTypes::new();
        let uri = PartUri::new("/word/media/image1.png").unwrap();
        assert_eq!(ct.get(&uri), Some("image/png"));
    }
}
