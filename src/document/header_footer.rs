//! Header and Footer elements

use crate::document::Paragraph;
use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

/// A header or footer part
#[derive(Clone, Debug, Default)]
pub struct HeaderFooter {
    /// Paragraphs in the header/footer
    pub paragraphs: Vec<Paragraph>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
    /// Whether this is a header (true) or footer (false)
    pub is_header: bool,
}

impl HeaderFooter {
    /// Parse from XML string
    pub fn from_xml(xml: &str, is_header: bool) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut hf = HeaderFooter {
            is_header,
            ..Default::default()
        };
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"hdr" | b"ftr" => {
                            // Root element, continue
                        }
                        b"p" => {
                            let para = Paragraph::from_reader(&mut reader, &e)?;
                            hf.paragraphs.push(para);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            hf.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"p" {
                        let para = Paragraph::from_empty(&e)?;
                        hf.paragraphs.push(para);
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
                        hf.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(hf)
    }

    /// Serialize to XML string
    pub fn to_xml(&self) -> Result<String> {
        let mut buffer = Cursor::new(Vec::new());
        let mut writer = Writer::new(&mut buffer);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let tag = if self.is_header { "w:hdr" } else { "w:ftr" };
        let mut start = BytesStart::new(tag);
        start.push_attribute(("xmlns:w", crate::xml::W));
        start.push_attribute(("xmlns:r", crate::xml::R));
        writer.write_event(Event::Start(start))?;

        for para in &self.paragraphs {
            para.write_to(&mut writer)?;
        }

        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new(tag)))?;

        let xml_bytes = buffer.into_inner();
        String::from_utf8(xml_bytes)
            .map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get all text
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, text: impl Into<String>) {
        self.paragraphs.push(Paragraph::new(text));
    }

    /// Create a new empty header
    pub fn new_header() -> Self {
        HeaderFooter {
            is_header: true,
            paragraphs: vec![Paragraph::default()],
            ..Default::default()
        }
    }

    /// Create a new empty footer
    pub fn new_footer() -> Self {
        HeaderFooter {
            is_header: false,
            paragraphs: vec![Paragraph::default()],
            ..Default::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_header() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>Header Text</w:t>
    </w:r>
  </w:p>
</w:hdr>"#;

        let hf = HeaderFooter::from_xml(xml, true).unwrap();
        assert!(hf.is_header);
        assert_eq!(hf.paragraphs.len(), 1);
        assert_eq!(hf.text(), "Header Text");
    }

    #[test]
    fn test_header_roundtrip() {
        let mut hf = HeaderFooter::new_header();
        hf.paragraphs.clear();
        hf.add_paragraph("Test Header");

        let xml = hf.to_xml().unwrap();
        let hf2 = HeaderFooter::from_xml(&xml, true).unwrap();
        assert_eq!(hf2.text(), "Test Header");
    }

    #[test]
    fn test_footer() {
        let mut hf = HeaderFooter::new_footer();
        hf.paragraphs.clear();
        hf.add_paragraph("Page 1");

        let xml = hf.to_xml().unwrap();
        assert!(xml.contains("w:ftr"));
        assert!(!xml.contains("w:hdr"));
    }
}
