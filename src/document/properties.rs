//! Core properties (core.xml) - Dublin Core metadata

use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

/// Core properties from core.xml (Dublin Core metadata)
#[derive(Clone, Debug, Default)]
pub struct CoreProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
    pub category: Option<String>,
    pub content_status: Option<String>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

impl CoreProperties {
    /// Parse from XML string
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut props = CoreProperties::default();
        let mut buf = Vec::new();
        let mut current_element: Option<String> = None;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    let local_str = String::from_utf8_lossy(local.as_ref()).to_string();

                    match local_str.as_str() {
                        "coreProperties" => {
                            // Root element, continue
                        }
                        "title" | "subject" | "creator" | "description" | "keywords"
                        | "lastModifiedBy" | "revision" | "created" | "modified" | "category"
                        | "contentStatus" => {
                            current_element = Some(local_str);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            props.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Text(t) => {
                    if let Some(ref elem) = current_element {
                        let text = t.unescape()?.to_string();
                        match elem.as_str() {
                            "title" => props.title = Some(text),
                            "subject" => props.subject = Some(text),
                            "creator" => props.creator = Some(text),
                            "keywords" => props.keywords = Some(text),
                            "description" => props.description = Some(text),
                            "lastModifiedBy" => props.last_modified_by = Some(text),
                            "revision" => props.revision = Some(text),
                            "created" => props.created = Some(text),
                            "modified" => props.modified = Some(text),
                            "category" => props.category = Some(text),
                            "contentStatus" => props.content_status = Some(text),
                            _ => {}
                        }
                    }
                }
                Event::End(e) => {
                    let local = e.name().local_name();
                    let local_str = String::from_utf8_lossy(local.as_ref()).to_string();
                    if current_element.as_deref() == Some(&local_str) {
                        current_element = None;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(props)
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

        let mut start = BytesStart::new("cp:coreProperties");
        start.push_attribute(("xmlns:cp", crate::xml::CP));
        start.push_attribute(("xmlns:dc", crate::xml::DC));
        start.push_attribute(("xmlns:dcterms", crate::xml::DCTERMS));
        start.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        start.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        writer.write_event(Event::Start(start))?;

        write_dc_element(&mut writer, "dc:title", &self.title)?;
        write_dc_element(&mut writer, "dc:subject", &self.subject)?;
        write_dc_element(&mut writer, "dc:creator", &self.creator)?;
        write_cp_element(&mut writer, "cp:keywords", &self.keywords)?;
        write_dc_element(&mut writer, "dc:description", &self.description)?;
        write_cp_element(&mut writer, "cp:lastModifiedBy", &self.last_modified_by)?;
        write_cp_element(&mut writer, "cp:revision", &self.revision)?;
        write_datetime_element(&mut writer, "dcterms:created", &self.created)?;
        write_datetime_element(&mut writer, "dcterms:modified", &self.modified)?;
        write_cp_element(&mut writer, "cp:category", &self.category)?;
        write_cp_element(&mut writer, "cp:contentStatus", &self.content_status)?;

        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;

        let xml_bytes = buffer.into_inner();
        String::from_utf8(xml_bytes)
            .map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }
}

fn write_dc_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: &Option<String>,
) -> Result<()> {
    if let Some(ref v) = value {
        writer.write_event(Event::Start(BytesStart::new(name)))?;
        writer.write_event(Event::Text(BytesText::new(v)))?;
        writer.write_event(Event::End(BytesEnd::new(name)))?;
    }
    Ok(())
}

fn write_cp_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: &Option<String>,
) -> Result<()> {
    write_dc_element(writer, name, value)
}

fn write_datetime_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    name: &str,
    value: &Option<String>,
) -> Result<()> {
    if let Some(ref v) = value {
        let mut start = BytesStart::new(name);
        start.push_attribute(("xsi:type", "dcterms:W3CDTF"));
        writer.write_event(Event::Start(start))?;
        writer.write_event(Event::Text(BytesText::new(v)))?;
        writer.write_event(Event::End(BytesEnd::new(name)))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_core_properties() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <cp:revision>3</cp:revision>
  <dcterms:created>2024-01-15T10:30:00Z</dcterms:created>
</cp:coreProperties>"#;

        let props = CoreProperties::from_xml(xml).unwrap();
        assert_eq!(props.title.as_deref(), Some("Test Document"));
        assert_eq!(props.creator.as_deref(), Some("Test Author"));
        assert_eq!(props.revision.as_deref(), Some("3"));
        assert_eq!(props.created.as_deref(), Some("2024-01-15T10:30:00Z"));
        assert!(props.subject.is_none());
    }

    #[test]
    fn test_core_properties_roundtrip() {
        let props = CoreProperties {
            title: Some("My Doc".into()),
            creator: Some("Author".into()),
            modified: Some("2024-06-01T12:00:00Z".into()),
            ..Default::default()
        };

        let xml = props.to_xml().unwrap();
        let props2 = CoreProperties::from_xml(&xml).unwrap();

        assert_eq!(props2.title.as_deref(), Some("My Doc"));
        assert_eq!(props2.creator.as_deref(), Some("Author"));
        assert_eq!(props2.modified.as_deref(), Some("2024-06-01T12:00:00Z"));
    }
}
