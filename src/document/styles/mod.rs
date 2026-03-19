//! Styles module - parsing and managing styles.xml

mod types;

pub use types::{DocDefaults, Style, StyleType};

use crate::error::Result;
use crate::xml::{get_attr, get_w_val, parse_bool, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Cursor};

use super::{ParagraphProperties, RunProperties};

/// Collection of styles from styles.xml
#[derive(Clone, Debug, Default)]
pub struct Styles {
    /// Document defaults
    pub doc_defaults: Option<DocDefaults>,
    /// Style definitions
    pub styles: Vec<Style>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Styles {
    /// Parse from XML string
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut styles = Styles::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"styles" => {
                            // Root element, continue parsing children
                        }
                        b"docDefaults" => {
                            styles.doc_defaults = Some(parse_doc_defaults(&mut reader)?);
                        }
                        b"style" => {
                            styles.styles.push(parse_style(&mut reader, &e)?);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            styles.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"style" {
                        styles.styles.push(parse_style_from_empty(&e)?);
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
                        styles.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(styles)
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

        let mut start = BytesStart::new("w:styles");
        start.push_attribute(("xmlns:w", crate::xml::W));
        start.push_attribute(("xmlns:r", crate::xml::R));
        writer.write_event(Event::Start(start))?;

        // Write doc defaults
        if let Some(ref defaults) = self.doc_defaults {
            write_doc_defaults(&mut writer, defaults)?;
        }

        // Write styles
        for style in &self.styles {
            write_style(&mut writer, style)?;
        }

        // Write unknown children
        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:styles")))?;

        let xml_bytes = buffer.into_inner();
        String::from_utf8(xml_bytes)
            .map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get style by ID
    pub fn get(&self, style_id: &str) -> Option<&Style> {
        self.styles.iter().find(|s| s.style_id == style_id)
    }

    /// Get mutable style by ID
    pub fn get_mut(&mut self, style_id: &str) -> Option<&mut Style> {
        self.styles.iter_mut().find(|s| s.style_id == style_id)
    }

    /// Get style by display name
    pub fn get_by_name(&self, name: &str) -> Option<&Style> {
        self.styles.iter().find(|s| s.name.as_deref() == Some(name))
    }

    /// Iterate all styles
    pub fn iter(&self) -> impl Iterator<Item = &Style> {
        self.styles.iter()
    }

    /// Iterate paragraph styles
    pub fn paragraph_styles(&self) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(|s| s.style_type == Some(StyleType::Paragraph))
    }

    /// Iterate character styles
    pub fn character_styles(&self) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(|s| s.style_type == Some(StyleType::Character))
    }

    /// Add a style
    pub fn add(&mut self, style: Style) {
        self.styles.push(style);
    }

    /// Remove a style by ID
    pub fn remove(&mut self, style_id: &str) -> Option<Style> {
        let idx = self.styles.iter().position(|s| s.style_id == style_id)?;
        Some(self.styles.remove(idx))
    }
}

// === Parsing helpers ===

fn parse_doc_defaults<R: BufRead>(reader: &mut Reader<R>) -> Result<DocDefaults> {
    let mut defaults = DocDefaults::default();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                match local.as_ref() {
                    b"rPrDefault" => {
                        defaults.run_properties = parse_default_rpr(reader)?;
                    }
                    b"pPrDefault" => {
                        defaults.paragraph_properties = parse_default_ppr(reader)?;
                    }
                    _ => {
                        let raw = RawXmlElement::from_reader(reader, &e)?;
                        defaults.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"docDefaults" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(defaults)
}

fn parse_default_rpr<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<RunProperties>> {
    let mut result = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"rPr" {
                    result = Some(RunProperties::from_reader(reader)?);
                } else {
                    skip_element(reader, &e)?;
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"rPrDefault" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

fn parse_default_ppr<R: BufRead>(reader: &mut Reader<R>) -> Result<Option<ParagraphProperties>> {
    let mut result = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"pPr" {
                    result = Some(ParagraphProperties::from_reader(reader)?);
                } else {
                    skip_element(reader, &e)?;
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"pPrDefault" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

fn parse_style<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Style> {
    let unknown_attrs: Vec<(String, String)> = start
        .attributes()
        .filter_map(|a| a.ok())
        .filter(|a| {
            let key = String::from_utf8_lossy(a.key.as_ref());
            !matches!(
                key.as_ref(),
                "w:type"
                    | "type"
                    | "w:styleId"
                    | "styleId"
                    | "w:default"
                    | "default"
                    | "w:customStyle"
                    | "customStyle"
            )
        })
        .map(|a| {
            (
                String::from_utf8_lossy(a.key.as_ref()).to_string(),
                String::from_utf8_lossy(&a.value).to_string(),
            )
        })
        .collect();

    let mut style = Style {
        style_type: get_attr(start, "w:type")
            .or_else(|| get_attr(start, "type"))
            .and_then(|v| StyleType::parse(&v)),
        style_id: get_attr(start, "w:styleId")
            .or_else(|| get_attr(start, "styleId"))
            .unwrap_or_default(),
        is_default: get_attr(start, "w:default")
            .or_else(|| get_attr(start, "default"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false),
        is_custom: get_attr(start, "w:customStyle")
            .or_else(|| get_attr(start, "customStyle"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false),
        unknown_attrs,
        ..Default::default()
    };

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                match local.as_ref() {
                    b"pPr" => {
                        style.paragraph_properties =
                            Some(ParagraphProperties::from_reader(reader)?);
                    }
                    b"rPr" => {
                        style.run_properties = Some(RunProperties::from_reader(reader)?);
                    }
                    _ => {
                        let raw = RawXmlElement::from_reader(reader, &e)?;
                        style.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
            }
            Event::Empty(e) => {
                let local = e.name().local_name();
                match local.as_ref() {
                    b"name" => {
                        style.name = get_w_val(&e);
                    }
                    b"basedOn" => {
                        style.based_on = get_w_val(&e);
                    }
                    b"next" => {
                        style.next_style = get_w_val(&e);
                    }
                    b"link" => {
                        style.link = get_w_val(&e);
                    }
                    b"uiPriority" => {
                        style.ui_priority = get_w_val(&e).and_then(|v| v.parse().ok());
                    }
                    b"semiHidden" => {
                        style.semi_hidden = parse_bool(&e);
                    }
                    b"unhideWhenUsed" => {
                        style.unhide_when_used = parse_bool(&e);
                    }
                    b"qFormat" => {
                        style.qformat = parse_bool(&e);
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
                        style.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"style" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(style)
}

fn parse_style_from_empty(start: &BytesStart) -> Result<Style> {
    Ok(Style {
        style_type: get_attr(start, "w:type")
            .or_else(|| get_attr(start, "type"))
            .and_then(|v| StyleType::parse(&v)),
        style_id: get_attr(start, "w:styleId")
            .or_else(|| get_attr(start, "styleId"))
            .unwrap_or_default(),
        is_default: get_attr(start, "w:default")
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false),
        ..Default::default()
    })
}

// === Writing helpers ===

fn write_doc_defaults<W: std::io::Write>(
    writer: &mut Writer<W>,
    defaults: &DocDefaults,
) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("w:docDefaults")))?;

    if let Some(ref rpr) = defaults.run_properties {
        writer.write_event(Event::Start(BytesStart::new("w:rPrDefault")))?;
        rpr.write_to(writer)?;
        writer.write_event(Event::End(BytesEnd::new("w:rPrDefault")))?;
    }

    if let Some(ref ppr) = defaults.paragraph_properties {
        writer.write_event(Event::Start(BytesStart::new("w:pPrDefault")))?;
        ppr.write_to(writer)?;
        writer.write_event(Event::End(BytesEnd::new("w:pPrDefault")))?;
    }

    for child in &defaults.unknown_children {
        child.write_to(writer)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:docDefaults")))?;
    Ok(())
}

fn write_style<W: std::io::Write>(writer: &mut Writer<W>, style: &Style) -> Result<()> {
    let mut start = BytesStart::new("w:style");

    if let Some(ref st) = style.style_type {
        start.push_attribute(("w:type", st.as_str()));
    }
    if !style.style_id.is_empty() {
        start.push_attribute(("w:styleId", style.style_id.as_str()));
    }
    if style.is_default {
        start.push_attribute(("w:default", "1"));
    }
    if style.is_custom {
        start.push_attribute(("w:customStyle", "1"));
    }
    for (key, value) in &style.unknown_attrs {
        start.push_attribute((key.as_str(), value.as_str()));
    }

    writer.write_event(Event::Start(start))?;

    // Name
    if let Some(ref name) = style.name {
        let mut elem = BytesStart::new("w:name");
        elem.push_attribute(("w:val", name.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    // Based on
    if let Some(ref based_on) = style.based_on {
        let mut elem = BytesStart::new("w:basedOn");
        elem.push_attribute(("w:val", based_on.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    // Next style
    if let Some(ref next) = style.next_style {
        let mut elem = BytesStart::new("w:next");
        elem.push_attribute(("w:val", next.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    // Link
    if let Some(ref link) = style.link {
        let mut elem = BytesStart::new("w:link");
        elem.push_attribute(("w:val", link.as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    // UI priority
    if let Some(priority) = style.ui_priority {
        let mut elem = BytesStart::new("w:uiPriority");
        elem.push_attribute(("w:val", priority.to_string().as_str()));
        writer.write_event(Event::Empty(elem))?;
    }

    if style.semi_hidden {
        writer.write_event(Event::Empty(BytesStart::new("w:semiHidden")))?;
    }

    if style.unhide_when_used {
        writer.write_event(Event::Empty(BytesStart::new("w:unhideWhenUsed")))?;
    }

    if style.qformat {
        writer.write_event(Event::Empty(BytesStart::new("w:qFormat")))?;
    }

    // Paragraph properties
    if let Some(ref ppr) = style.paragraph_properties {
        ppr.write_to(writer)?;
    }

    // Run properties
    if let Some(ref rpr) = style.run_properties {
        rpr.write_to(writer)?;
    }

    // Unknown children
    for child in &style.unknown_children {
        child.write_to(writer)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:style")))?;
    Ok(())
}

fn skip_element<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<()> {
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

    #[test]
    fn test_parse_styles() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
</w:styles>"#;

        let styles = Styles::from_xml(xml).unwrap();
        assert!(styles.doc_defaults.is_some());
        assert_eq!(styles.styles.len(), 3);

        let normal = styles.get("Normal").unwrap();
        assert_eq!(normal.name.as_deref(), Some("Normal"));
        assert!(normal.is_default);
        assert!(normal.qformat);

        let h1 = styles.get("Heading1").unwrap();
        assert_eq!(h1.based_on.as_deref(), Some("Normal"));
        assert_eq!(h1.next_style.as_deref(), Some("Normal"));
        assert!(h1.run_properties.is_some());
        assert!(h1.paragraph_properties.is_some());

        let strong = styles.get("Strong").unwrap();
        assert_eq!(strong.style_type, Some(StyleType::Character));

        // Test by_name
        let found = styles.get_by_name("heading 1").unwrap();
        assert_eq!(found.style_id, "Heading1");

        // Test iterators
        assert_eq!(styles.paragraph_styles().count(), 2);
        assert_eq!(styles.character_styles().count(), 1);
    }

    #[test]
    fn test_styles_roundtrip() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"#;

        let styles = Styles::from_xml(xml).unwrap();
        let output = styles.to_xml().unwrap();
        let styles2 = Styles::from_xml(&output).unwrap();

        assert_eq!(styles2.styles.len(), 1);
        assert_eq!(
            styles2.get("Normal").unwrap().name.as_deref(),
            Some("Normal")
        );
    }
}
