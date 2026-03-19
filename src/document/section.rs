//! Section properties (w:sectPr)

use crate::error::Result;
use crate::xml::{get_attr, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

/// Section properties (w:sectPr)
#[derive(Clone, Debug, Default)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub page_margin: Option<PageMargin>,
    pub columns: Option<Columns>,
    pub header_references: Vec<HeaderFooterRef>,
    pub footer_references: Vec<HeaderFooterRef>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
    /// Unknown attributes (preserved for round-trip)
    pub unknown_attrs: Vec<(String, String)>,
}

/// Page size
#[derive(Clone, Debug, Default)]
pub struct PageSize {
    /// Width in twips
    pub width: Option<u32>,
    /// Height in twips
    pub height: Option<u32>,
    /// Orientation
    pub orient: Option<PageOrientation>,
}

/// Page orientation
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// Page margins (all values in twips)
#[derive(Clone, Debug, Default)]
pub struct PageMargin {
    pub top: Option<i32>,
    pub bottom: Option<i32>,
    pub left: Option<i32>,
    pub right: Option<i32>,
    pub header: Option<i32>,
    pub footer: Option<i32>,
    pub gutter: Option<i32>,
}

/// Column settings
#[derive(Clone, Debug, Default)]
pub struct Columns {
    pub count: Option<u32>,
    /// Space between columns in twips
    pub space: Option<u32>,
    pub equal_width: Option<bool>,
    pub unknown_children: Vec<RawXmlNode>,
}

/// Header or footer reference
#[derive(Clone, Debug)]
pub struct HeaderFooterRef {
    pub ref_type: HeaderFooterType,
    pub r_id: String,
}

/// Header/footer reference type
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HeaderFooterType {
    Default,
    First,
    Even,
}

impl SectionProperties {
    /// Parse from reader (after w:sectPr start tag)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let mut sect = SectionProperties::default();

        // Parse attributes
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            let value = String::from_utf8_lossy(&attr.value).to_string();
            sect.unknown_attrs.push((key, value));
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"cols" => {
                            sect.columns = Some(parse_columns(reader, &e)?);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            sect.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"pgSz" => {
                            sect.page_size = Some(parse_page_size(&e));
                        }
                        b"pgMar" => {
                            sect.page_margin = Some(parse_page_margin(&e));
                        }
                        b"headerReference" => {
                            if let Some(r) = parse_header_footer_ref(&e) {
                                sect.header_references.push(r);
                            }
                        }
                        b"footerReference" => {
                            if let Some(r) = parse_header_footer_ref(&e) {
                                sect.footer_references.push(r);
                            }
                        }
                        b"cols" => {
                            sect.columns = Some(parse_columns_empty(&e));
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
                            sect.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::End(e) if e.name().local_name().as_ref() == b"sectPr" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(sect)
    }

    /// Write to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:sectPr");
        for (key, value) in &self.unknown_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        writer.write_event(Event::Start(start))?;

        // Header references
        for href in &self.header_references {
            write_header_footer_ref(writer, "w:headerReference", href)?;
        }

        // Footer references
        for fref in &self.footer_references {
            write_header_footer_ref(writer, "w:footerReference", fref)?;
        }

        // Page size
        if let Some(ref pg) = self.page_size {
            let mut elem = BytesStart::new("w:pgSz");
            if let Some(w) = pg.width {
                elem.push_attribute(("w:w", w.to_string().as_str()));
            }
            if let Some(h) = pg.height {
                elem.push_attribute(("w:h", h.to_string().as_str()));
            }
            if let Some(ref orient) = pg.orient {
                elem.push_attribute((
                    "w:orient",
                    match orient {
                        PageOrientation::Portrait => "portrait",
                        PageOrientation::Landscape => "landscape",
                    },
                ));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Page margin
        if let Some(ref m) = self.page_margin {
            let mut elem = BytesStart::new("w:pgMar");
            if let Some(v) = m.top {
                elem.push_attribute(("w:top", v.to_string().as_str()));
            }
            if let Some(v) = m.right {
                elem.push_attribute(("w:right", v.to_string().as_str()));
            }
            if let Some(v) = m.bottom {
                elem.push_attribute(("w:bottom", v.to_string().as_str()));
            }
            if let Some(v) = m.left {
                elem.push_attribute(("w:left", v.to_string().as_str()));
            }
            if let Some(v) = m.header {
                elem.push_attribute(("w:header", v.to_string().as_str()));
            }
            if let Some(v) = m.footer {
                elem.push_attribute(("w:footer", v.to_string().as_str()));
            }
            if let Some(v) = m.gutter {
                elem.push_attribute(("w:gutter", v.to_string().as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        // Columns
        if let Some(ref cols) = self.columns {
            let mut elem = BytesStart::new("w:cols");
            if let Some(n) = cols.count {
                elem.push_attribute(("w:num", n.to_string().as_str()));
            }
            if let Some(s) = cols.space {
                elem.push_attribute(("w:space", s.to_string().as_str()));
            }
            if let Some(eq) = cols.equal_width {
                elem.push_attribute(("w:equalWidth", if eq { "1" } else { "0" }));
            }
            if cols.unknown_children.is_empty() {
                writer.write_event(Event::Empty(elem))?;
            } else {
                writer.write_event(Event::Start(elem))?;
                for child in &cols.unknown_children {
                    child.write_to(writer)?;
                }
                writer.write_event(Event::End(BytesEnd::new("w:cols")))?;
            }
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
        Ok(())
    }

    /// Get page orientation
    pub fn orientation(&self) -> Option<&PageOrientation> {
        self.page_size.as_ref()?.orient.as_ref()
    }

    /// Set page to A4 portrait (210mm x 297mm)
    pub fn set_a4_portrait(&mut self) {
        self.page_size = Some(PageSize {
            width: Some(11906),  // 210mm in twips
            height: Some(16838), // 297mm in twips
            orient: Some(PageOrientation::Portrait),
        });
    }

    /// Set page to A4 landscape
    pub fn set_a4_landscape(&mut self) {
        self.page_size = Some(PageSize {
            width: Some(16838),
            height: Some(11906),
            orient: Some(PageOrientation::Landscape),
        });
    }

    /// Set page to US Letter portrait (8.5" x 11")
    pub fn set_letter_portrait(&mut self) {
        self.page_size = Some(PageSize {
            width: Some(12240),  // 8.5 inches in twips
            height: Some(15840), // 11 inches in twips
            orient: Some(PageOrientation::Portrait),
        });
    }
}

// === Parsing helpers ===

fn parse_page_size(e: &BytesStart) -> PageSize {
    PageSize {
        width: get_attr(e, "w:w")
            .or_else(|| get_attr(e, "w"))
            .and_then(|v| v.parse().ok()),
        height: get_attr(e, "w:h")
            .or_else(|| get_attr(e, "h"))
            .and_then(|v| v.parse().ok()),
        orient: get_attr(e, "w:orient")
            .or_else(|| get_attr(e, "orient"))
            .and_then(|v| match v.as_str() {
                "portrait" => Some(PageOrientation::Portrait),
                "landscape" => Some(PageOrientation::Landscape),
                _ => None,
            }),
    }
}

fn parse_page_margin(e: &BytesStart) -> PageMargin {
    PageMargin {
        top: get_attr(e, "w:top").and_then(|v| v.parse().ok()),
        bottom: get_attr(e, "w:bottom").and_then(|v| v.parse().ok()),
        left: get_attr(e, "w:left").and_then(|v| v.parse().ok()),
        right: get_attr(e, "w:right").and_then(|v| v.parse().ok()),
        header: get_attr(e, "w:header").and_then(|v| v.parse().ok()),
        footer: get_attr(e, "w:footer").and_then(|v| v.parse().ok()),
        gutter: get_attr(e, "w:gutter").and_then(|v| v.parse().ok()),
    }
}

fn parse_header_footer_ref(e: &BytesStart) -> Option<HeaderFooterRef> {
    let r_id = get_attr(e, "r:id")?;
    let ref_type = match get_attr(e, "w:type")
        .or_else(|| get_attr(e, "type"))
        .as_deref()
    {
        Some("first") => HeaderFooterType::First,
        Some("even") => HeaderFooterType::Even,
        _ => HeaderFooterType::Default,
    };
    Some(HeaderFooterRef { ref_type, r_id })
}

fn parse_columns<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Columns> {
    let mut cols = parse_columns_empty(start);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let raw = RawXmlElement::from_reader(reader, &e)?;
                cols.unknown_children.push(RawXmlNode::Element(raw));
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
                cols.unknown_children.push(RawXmlNode::Element(raw));
            }
            Event::End(e) if e.name().local_name().as_ref() == b"cols" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(cols)
}

fn parse_columns_empty(e: &BytesStart) -> Columns {
    Columns {
        count: get_attr(e, "w:num").and_then(|v| v.parse().ok()),
        space: get_attr(e, "w:space").and_then(|v| v.parse().ok()),
        equal_width: get_attr(e, "w:equalWidth").map(|v| v == "1" || v == "true"),
        unknown_children: Vec::new(),
    }
}

fn write_header_footer_ref<W: std::io::Write>(
    writer: &mut Writer<W>,
    tag: &str,
    href: &HeaderFooterRef,
) -> Result<()> {
    let mut elem = BytesStart::new(tag);
    elem.push_attribute((
        "w:type",
        match href.ref_type {
            HeaderFooterType::Default => "default",
            HeaderFooterType::First => "first",
            HeaderFooterType::Even => "even",
        },
    ));
    elem.push_attribute(("r:id", href.r_id.as_str()));
    writer.write_event(Event::Empty(elem))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_section_properties() {
        let xml = r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:headerReference w:type="default" r:id="rId7"/>
  <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720"/>
</w:sectPr>"#;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        if let Event::Start(e) = reader.read_event_into(&mut buf).unwrap() {
            let sect = SectionProperties::from_reader(&mut reader, &e).unwrap();

            let pg = sect.page_size.as_ref().unwrap();
            assert_eq!(pg.width, Some(11906));
            assert_eq!(pg.height, Some(16838));
            assert_eq!(pg.orient, Some(PageOrientation::Portrait));

            let margin = sect.page_margin.as_ref().unwrap();
            assert_eq!(margin.top, Some(1440));
            assert_eq!(margin.left, Some(1440));

            assert_eq!(sect.header_references.len(), 1);
            assert_eq!(sect.header_references[0].r_id, "rId7");

            assert_eq!(sect.columns.as_ref().unwrap().space, Some(720));
        }
    }

    #[test]
    fn test_set_a4() {
        let mut sect = SectionProperties::default();
        sect.set_a4_portrait();
        let pg = sect.page_size.as_ref().unwrap();
        assert_eq!(pg.width, Some(11906));
        assert_eq!(pg.height, Some(16838));
    }
}
