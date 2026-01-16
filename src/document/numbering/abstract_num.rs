//! Abstract numbering definitions

use crate::error::Result;
use crate::xml::{get_w_val, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::BufRead;

use super::level::Level;
use super::types::NumberFormat;

/// Abstract numbering definition (w:abstractNum)
#[derive(Clone, Debug, Default)]
pub struct AbstractNum {
    /// Abstract numbering ID
    pub abstract_num_id: u32,
    /// Multi-level type
    pub multi_level_type: Option<String>,
    /// Level definitions
    pub levels: HashMap<u8, Level>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl AbstractNum {
    /// Create a new abstract numbering definition
    pub fn new(id: u32) -> Self {
        AbstractNum {
            abstract_num_id: id,
            multi_level_type: Some("hybridMultilevel".to_string()),
            ..Default::default()
        }
    }

    /// Add a level to this abstract numbering
    pub fn add_level(&mut self, level: Level) {
        self.levels.insert(level.ilvl, level);
    }

    /// Create a simple bullet list definition
    pub fn bullet_list(id: u32) -> Self {
        let mut abs = Self::new(id);
        abs.add_level(
            Level::new(0)
                .with_format(NumberFormat::Bullet)
                .with_text("•")
                .with_justification("left"),
        );
        abs
    }

    /// Create a simple decimal numbered list definition
    pub fn decimal_list(id: u32) -> Self {
        let mut abs = Self::new(id);
        abs.add_level(
            Level::new(0)
                .with_format(NumberFormat::Decimal)
                .with_text("%1.")
                .with_justification("left"),
        );
        abs
    }

    /// Create a Chinese numbered list definition (一、二、三)
    pub fn chinese_list(id: u32) -> Self {
        let mut abs = Self::new(id);
        abs.add_level(
            Level::new(0)
                .with_format(NumberFormat::ChineseCounting)
                .with_text("%1、")
                .with_justification("left"),
        );
        abs
    }

    pub(crate) fn from_reader<R: BufRead>(
        reader: &mut Reader<R>,
        start: &BytesStart,
    ) -> Result<Self> {
        let mut abs_num = AbstractNum::default();

        // Get abstractNumId attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:abstractNumId" || key == b"abstractNumId" {
                abs_num.abstract_num_id =
                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"lvl" => {
                            let lvl = Level::from_reader(reader, &e)?;
                            abs_num.levels.insert(lvl.ilvl, lvl);
                        }
                        _ => {
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            abs_num.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"multiLevelType" => {
                            abs_num.multi_level_type = get_w_val(&e);
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
                            abs_num.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"abstractNum" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(abs_num)
    }

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:abstractNum");
        start.push_attribute(("w:abstractNumId", self.abstract_num_id.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Multi-level type
        if let Some(mlt) = &self.multi_level_type {
            let mut elem = BytesStart::new("w:multiLevelType");
            elem.push_attribute(("w:val", mlt.as_str()));
            writer.write_event(Event::Empty(elem))?;
        }

        // Write levels (sorted)
        let mut level_ids: Vec<_> = self.levels.keys().collect();
        level_ids.sort();
        for id in level_ids {
            self.levels[id].write_to(writer)?;
        }

        // Unknown children
        for child in &self.unknown_children {
            child.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
        Ok(())
    }
}
