//! Numbering instance definitions

use crate::error::Result;
use crate::xml::get_w_val;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

use super::level::LevelOverride;

/// Numbering instance (w:num)
#[derive(Clone, Debug)]
pub struct Num {
    /// Numbering ID (referenced by paragraphs)
    pub num_id: u32,
    /// Referenced abstract numbering ID
    pub abstract_num_id: u32,
    /// Level overrides
    pub level_overrides: Vec<LevelOverride>,
}

impl Num {
    /// Create a new numbering instance
    pub fn new(num_id: u32, abstract_num_id: u32) -> Self {
        Num {
            num_id,
            abstract_num_id,
            level_overrides: Vec::new(),
        }
    }

    pub(crate) fn from_reader<R: BufRead>(
        reader: &mut Reader<R>,
        start: &BytesStart,
    ) -> Result<Self> {
        let mut num_id = 0u32;
        let mut abstract_num_id = 0u32;
        let mut level_overrides = Vec::new();

        // Get numId attribute
        for attr in start.attributes().filter_map(|a| a.ok()) {
            let key = attr.key.as_ref();
            if key == b"w:numId" || key == b"numId" {
                num_id = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
            }
        }

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"lvlOverride" {
                        let lo = LevelOverride::from_reader(reader, &e)?;
                        level_overrides.push(lo);
                    } else {
                        super::skip_element(reader, &e)?;
                    }
                }
                Event::Empty(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"abstractNumId" {
                        abstract_num_id =
                            get_w_val(&e).and_then(|v| v.parse().ok()).unwrap_or(0);
                    }
                }
                Event::End(e) => {
                    if e.name().local_name().as_ref() == b"num" {
                        break;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(Num {
            num_id,
            abstract_num_id,
            level_overrides,
        })
    }

    pub(crate) fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new("w:num");
        start.push_attribute(("w:numId", self.num_id.to_string().as_str()));
        writer.write_event(Event::Start(start))?;

        // Abstract num ID
        let mut elem = BytesStart::new("w:abstractNumId");
        elem.push_attribute(("w:val", self.abstract_num_id.to_string().as_str()));
        writer.write_event(Event::Empty(elem))?;

        // Level overrides
        for lo in &self.level_overrides {
            lo.write_to(writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:num")))?;
        Ok(())
    }
}
