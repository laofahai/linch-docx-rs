//! Numbering definitions (numbering.xml)
//!
//! This module handles list numbering in DOCX documents.

mod abstract_num;
mod level;
mod num;
mod types;

pub use abstract_num::AbstractNum;
pub use level::{Level, LevelOverride};
pub use num::Num;
pub use types::{NumberFormat, NumberingInfo};

use crate::error::Result;
use crate::xml::{RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::BufRead;

/// Numbering definitions from numbering.xml
#[derive(Clone, Debug, Default)]
pub struct Numbering {
    /// Abstract numbering definitions
    pub abstract_nums: HashMap<u32, AbstractNum>,
    /// Numbering instances
    pub nums: HashMap<u32, Num>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
    /// Next available abstract num ID
    next_abstract_id: u32,
    /// Next available num ID
    next_num_id: u32,
}

impl Numbering {
    /// Create a new empty numbering definitions
    pub fn new() -> Self {
        Numbering::default()
    }

    /// Parse numbering.xml content
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut numbering = Numbering::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    match local.as_ref() {
                        b"abstractNum" => {
                            let abs_num = AbstractNum::from_reader(&mut reader, &e)?;
                            let id = abs_num.abstract_num_id;
                            numbering.abstract_nums.insert(id, abs_num);
                            if id >= numbering.next_abstract_id {
                                numbering.next_abstract_id = id + 1;
                            }
                        }
                        b"num" => {
                            let num = Num::from_reader(&mut reader, &e)?;
                            let id = num.num_id;
                            numbering.nums.insert(id, num);
                            if id >= numbering.next_num_id {
                                numbering.next_num_id = id + 1;
                            }
                        }
                        b"numbering" => {
                            // Root element, continue
                        }
                        _ => {
                            // Unknown - preserve
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            numbering.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Empty(e) => {
                    // Empty elements at root level - preserve
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
                    numbering.unknown_children.push(RawXmlNode::Element(raw));
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(numbering)
    }

    /// Serialize to XML
    pub fn to_xml(&self) -> Result<String> {
        let mut buffer = Vec::new();
        let mut writer = Writer::new(&mut buffer);

        // XML declaration
        writer.write_event(Event::Decl(quick_xml::events::BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with namespaces
        let mut start = BytesStart::new("w:numbering");
        start.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        start.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(start))?;

        // Write abstract nums (sorted by ID for consistency)
        let mut abs_ids: Vec<_> = self.abstract_nums.keys().collect();
        abs_ids.sort();
        for id in abs_ids {
            self.abstract_nums[id].write_to(&mut writer)?;
        }

        // Write nums (sorted by ID)
        let mut num_ids: Vec<_> = self.nums.keys().collect();
        num_ids.sort();
        for id in num_ids {
            self.nums[id].write_to(&mut writer)?;
        }

        // Write unknown children
        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;

        String::from_utf8(buffer).map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get the format for a specific numId and level
    pub fn get_format(&self, num_id: u32, level: u8) -> Option<&NumberFormat> {
        let num = self.nums.get(&num_id)?;
        let abs_num = self.abstract_nums.get(&num.abstract_num_id)?;
        let lvl = abs_num.levels.get(&level)?;
        lvl.num_fmt.as_ref()
    }

    /// Check if a numId represents a bullet list
    pub fn is_bullet_list(&self, num_id: u32) -> bool {
        if let Some(fmt) = self.get_format(num_id, 0) {
            fmt.is_bullet()
        } else {
            false
        }
    }

    /// Get level text for a specific numId and level
    pub fn get_level_text(&self, num_id: u32, level: u8) -> Option<&str> {
        let num = self.nums.get(&num_id)?;
        let abs_num = self.abstract_nums.get(&num.abstract_num_id)?;
        let lvl = abs_num.levels.get(&level)?;
        lvl.level_text.as_deref()
    }

    /// Get the level definition for a specific numId and level
    pub fn get_level(&self, num_id: u32, level: u8) -> Option<&Level> {
        let num = self.nums.get(&num_id)?;
        let abs_num = self.abstract_nums.get(&num.abstract_num_id)?;
        abs_num.levels.get(&level)
    }

    /// Add a bullet list definition and return the numId
    pub fn add_bullet_list(&mut self) -> u32 {
        let abs_id = self.next_abstract_id;
        self.next_abstract_id += 1;

        let num_id = self.next_num_id;
        self.next_num_id += 1;

        let abs_num = AbstractNum::bullet_list(abs_id);
        self.abstract_nums.insert(abs_id, abs_num);

        let num = Num::new(num_id, abs_id);
        self.nums.insert(num_id, num);

        num_id
    }

    /// Add a decimal numbered list definition and return the numId
    pub fn add_decimal_list(&mut self) -> u32 {
        let abs_id = self.next_abstract_id;
        self.next_abstract_id += 1;

        let num_id = self.next_num_id;
        self.next_num_id += 1;

        let abs_num = AbstractNum::decimal_list(abs_id);
        self.abstract_nums.insert(abs_id, abs_num);

        let num = Num::new(num_id, abs_id);
        self.nums.insert(num_id, num);

        num_id
    }

    /// Add a Chinese numbered list definition (一、二、三) and return the numId
    pub fn add_chinese_list(&mut self) -> u32 {
        let abs_id = self.next_abstract_id;
        self.next_abstract_id += 1;

        let num_id = self.next_num_id;
        self.next_num_id += 1;

        let abs_num = AbstractNum::chinese_list(abs_id);
        self.abstract_nums.insert(abs_id, abs_num);

        let num = Num::new(num_id, abs_id);
        self.nums.insert(num_id, num);

        num_id
    }

    /// Add a custom abstract numbering definition and return the numId
    pub fn add_abstract_num(&mut self, mut abs_num: AbstractNum) -> u32 {
        let abs_id = self.next_abstract_id;
        self.next_abstract_id += 1;
        abs_num.abstract_num_id = abs_id;

        let num_id = self.next_num_id;
        self.next_num_id += 1;

        self.abstract_nums.insert(abs_id, abs_num);

        let num = Num::new(num_id, abs_id);
        self.nums.insert(num_id, num);

        num_id
    }
}

/// Skip an element and all its children
pub(crate) fn skip_element<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<()> {
    let name = start.name();
    let mut depth = 1;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name() == name => depth += 1,
            Event::End(e) if e.name() == name => {
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

    const SAMPLE_NUMBERING: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>"#;

    #[test]
    fn test_parse_numbering() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();

        // Should have 2 abstract nums
        assert_eq!(numbering.abstract_nums.len(), 2);

        // Should have 2 nums
        assert_eq!(numbering.nums.len(), 2);

        // Check abstract num 0
        let abs0 = numbering.abstract_nums.get(&0).unwrap();
        assert_eq!(abs0.multi_level_type, Some("hybridMultilevel".to_string()));
        assert_eq!(abs0.levels.len(), 2);

        // Check level 0
        let lvl0 = abs0.levels.get(&0).unwrap();
        assert_eq!(lvl0.start, Some(1));
        assert_eq!(lvl0.num_fmt, Some(NumberFormat::Decimal));
        assert_eq!(lvl0.level_text, Some("%1.".to_string()));

        // Check num 1 references abstract num 0
        let num1 = numbering.nums.get(&1).unwrap();
        assert_eq!(num1.abstract_num_id, 0);
    }

    #[test]
    fn test_is_bullet_list() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();

        // numId 1 references abstractNum 0 which is decimal
        assert!(!numbering.is_bullet_list(1));

        // numId 2 references abstractNum 1 which is bullet
        assert!(numbering.is_bullet_list(2));
    }

    #[test]
    fn test_roundtrip() {
        let numbering = Numbering::from_xml(SAMPLE_NUMBERING).unwrap();
        let xml = numbering.to_xml().unwrap();

        // Parse again
        let numbering2 = Numbering::from_xml(&xml).unwrap();

        // Should have same structure
        assert_eq!(
            numbering.abstract_nums.len(),
            numbering2.abstract_nums.len()
        );
        assert_eq!(numbering.nums.len(), numbering2.nums.len());

        // Check a specific value
        assert_eq!(numbering.get_format(1, 0), numbering2.get_format(1, 0));
    }

    #[test]
    fn test_add_bullet_list() {
        let mut numbering = Numbering::new();
        let num_id = numbering.add_bullet_list();

        assert_eq!(num_id, 0);
        assert!(numbering.is_bullet_list(num_id));
    }

    #[test]
    fn test_add_decimal_list() {
        let mut numbering = Numbering::new();
        let num_id = numbering.add_decimal_list();

        assert_eq!(num_id, 0);
        assert!(!numbering.is_bullet_list(num_id));
        assert_eq!(numbering.get_format(num_id, 0), Some(&NumberFormat::Decimal));
    }

    #[test]
    fn test_add_chinese_list() {
        let mut numbering = Numbering::new();
        let num_id = numbering.add_chinese_list();

        assert_eq!(num_id, 0);
        assert_eq!(
            numbering.get_format(num_id, 0),
            Some(&NumberFormat::ChineseCounting)
        );
    }
}
