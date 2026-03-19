//! Footnotes and endnotes (footnotes.xml / endnotes.xml)

use crate::document::Paragraph;
use crate::error::Result;
use crate::xml::{get_attr, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

/// A single footnote or endnote
#[derive(Clone, Debug)]
pub struct Note {
    /// Note ID
    pub id: i32,
    /// Note type (normal, separator, continuationSeparator)
    pub note_type: Option<String>,
    /// Paragraphs
    pub paragraphs: Vec<Paragraph>,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Note {
    /// Create a new note with text
    pub fn new(id: i32, text: impl Into<String>) -> Self {
        Note {
            id,
            note_type: None,
            paragraphs: vec![Paragraph::new(text)],
            unknown_children: Vec::new(),
        }
    }

    /// Get all text
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Is this a regular note (not a separator)
    pub fn is_regular(&self) -> bool {
        self.note_type.is_none() || self.note_type.as_deref() == Some("normal")
    }
}

/// Collection of footnotes or endnotes
#[derive(Clone, Debug, Default)]
pub struct Notes {
    /// Individual notes
    pub notes: Vec<Note>,
    /// Whether these are footnotes (true) or endnotes (false)
    pub is_footnotes: bool,
    /// Unknown children (preserved for round-trip)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Notes {
    /// Parse from XML string
    pub fn from_xml(xml: &str, is_footnotes: bool) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut notes = Notes {
            is_footnotes,
            ..Default::default()
        };
        let mut buf = Vec::new();

        let note_tag = if is_footnotes {
            b"footnote".as_slice()
        } else {
            b"endnote".as_slice()
        };

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == note_tag {
                        notes.notes.push(parse_note(&mut reader, &e)?);
                    } else if local.as_ref() != b"footnotes" && local.as_ref() != b"endnotes" {
                        let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                        notes.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(notes)
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

        let root_tag = if self.is_footnotes {
            "w:footnotes"
        } else {
            "w:endnotes"
        };
        let note_tag = if self.is_footnotes {
            "w:footnote"
        } else {
            "w:endnote"
        };

        let mut start = BytesStart::new(root_tag);
        start.push_attribute(("xmlns:w", crate::xml::W));
        start.push_attribute(("xmlns:r", crate::xml::R));
        writer.write_event(Event::Start(start))?;

        for note in &self.notes {
            let mut note_start = BytesStart::new(note_tag);
            note_start.push_attribute(("w:id", note.id.to_string().as_str()));
            if let Some(ref nt) = note.note_type {
                note_start.push_attribute(("w:type", nt.as_str()));
            }

            writer.write_event(Event::Start(note_start))?;

            for para in &note.paragraphs {
                para.write_to(&mut writer)?;
            }

            for child in &note.unknown_children {
                child.write_to(&mut writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new(note_tag)))?;
        }

        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new(root_tag)))?;

        let xml_bytes = buffer.into_inner();
        String::from_utf8(xml_bytes)
            .map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get a note by ID
    pub fn get(&self, id: i32) -> Option<&Note> {
        self.notes.iter().find(|n| n.id == id)
    }

    /// Get regular notes (excluding separators)
    pub fn regular_notes(&self) -> impl Iterator<Item = &Note> {
        self.notes.iter().filter(|n| n.is_regular())
    }

    /// Next available ID
    pub fn next_id(&self) -> i32 {
        self.notes.iter().map(|n| n.id).max().unwrap_or(0) + 1
    }

    /// Add a note, returns the assigned ID
    pub fn add(&mut self, text: impl Into<String>) -> i32 {
        let id = self.next_id();
        self.notes.push(Note::new(id, text));
        id
    }
}

fn parse_note<R: std::io::BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Note> {
    let id = get_attr(start, "w:id")
        .or_else(|| get_attr(start, "id"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let note_type = get_attr(start, "w:type").or_else(|| get_attr(start, "type"));

    let mut note = Note {
        id,
        note_type,
        paragraphs: Vec::new(),
        unknown_children: Vec::new(),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"p" {
                    note.paragraphs.push(Paragraph::from_reader(reader, &e)?);
                } else {
                    let raw = RawXmlElement::from_reader(reader, &e)?;
                    note.unknown_children.push(RawXmlNode::Element(raw));
                }
            }
            Event::Empty(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"p" {
                    note.paragraphs.push(Paragraph::from_empty(&e)?);
                }
            }
            Event::End(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"footnote" || local.as_ref() == b"endnote" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(note)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_footnotes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:t>This is footnote 1</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>"#;

        let notes = Notes::from_xml(xml, true).unwrap();
        assert_eq!(notes.notes.len(), 2);
        assert!(!notes.notes[0].is_regular());
        assert!(notes.notes[1].is_regular());
        assert_eq!(notes.get(1).unwrap().text(), "This is footnote 1");
        assert_eq!(notes.regular_notes().count(), 1);
    }

    #[test]
    fn test_notes_roundtrip() {
        let mut notes = Notes {
            is_footnotes: true,
            ..Default::default()
        };
        let id = notes.add("Test footnote");
        assert_eq!(id, 1);

        let xml = notes.to_xml().unwrap();
        let notes2 = Notes::from_xml(&xml, true).unwrap();
        assert_eq!(notes2.notes.len(), 1);
        assert_eq!(notes2.get(1).unwrap().text(), "Test footnote");
    }

    #[test]
    fn test_endnotes() {
        let mut notes = Notes {
            is_footnotes: false,
            ..Default::default()
        };
        notes.add("Endnote content");

        let xml = notes.to_xml().unwrap();
        assert!(xml.contains("w:endnotes"));
        assert!(xml.contains("w:endnote"));
    }
}
