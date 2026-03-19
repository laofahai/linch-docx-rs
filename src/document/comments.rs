//! Comments (comments.xml)

use crate::document::Paragraph;
use crate::error::Result;
use crate::xml::{get_attr, RawXmlElement, RawXmlNode};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

/// A single comment
#[derive(Clone, Debug)]
pub struct Comment {
    /// Comment ID
    pub id: u32,
    /// Author name
    pub author: String,
    /// Author initials
    pub initials: Option<String>,
    /// Date (ISO 8601)
    pub date: Option<String>,
    /// Comment content paragraphs
    pub paragraphs: Vec<Paragraph>,
    /// Unknown children (preserved)
    pub unknown_children: Vec<RawXmlNode>,
}

impl Comment {
    /// Create a new comment
    pub fn new(id: u32, author: impl Into<String>, text: impl Into<String>) -> Self {
        Comment {
            id,
            author: author.into(),
            initials: None,
            date: None,
            paragraphs: vec![Paragraph::new(text)],
            unknown_children: Vec::new(),
        }
    }

    /// Get comment text
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Collection of comments from comments.xml
#[derive(Clone, Debug, Default)]
pub struct Comments {
    pub comments: Vec<Comment>,
    pub unknown_children: Vec<RawXmlNode>,
}

impl Comments {
    /// Parse from XML string
    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut comments = Comments::default();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let local = e.name().local_name();
                    if local.as_ref() == b"comment" {
                        comments.comments.push(parse_comment(&mut reader, &e)?);
                    } else if local.as_ref() != b"comments" {
                        let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                        comments.unknown_children.push(RawXmlNode::Element(raw));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(comments)
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

        let mut start = BytesStart::new("w:comments");
        start.push_attribute(("xmlns:w", crate::xml::W));
        start.push_attribute(("xmlns:r", crate::xml::R));
        writer.write_event(Event::Start(start))?;

        for comment in &self.comments {
            let mut cs = BytesStart::new("w:comment");
            cs.push_attribute(("w:id", comment.id.to_string().as_str()));
            cs.push_attribute(("w:author", comment.author.as_str()));
            if let Some(ref initials) = comment.initials {
                cs.push_attribute(("w:initials", initials.as_str()));
            }
            if let Some(ref date) = comment.date {
                cs.push_attribute(("w:date", date.as_str()));
            }
            writer.write_event(Event::Start(cs))?;

            for para in &comment.paragraphs {
                para.write_to(&mut writer)?;
            }
            for child in &comment.unknown_children {
                child.write_to(&mut writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new("w:comment")))?;
        }

        for child in &self.unknown_children {
            child.write_to(&mut writer)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:comments")))?;

        let xml_bytes = buffer.into_inner();
        String::from_utf8(xml_bytes)
            .map_err(|e| crate::error::Error::InvalidDocument(e.to_string()))
    }

    /// Get a comment by ID
    pub fn get(&self, id: u32) -> Option<&Comment> {
        self.comments.iter().find(|c| c.id == id)
    }

    /// Next available ID
    pub fn next_id(&self) -> u32 {
        self.comments.iter().map(|c| c.id).max().unwrap_or(0) + 1
    }

    /// Add a comment, returns the assigned ID
    pub fn add(&mut self, author: impl Into<String>, text: impl Into<String>) -> u32 {
        let id = self.next_id();
        self.comments.push(Comment::new(id, author, text));
        id
    }
}

fn parse_comment<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart,
) -> Result<Comment> {
    let id = get_attr(start, "w:id")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let author = get_attr(start, "w:author").unwrap_or_default();
    let initials = get_attr(start, "w:initials");
    let date = get_attr(start, "w:date");

    let mut comment = Comment {
        id,
        author,
        initials,
        date,
        paragraphs: Vec::new(),
        unknown_children: Vec::new(),
    };

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                if local.as_ref() == b"p" {
                    comment.paragraphs.push(Paragraph::from_reader(reader, &e)?);
                } else {
                    let raw = RawXmlElement::from_reader(reader, &e)?;
                    comment.unknown_children.push(RawXmlNode::Element(raw));
                }
            }
            Event::Empty(e) => {
                if e.name().local_name().as_ref() == b"p" {
                    comment.paragraphs.push(Paragraph::from_empty(&e)?);
                }
            }
            Event::End(e) if e.name().local_name().as_ref() == b"comment" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(comment)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_comments() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Alice" w:date="2024-01-15T10:30:00Z" w:initials="A">
    <w:p><w:r><w:t>Great point!</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Bob">
    <w:p><w:r><w:t>Need to revise this.</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#;

        let comments = Comments::from_xml(xml).unwrap();
        assert_eq!(comments.comments.len(), 2);

        let c1 = comments.get(1).unwrap();
        assert_eq!(c1.author, "Alice");
        assert_eq!(c1.text(), "Great point!");
        assert_eq!(c1.initials.as_deref(), Some("A"));
        assert_eq!(c1.date.as_deref(), Some("2024-01-15T10:30:00Z"));

        let c2 = comments.get(2).unwrap();
        assert_eq!(c2.author, "Bob");
        assert_eq!(c2.text(), "Need to revise this.");
    }

    #[test]
    fn test_comments_roundtrip() {
        let mut comments = Comments::default();
        comments.add("Alice", "First comment");
        comments.add("Bob", "Second comment");

        let xml = comments.to_xml().unwrap();
        let comments2 = Comments::from_xml(&xml).unwrap();

        assert_eq!(comments2.comments.len(), 2);
        assert_eq!(comments2.get(1).unwrap().text(), "First comment");
        assert_eq!(comments2.get(2).unwrap().author, "Bob");
    }
}
