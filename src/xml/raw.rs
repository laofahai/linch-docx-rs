//! Raw XML node types for round-trip preservation

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::BufRead;

use crate::error::{Error, Result};

/// Raw XML node for preserving unknown elements during round-trip
#[derive(Clone, Debug)]
pub enum RawXmlNode {
    /// Element node
    Element(RawXmlElement),
    /// Text node
    Text(String),
    /// Comment node
    Comment(String),
}

/// Raw XML element with attributes and children
#[derive(Clone, Debug)]
pub struct RawXmlElement {
    /// Full element name (with prefix, e.g., "w:customXml")
    pub name: String,
    /// Attributes as (name, value) pairs
    pub attributes: Vec<(String, String)>,
    /// Child nodes
    pub children: Vec<RawXmlNode>,
    /// Whether this was a self-closing element
    pub self_closing: bool,
}

impl RawXmlElement {
    /// Create a new empty element
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            attributes: Vec::new(),
            children: Vec::new(),
            self_closing: false,
        }
    }

    /// Read a complete element from XML reader (starting after the start tag was read)
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let name = String::from_utf8_lossy(start.name().as_ref()).to_string();

        let attributes = start
            .attributes()
            .filter_map(|a| a.ok())
            .map(|a| {
                (
                    String::from_utf8_lossy(a.key.as_ref()).to_string(),
                    String::from_utf8_lossy(&a.value).to_string(),
                )
            })
            .collect();

        let mut children = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let child = Self::from_reader(reader, &e)?;
                    children.push(RawXmlNode::Element(child));
                }
                Event::Empty(e) => {
                    let elem = Self {
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
                    children.push(RawXmlNode::Element(elem));
                }
                Event::Text(t) => {
                    let text = t.unescape()?.to_string();
                    if !text.is_empty() {
                        children.push(RawXmlNode::Text(text));
                    }
                }
                Event::Comment(c) => {
                    children.push(RawXmlNode::Comment(String::from_utf8_lossy(&c).to_string()));
                }
                Event::End(e) => {
                    let end_name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                    if end_name == name {
                        break;
                    }
                }
                Event::Eof => return Err(Error::InvalidDocument("Unexpected EOF".into())),
                _ => {}
            }
            buf.clear();
        }

        Ok(Self {
            name,
            attributes,
            children,
            self_closing: false,
        })
    }

    /// Create from empty element tag
    pub fn from_empty(e: &BytesStart) -> Self {
        Self {
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
        }
    }

    /// Write element to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new(&self.name);
        for (key, value) in &self.attributes {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        if self.children.is_empty() && self.self_closing {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            for child in &self.children {
                child.write_to(writer)?;
            }
            writer.write_event(Event::End(BytesEnd::new(&self.name)))?;
        }

        Ok(())
    }

    /// Add an attribute
    pub fn with_attr(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        self.attributes.push((name.into(), value.into()));
        self
    }

    /// Add a child element
    pub fn with_child(mut self, child: RawXmlElement) -> Self {
        self.children.push(RawXmlNode::Element(child));
        self
    }

    /// Add a text child
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.children.push(RawXmlNode::Text(text.into()));
        self
    }
}

impl RawXmlNode {
    /// Write node to XML writer
    pub fn write_to<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            RawXmlNode::Element(e) => e.write_to(writer),
            RawXmlNode::Text(t) => {
                writer.write_event(Event::Text(BytesText::new(t)))?;
                Ok(())
            }
            RawXmlNode::Comment(c) => {
                writer.write_event(Event::Comment(BytesText::new(c)))?;
                Ok(())
            }
        }
    }
}
