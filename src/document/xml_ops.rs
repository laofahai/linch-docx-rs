//! Document XML parsing and serialization

use crate::document::Body;
use crate::error::{Error, Result};
use crate::xml;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Cursor};

/// Parse document.xml content
pub fn parse_document_xml(xml_str: &str) -> Result<Body> {
    let mut reader = Reader::from_str(xml_str);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut body = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.name().local_name();
                match local.as_ref() {
                    b"body" => {
                        body = Some(Body::from_reader(&mut reader)?);
                    }
                    b"document" => {}
                    _ => {
                        skip_element(&mut reader, &e)?;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    body.ok_or_else(|| Error::InvalidDocument("Missing w:body element".into()))
}

/// Serialize body to document.xml content
pub fn serialize_document_xml(body: &Body) -> Result<String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut writer = Writer::new(&mut buffer);

    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut doc_start = BytesStart::new("w:document");
    for (attr, value) in xml::document_namespaces() {
        doc_start.push_attribute((attr, value));
    }
    writer.write_event(Event::Start(doc_start))?;

    body.write_to(&mut writer)?;

    writer.write_event(Event::End(BytesEnd::new("w:document")))?;

    let xml_bytes = buffer.into_inner();
    String::from_utf8(xml_bytes).map_err(|e| Error::InvalidDocument(e.to_string()))
}

/// Skip an element and all its children
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

    const SIMPLE_DOC: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>This is a heading</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    #[test]
    fn test_parse_simple_document() {
        let body = parse_document_xml(SIMPLE_DOC).unwrap();

        let paras: Vec<_> = body.paragraphs().collect();
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text(), "Hello, World!");
        assert_eq!(paras[1].text(), "This is a heading");
        assert_eq!(paras[1].style(), Some("Heading1"));

        let runs: Vec<_> = paras[1].runs().collect();
        assert_eq!(runs.len(), 1);
        assert!(runs[0].bold());
    }

    #[test]
    fn test_parse_with_formatting() {
        let xml = r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:sz w:val="28"/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>Formatted text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document_xml(xml).unwrap();
        let para = body.paragraphs().next().unwrap();
        let run = para.runs().next().unwrap();

        assert!(run.bold());
        assert!(run.italic());
        assert_eq!(run.font_size_pt(), Some(14.0));
        assert_eq!(run.color(), Some("FF0000"));
    }
}
