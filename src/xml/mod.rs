//! XML utilities and raw element preservation for round-trip support

mod namespace;
mod raw;

pub use namespace::*;
pub use raw::{RawXmlElement, RawXmlNode};

use quick_xml::events::BytesStart;

/// Helper to get attribute value from BytesStart
pub fn get_attr(element: &BytesStart, name: &str) -> Option<String> {
    element
        .attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name.as_bytes())
        .map(|a| String::from_utf8_lossy(&a.value).to_string())
}

/// Helper to get w:val attribute (common in OOXML)
pub fn get_w_val(element: &BytesStart) -> Option<String> {
    get_attr(element, "w:val").or_else(|| get_attr(element, "val"))
}

/// Parse a boolean value from OOXML (handles "1", "true", "on", or missing val)
pub fn parse_bool(element: &BytesStart) -> bool {
    match get_w_val(element) {
        None => true, // No val attribute means true (e.g., <w:b/>)
        Some(v) => matches!(v.as_str(), "1" | "true" | "on"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;
    use quick_xml::Reader;

    #[test]
    fn test_raw_element_roundtrip() {
        let xml = r#"<w:custom foo="bar"><w:child>text</w:child></w:custom>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        if let Event::Start(e) = reader.read_event_into(&mut buf).unwrap() {
            let elem = RawXmlElement::from_reader(&mut reader, &e).unwrap();

            assert_eq!(elem.name, "w:custom");
            assert_eq!(elem.attributes.len(), 1);
            assert_eq!(elem.children.len(), 1);
        }
    }

    #[test]
    fn test_namespace_constants() {
        assert!(W.contains("wordprocessingml"));
        assert!(R.contains("relationships"));
    }
}
