//! XML namespaces used in OOXML

/// WordprocessingML main namespace
pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Relationships namespace
pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Drawing namespace
pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
/// DrawingML main namespace
pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Pictures namespace
pub const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
/// Content Types namespace
pub const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
/// Package Relationships namespace
pub const PR: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
/// Core Properties namespace (Dublin Core)
pub const CP: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
/// Dublin Core namespace
pub const DC: &str = "http://purl.org/dc/elements/1.1/";
/// Dublin Core Terms namespace
pub const DCTERMS: &str = "http://purl.org/dc/terms/";

/// Standard namespace declarations for document.xml
pub fn document_namespaces() -> Vec<(&'static str, &'static str)> {
    vec![
        ("xmlns:w", W),
        ("xmlns:r", R),
        ("xmlns:wp", WP),
        ("xmlns:a", A),
        ("xmlns:pic", PIC),
    ]
}

/// Minimal namespace declarations for document.xml
pub fn minimal_document_namespaces() -> Vec<(&'static str, &'static str)> {
    vec![("xmlns:w", W), ("xmlns:r", R)]
}
