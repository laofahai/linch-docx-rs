//! Error types for linch-docx-rs

use thiserror::Error;

/// Main error type
#[derive(Debug, Error)]
pub enum Error {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("XML encoding error: {0}")]
    XmlEncoding(#[from] quick_xml::encoding::EncodingError),

    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),

    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("Missing required part: {0}")]
    MissingPart(String),

    #[error("Invalid part URI: {0}")]
    InvalidPartUri(String),

    #[error("Invalid content type: {0}")]
    InvalidContentType(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    #[error("Missing attribute '{attr}' on element '{element}'")]
    MissingAttribute { element: String, attr: String },

    #[error("Invalid document: {0}")]
    InvalidDocument(String),

    #[error("Part not found: {0}")]
    PartNotFound(String),
}

/// Result type alias
pub type Result<T> = std::result::Result<T, Error>;
