//! Open Packaging Convention (OPC) implementation
//!
//! This module handles the ZIP-based package format used by DOCX files.

mod content_types;
mod package;
mod part;
mod part_uri;
mod relationships;

pub use content_types::{ContentTypes, MAIN_DOCUMENT, RELATIONSHIPS, STYLES, XML};
pub use package::Package;
pub use part::Part;
pub use part_uri::{well_known, PartUri};
pub use relationships::{rel_types, Relationship, Relationships, TargetMode};
