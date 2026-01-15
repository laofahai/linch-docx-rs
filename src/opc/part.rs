//! Part representation for OPC packages

use crate::opc::{PartUri, Relationships};

/// A part within an OPC package
#[derive(Clone, Debug)]
pub struct Part {
    /// Part URI
    uri: PartUri,
    /// Content type
    content_type: String,
    /// Part data
    data: Vec<u8>,
    /// Part relationships (if any)
    relationships: Option<Relationships>,
    /// Whether this part has been modified
    modified: bool,
}

impl Part {
    /// Create a new part
    pub fn new(uri: PartUri, content_type: impl Into<String>, data: Vec<u8>) -> Self {
        Self {
            uri,
            content_type: content_type.into(),
            data,
            relationships: None,
            modified: false,
        }
    }

    /// Get the part URI
    pub fn uri(&self) -> &PartUri {
        &self.uri
    }

    /// Get the content type
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Get the raw data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get data as UTF-8 string
    pub fn data_as_str(&self) -> Result<&str, std::str::Utf8Error> {
        std::str::from_utf8(&self.data)
    }

    /// Set the data
    pub fn set_data(&mut self, data: Vec<u8>) {
        self.data = data;
        self.modified = true;
    }

    /// Get relationships
    pub fn relationships(&self) -> Option<&Relationships> {
        self.relationships.as_ref()
    }

    /// Get mutable relationships
    pub fn relationships_mut(&mut self) -> Option<&mut Relationships> {
        self.relationships.as_mut()
    }

    /// Set relationships
    pub fn set_relationships(&mut self, rels: Relationships) {
        self.relationships = Some(rels);
    }

    /// Ensure relationships exist, creating if needed
    pub fn ensure_relationships(&mut self) -> &mut Relationships {
        if self.relationships.is_none() {
            self.relationships = Some(Relationships::new());
        }
        self.relationships.as_mut().unwrap()
    }

    /// Check if the part has been modified
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    /// Mark the part as modified
    pub fn mark_modified(&mut self) {
        self.modified = true;
    }

    /// Get the relationships URI for this part
    pub fn relationships_uri(&self) -> PartUri {
        self.uri.relationships_uri()
    }
}
