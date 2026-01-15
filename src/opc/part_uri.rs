//! Part URI handling for OPC packages

use crate::error::{Error, Result};
use std::fmt;

/// Represents a URI to a part within an OPC package.
///
/// Part URIs are always absolute paths starting with '/'.
/// Example: `/word/document.xml`
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct PartUri {
    path: String,
}

impl PartUri {
    /// Create a new PartUri from a string.
    ///
    /// The path will be normalized (leading '/' ensured, no trailing '/').
    pub fn new(path: &str) -> Result<Self> {
        let path = path.trim();

        if path.is_empty() {
            return Err(Error::InvalidPartUri("empty path".into()));
        }

        // Normalize: ensure leading '/', remove trailing '/'
        let normalized = if path.starts_with('/') {
            path.to_string()
        } else {
            format!("/{}", path)
        };

        let normalized = normalized.trim_end_matches('/').to_string();

        // Validate: no double slashes, no '..' for now
        if normalized.contains("//") {
            return Err(Error::InvalidPartUri(format!(
                "invalid path '{}': contains double slashes",
                path
            )));
        }

        Ok(Self { path: normalized })
    }

    /// Create PartUri without validation (for internal use)
    pub(crate) fn from_string_unchecked(path: String) -> Self {
        Self { path }
    }

    /// Get the path as a string slice
    pub fn as_str(&self) -> &str {
        &self.path
    }

    /// Get the file name portion
    pub fn file_name(&self) -> Option<&str> {
        self.path.rsplit('/').next()
    }

    /// Get the file extension
    pub fn extension(&self) -> Option<&str> {
        self.file_name()
            .and_then(|name| name.rsplit('.').next())
            .filter(|ext| !ext.is_empty() && !ext.contains('/'))
    }

    /// Get the parent directory URI
    pub fn parent(&self) -> Option<PartUri> {
        let pos = self.path.rfind('/')?;
        if pos == 0 {
            None
        } else {
            Some(PartUri {
                path: self.path[..pos].to_string(),
            })
        }
    }

    /// Get the relationships URI for this part.
    ///
    /// For `/word/document.xml`, returns `/word/_rels/document.xml.rels`
    pub fn relationships_uri(&self) -> PartUri {
        let file_name = self.file_name().unwrap_or("");
        let parent = self.parent().map(|p| p.path).unwrap_or_default();

        let rels_path = format!("{}/_rels/{}.rels", parent, file_name);
        PartUri { path: rels_path }
    }

    /// Resolve a relative path against this URI.
    ///
    /// For `/word/document.xml` and `../media/image1.png`, returns `/media/image1.png`
    pub fn resolve(&self, relative: &str) -> Result<PartUri> {
        if relative.starts_with('/') {
            // Absolute path
            return PartUri::new(relative);
        }

        let base_dir = self.parent().map(|p| p.path).unwrap_or_default();
        let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

        for segment in relative.split('/') {
            match segment {
                "" | "." => continue,
                ".." => {
                    parts.pop();
                }
                s => parts.push(s),
            }
        }

        let resolved = format!("/{}", parts.join("/"));
        PartUri::new(&resolved)
    }

    /// Check if this URI points to a relationships file
    pub fn is_relationships(&self) -> bool {
        self.path.contains("/_rels/") && self.path.ends_with(".rels")
    }
}

impl fmt::Display for PartUri {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.path)
    }
}

impl std::str::FromStr for PartUri {
    type Err = Error;

    fn from_str(s: &str) -> Result<Self> {
        PartUri::new(s)
    }
}

/// Well-known part URIs
pub mod well_known {
    use super::PartUri;

    pub fn content_types() -> PartUri {
        PartUri::from_string_unchecked("/[Content_Types].xml".into())
    }

    pub fn package_rels() -> PartUri {
        PartUri::from_string_unchecked("/_rels/.rels".into())
    }

    pub fn document() -> PartUri {
        PartUri::from_string_unchecked("/word/document.xml".into())
    }

    pub fn styles() -> PartUri {
        PartUri::from_string_unchecked("/word/styles.xml".into())
    }

    pub fn numbering() -> PartUri {
        PartUri::from_string_unchecked("/word/numbering.xml".into())
    }

    pub fn core_props() -> PartUri {
        PartUri::from_string_unchecked("/docProps/core.xml".into())
    }

    pub fn app_props() -> PartUri {
        PartUri::from_string_unchecked("/docProps/app.xml".into())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_with_leading_slash() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.as_str(), "/word/document.xml");
    }

    #[test]
    fn test_new_without_leading_slash() {
        let uri = PartUri::new("word/document.xml").unwrap();
        assert_eq!(uri.as_str(), "/word/document.xml");
    }

    #[test]
    fn test_file_name() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.file_name(), Some("document.xml"));
    }

    #[test]
    fn test_extension() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.extension(), Some("xml"));
    }

    #[test]
    fn test_parent() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.parent().unwrap().as_str(), "/word");
    }

    #[test]
    fn test_relationships_uri() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.relationships_uri().as_str(), "/word/_rels/document.xml.rels");
    }

    #[test]
    fn test_resolve_relative() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        let resolved = uri.resolve("../media/image1.png").unwrap();
        assert_eq!(resolved.as_str(), "/media/image1.png");
    }

    #[test]
    fn test_resolve_same_dir() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        let resolved = uri.resolve("styles.xml").unwrap();
        assert_eq!(resolved.as_str(), "/word/styles.xml");
    }

    #[test]
    fn test_is_relationships() {
        let rels = PartUri::new("/word/_rels/document.xml.rels").unwrap();
        assert!(rels.is_relationships());

        let doc = PartUri::new("/word/document.xml").unwrap();
        assert!(!doc.is_relationships());
    }
}
