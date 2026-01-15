//! OPC Package implementation
//!
//! Handles reading and writing DOCX files as ZIP packages

use crate::error::{Error, Result};
use crate::opc::{ContentTypes, Part, PartUri, Relationships};
use crate::opc::relationships::rel_types;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::Path;
use zip::read::ZipArchive;
use zip::write::{FileOptions, ZipWriter};
use zip::CompressionMethod;

/// An OPC package (ZIP-based container for DOCX, XLSX, PPTX, etc.)
#[derive(Debug)]
pub struct Package {
    /// All parts in the package
    parts: HashMap<PartUri, Part>,
    /// Package-level relationships (/_rels/.rels)
    relationships: Relationships,
    /// Content types ([Content_Types].xml)
    content_types: ContentTypes,
}

impl Package {
    /// Create a new empty package
    pub fn new() -> Self {
        Self {
            parts: HashMap::new(),
            relationships: Relationships::new(),
            content_types: ContentTypes::new(),
        }
    }

    /// Open a package from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_reader(reader)
    }

    /// Open a package from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        Self::from_reader(cursor)
    }

    /// Open a package from a reader
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut archive = ZipArchive::new(reader)?;
        let mut package = Self::new();

        // Step 1: Read [Content_Types].xml
        package.content_types = Self::read_content_types(&mut archive)?;

        // Step 2: Read package relationships (/_rels/.rels)
        package.relationships = Self::read_package_rels(&mut archive)?;

        // Step 3: Read all parts
        package.read_parts(&mut archive)?;

        // Step 4: Read part relationships
        package.read_part_relationships(&mut archive)?;

        Ok(package)
    }

    /// Save the package to a file
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let file = File::create(path)?;
        self.write_to(file)
    }

    /// Save the package to bytes
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut buf = Vec::new();
        let cursor = Cursor::new(&mut buf);
        self.write_to(cursor)?;
        Ok(buf)
    }

    /// Write the package to a writer
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options: FileOptions<()> = FileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        // Write [Content_Types].xml
        zip.start_file("[Content_Types].xml", options)?;
        self.content_types.write_to(&mut zip)?;

        // Write package relationships
        if !self.relationships.is_empty() {
            zip.start_file("_rels/.rels", options)?;
            self.relationships.write_to(&mut zip)?;
        }

        // Write all parts
        for (uri, part) in &self.parts {
            let path = &uri.as_str()[1..]; // Remove leading '/'
            zip.start_file(path, options)?;
            zip.write_all(part.data())?;

            // Write part relationships if any
            if let Some(rels) = part.relationships() {
                if !rels.is_empty() {
                    let rels_uri = uri.relationships_uri();
                    let rels_path = &rels_uri.as_str()[1..];
                    zip.start_file(rels_path, options)?;
                    rels.write_to(&mut zip)?;
                }
            }
        }

        zip.finish()?;
        Ok(())
    }

    /// Get a part by URI
    pub fn part(&self, uri: &PartUri) -> Option<&Part> {
        self.parts.get(uri)
    }

    /// Get a mutable part by URI
    pub fn part_mut(&mut self, uri: &PartUri) -> Option<&mut Part> {
        self.parts.get_mut(uri)
    }

    /// Add a part to the package
    pub fn add_part(&mut self, part: Part) {
        let uri = part.uri().clone();
        self.content_types.add_override(&uri, part.content_type());
        self.parts.insert(uri, part);
    }

    /// Remove a part from the package
    pub fn remove_part(&mut self, uri: &PartUri) -> Option<Part> {
        self.content_types.remove_override(uri);
        self.parts.remove(uri)
    }

    /// Get all part URIs
    pub fn part_uris(&self) -> impl Iterator<Item = &PartUri> {
        self.parts.keys()
    }

    /// Get all parts
    pub fn parts(&self) -> impl Iterator<Item = (&PartUri, &Part)> {
        self.parts.iter()
    }

    /// Get package-level relationships
    pub fn relationships(&self) -> &Relationships {
        &self.relationships
    }

    /// Get mutable package-level relationships
    pub fn relationships_mut(&mut self) -> &mut Relationships {
        &mut self.relationships
    }

    /// Get content types
    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    /// Get mutable content types
    pub fn content_types_mut(&mut self) -> &mut ContentTypes {
        &mut self.content_types
    }

    /// Get a part by relationship type from package relationships
    pub fn part_by_rel_type(&self, rel_type: &str) -> Option<&Part> {
        let rel = self.relationships.by_type(rel_type)?;
        let uri = PartUri::new(&rel.target).ok()?;
        self.parts.get(&uri)
    }

    /// Get the main document part
    pub fn main_document_part(&self) -> Option<&Part> {
        self.part_by_rel_type(rel_types::OFFICE_DOCUMENT)
    }

    /// Get the main document part mutably
    pub fn main_document_part_mut(&mut self) -> Option<&mut Part> {
        let rel = self.relationships.by_type(rel_types::OFFICE_DOCUMENT)?;
        let uri = PartUri::new(&rel.target).ok()?;
        self.parts.get_mut(&uri)
    }

    /// Add a package-level relationship
    pub fn add_relationship(&mut self, rel_type: &str, target: &str) -> String {
        self.relationships.add(rel_type, target)
    }

    // === Private methods ===

    fn read_content_types<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<ContentTypes> {
        let mut file = archive
            .by_name("[Content_Types].xml")
            .map_err(|_| Error::MissingPart("[Content_Types].xml".into()))?;

        let mut content = String::new();
        file.read_to_string(&mut content)?;

        ContentTypes::from_xml(&content)
    }

    fn read_package_rels<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Relationships> {
        match archive.by_name("_rels/.rels") {
            Ok(mut file) => {
                let mut content = String::new();
                file.read_to_string(&mut content)?;
                Relationships::from_xml(&content)
            }
            Err(_) => Ok(Relationships::new()),
        }
    }

    fn read_parts<R: Read + Seek>(&mut self, archive: &mut ZipArchive<R>) -> Result<()> {
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();

            // Skip directories
            if name.ends_with('/') {
                continue;
            }

            // Skip special files
            if name == "[Content_Types].xml" {
                continue;
            }

            // Skip relationship files (handle separately)
            if name.contains("_rels/") && name.ends_with(".rels") {
                continue;
            }

            let uri = PartUri::new(&format!("/{}", name))?;

            // Get content type
            let content_type = self
                .content_types
                .get(&uri)
                .unwrap_or("application/octet-stream")
                .to_string();

            // Read data
            let mut data = Vec::new();
            file.read_to_end(&mut data)?;

            let part = Part::new(uri.clone(), content_type, data);
            self.parts.insert(uri, part);
        }

        Ok(())
    }

    fn read_part_relationships<R: Read + Seek>(&mut self, archive: &mut ZipArchive<R>) -> Result<()> {
        // Collect all part URIs first to avoid borrow issues
        let part_uris: Vec<PartUri> = self.parts.keys().cloned().collect();

        for uri in part_uris {
            let rels_path = uri.relationships_uri();
            let rels_zip_path = rels_path.as_str()[1..].to_string(); // Remove leading '/'

            if let Ok(mut file) = archive.by_name(&rels_zip_path) {
                let mut content = String::new();
                file.read_to_string(&mut content)?;
                let rels = Relationships::from_xml(&content)?;

                if let Some(part) = self.parts.get_mut(&uri) {
                    part.set_relationships(rels);
                }
            }
        }

        Ok(())
    }
}

impl Default for Package {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_package() {
        let pkg = Package::new();
        assert!(pkg.parts.is_empty());
        assert!(pkg.relationships.is_empty());
    }

    #[test]
    fn test_add_part() {
        let mut pkg = Package::new();
        let uri = PartUri::new("/word/document.xml").unwrap();
        let part = Part::new(uri.clone(), "application/xml", b"<doc/>".to_vec());

        pkg.add_part(part);

        assert!(pkg.part(&uri).is_some());
        assert_eq!(pkg.part(&uri).unwrap().data(), b"<doc/>");
    }

    #[test]
    fn test_roundtrip_empty() {
        let pkg = Package::new();
        let bytes = pkg.to_bytes().unwrap();

        let pkg2 = Package::from_bytes(&bytes).unwrap();
        assert!(pkg2.parts.is_empty());
    }

    #[test]
    fn test_roundtrip_with_parts() {
        let mut pkg = Package::new();

        // Add document part
        let doc_uri = PartUri::new("/word/document.xml").unwrap();
        let doc_data = b"<?xml version=\"1.0\"?><document/>".to_vec();
        let doc_part = Part::new(
            doc_uri.clone(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            doc_data,
        );
        pkg.add_part(doc_part);

        // Add relationship
        pkg.relationships_mut()
            .add(rel_types::OFFICE_DOCUMENT, "word/document.xml");

        // Save and reload
        let bytes = pkg.to_bytes().unwrap();
        let pkg2 = Package::from_bytes(&bytes).unwrap();

        // Verify
        assert!(pkg2.part(&doc_uri).is_some());
        assert!(pkg2.main_document_part().is_some());
    }
}
