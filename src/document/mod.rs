//! Document model - high-level API for DOCX documents

mod body;
mod footnotes;
mod header_footer;
mod numbering;
mod paragraph;
mod properties;
mod run;
mod section;
mod styles;
mod table;
mod xml_ops;

pub use body::{BlockContent, Body};
pub use footnotes::{Note, Notes};
pub use header_footer::HeaderFooter;
pub use numbering::{AbstractNum, Level, LevelOverride, Num, NumberFormat, Numbering};
pub use paragraph::{
    Alignment, Hyperlink, Indentation, LineSpacing, Paragraph, ParagraphContent,
    ParagraphProperties,
};
pub use properties::CoreProperties;
pub use run::{BreakType, Run, RunContent, RunProperties};
pub use section::{
    Columns, HeaderFooterRef, HeaderFooterType, PageMargin, PageOrientation, PageSize,
    SectionProperties,
};
pub use styles::{DocDefaults, Style, StyleType, Styles};
pub use table::{
    GridColumn, Table, TableAlignment, TableBuilder, TableCell, TableCellProperties, TableRow,
    TableWidth, VMerge, VerticalAlignment,
};

use crate::error::{Error, Result};
use crate::opc::{Package, Part, PartUri};
use std::path::Path;

use xml_ops::{parse_document_xml, serialize_document_xml};

/// List of headers/footers keyed by relationship ID
type HeaderFooterList = Vec<(String, HeaderFooter)>;

/// A DOCX document
#[derive(Debug)]
pub struct Document {
    /// Underlying OPC package
    package: Package,
    /// Parsed document body
    body: Body,
    /// Numbering definitions (from numbering.xml)
    numbering: Option<Numbering>,
    /// Style definitions (from styles.xml)
    styles: Option<Styles>,
    /// Core properties (from core.xml)
    core_properties: Option<CoreProperties>,
    /// Headers (keyed by relationship ID)
    headers: HeaderFooterList,
    /// Footers (keyed by relationship ID)
    footers: HeaderFooterList,
    /// Footnotes
    footnotes: Option<Notes>,
    /// Endnotes
    endnotes: Option<Notes>,
}

impl Document {
    /// Open a document from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let package = Package::open(path)?;
        Self::from_package(package)
    }

    /// Open a document from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    /// Create document from an OPC package
    fn from_package(package: Package) -> Result<Self> {
        // Get main document part
        let doc_part = package
            .main_document_part()
            .ok_or_else(|| Error::MissingPart("Main document part not found".into()))?;

        // Parse document.xml
        let xml = doc_part.data_as_str()?;
        let body = parse_document_xml(xml)?;

        // Try to load numbering.xml
        let numbering = Self::load_numbering(&package);

        // Try to load styles.xml
        let styles = Self::load_styles(&package);

        // Try to load core properties
        let core_properties = Self::load_core_properties(&package);

        // Load headers and footers
        let (headers, footers) = Self::load_headers_footers(&package);

        // Load footnotes and endnotes
        let footnotes = Self::load_notes(&package, true);
        let endnotes = Self::load_notes(&package, false);

        Ok(Self {
            package,
            body,
            numbering,
            styles,
            core_properties,
            headers,
            footers,
            footnotes,
            endnotes,
        })
    }

    /// Load numbering definitions from the package
    fn load_numbering(package: &Package) -> Option<Numbering> {
        // First find the numbering part through relationships
        let doc_part = package.main_document_part()?;
        let rels = doc_part.relationships()?;
        let numbering_rel = rels.by_type(crate::opc::rel_types::NUMBERING)?;

        // Resolve the target URI
        let target = &numbering_rel.target;
        let numbering_uri = if target.starts_with('/') {
            PartUri::new(target).ok()?
        } else {
            PartUri::new(&format!("/word/{}", target)).ok()?
        };

        // Get the numbering part
        let numbering_part = package.part(&numbering_uri)?;
        let xml = numbering_part.data_as_str().ok()?;

        // Parse numbering.xml
        Numbering::from_xml(xml).ok()
    }

    /// Load styles from the package
    fn load_styles(package: &Package) -> Option<Styles> {
        let doc_part = package.main_document_part()?;
        let rels = doc_part.relationships()?;
        let styles_rel = rels.by_type(crate::opc::rel_types::STYLES)?;

        let target = &styles_rel.target;
        let styles_uri = if target.starts_with('/') {
            PartUri::new(target).ok()?
        } else {
            PartUri::new(&format!("/word/{}", target)).ok()?
        };

        let styles_part = package.part(&styles_uri)?;
        let xml_str = styles_part.data_as_str().ok()?;
        Styles::from_xml(xml_str).ok()
    }

    /// Load core properties from the package
    fn load_core_properties(package: &Package) -> Option<CoreProperties> {
        let rel = package
            .relationships()
            .by_type(crate::opc::rel_types::CORE_PROPERTIES)?;
        let target = &rel.target;
        let uri = if target.starts_with('/') {
            PartUri::new(target).ok()?
        } else {
            PartUri::new(&format!("/{}", target)).ok()?
        };
        let part = package.part(&uri)?;
        let xml_str = part.data_as_str().ok()?;
        CoreProperties::from_xml(xml_str).ok()
    }

    /// Load headers and footers from the package
    fn load_headers_footers(package: &Package) -> (HeaderFooterList, HeaderFooterList) {
        let mut headers = Vec::new();
        let mut footers = Vec::new();

        let doc_part = match package.main_document_part() {
            Some(p) => p,
            None => return (headers, footers),
        };
        let rels = match doc_part.relationships() {
            Some(r) => r,
            None => return (headers, footers),
        };

        // Load headers
        for rel in rels.all_by_type(crate::opc::rel_types::HEADER) {
            let uri = if rel.target.starts_with('/') {
                PartUri::new(&rel.target).ok()
            } else {
                PartUri::new(&format!("/word/{}", rel.target)).ok()
            };
            if let Some(uri) = uri {
                if let Some(part) = package.part(&uri) {
                    if let Ok(xml_str) = part.data_as_str() {
                        if let Ok(hf) = HeaderFooter::from_xml(xml_str, true) {
                            headers.push((rel.id.clone(), hf));
                        }
                    }
                }
            }
        }

        // Load footers
        for rel in rels.all_by_type(crate::opc::rel_types::FOOTER) {
            let uri = if rel.target.starts_with('/') {
                PartUri::new(&rel.target).ok()
            } else {
                PartUri::new(&format!("/word/{}", rel.target)).ok()
            };
            if let Some(uri) = uri {
                if let Some(part) = package.part(&uri) {
                    if let Ok(xml_str) = part.data_as_str() {
                        if let Ok(hf) = HeaderFooter::from_xml(xml_str, false) {
                            footers.push((rel.id.clone(), hf));
                        }
                    }
                }
            }
        }

        (headers, footers)
    }

    /// Load footnotes or endnotes from the package
    fn load_notes(package: &Package, is_footnotes: bool) -> Option<Notes> {
        let doc_part = package.main_document_part()?;
        let rels = doc_part.relationships()?;
        let rel_type = if is_footnotes {
            crate::opc::rel_types::FOOTNOTES
        } else {
            crate::opc::rel_types::ENDNOTES
        };
        let rel = rels.by_type(rel_type)?;
        let target = &rel.target;
        let uri = if target.starts_with('/') {
            PartUri::new(target).ok()?
        } else {
            PartUri::new(&format!("/word/{}", target)).ok()?
        };
        let part = package.part(&uri)?;
        let xml_str = part.data_as_str().ok()?;
        Notes::from_xml(xml_str, is_footnotes).ok()
    }

    /// Create a new empty document
    pub fn new() -> Self {
        Self {
            package: Package::new(),
            body: Body::default(),
            numbering: None,
            styles: None,
            core_properties: None,
            headers: Vec::new(),
            footers: Vec::new(),
            footnotes: None,
            endnotes: None,
        }
    }

    /// Save the document to a file
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.update_package()?;
        self.package.save(path)
    }

    /// Save the document to bytes
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.update_package()?;
        self.package.to_bytes()
    }

    /// Update the package with current body content
    fn update_package(&mut self) -> Result<()> {
        let xml = serialize_document_xml(&self.body)?;
        let uri = PartUri::new("/word/document.xml")?;

        // Update or add the document part
        let part = Part::new(
            uri.clone(),
            crate::opc::MAIN_DOCUMENT.to_string(),
            xml.into_bytes(),
        );
        self.package.add_part(part);

        // Ensure the relationship exists for the main document
        if self.package.main_document_part().is_none() {
            use crate::opc::rel_types;
            self.package
                .add_relationship(rel_types::OFFICE_DOCUMENT, uri.as_str());
        }

        // Update numbering.xml if present
        if let Some(ref numbering) = self.numbering {
            let numbering_xml = numbering.to_xml()?;
            let numbering_uri = PartUri::new("/word/numbering.xml")?;
            let numbering_part = Part::new(
                numbering_uri,
                crate::opc::NUMBERING.to_string(),
                numbering_xml.into_bytes(),
            );
            self.package.add_part(numbering_part);
        }

        // Update styles.xml if present
        if let Some(ref styles) = self.styles {
            let styles_xml = styles.to_xml()?;
            let styles_uri = PartUri::new("/word/styles.xml")?;
            let styles_part = Part::new(
                styles_uri,
                crate::opc::STYLES.to_string(),
                styles_xml.into_bytes(),
            );
            self.package.add_part(styles_part);
        }

        // Update core.xml if present
        if let Some(ref core_props) = self.core_properties {
            let core_xml = core_props.to_xml()?;
            let core_uri = PartUri::new("/docProps/core.xml")?;
            let core_part = Part::new(
                core_uri,
                crate::opc::CORE_PROPERTIES.to_string(),
                core_xml.into_bytes(),
            );
            self.package.add_part(core_part);
        }

        // Update headers
        for (r_id, hf) in &self.headers {
            let hf_xml = hf.to_xml()?;
            let hf_uri = PartUri::new(&format!("/word/header_{}.xml", r_id))?;
            let hf_part = Part::new(hf_uri, crate::opc::HEADER.to_string(), hf_xml.into_bytes());
            self.package.add_part(hf_part);
        }

        // Update footnotes
        if let Some(ref fn_notes) = self.footnotes {
            let fn_xml = fn_notes.to_xml()?;
            let fn_uri = PartUri::new("/word/footnotes.xml")?;
            let fn_part = Part::new(
                fn_uri,
                crate::opc::FOOTNOTES.to_string(),
                fn_xml.into_bytes(),
            );
            self.package.add_part(fn_part);
        }

        // Update endnotes
        if let Some(ref en_notes) = self.endnotes {
            let en_xml = en_notes.to_xml()?;
            let en_uri = PartUri::new("/word/endnotes.xml")?;
            let en_part = Part::new(
                en_uri,
                crate::opc::ENDNOTES.to_string(),
                en_xml.into_bytes(),
            );
            self.package.add_part(en_part);
        }

        // Update footers
        for (r_id, hf) in &self.footers {
            let hf_xml = hf.to_xml()?;
            let hf_uri = PartUri::new(&format!("/word/footer_{}.xml", r_id))?;
            let hf_part = Part::new(hf_uri, crate::opc::FOOTER.to_string(), hf_xml.into_bytes());
            self.package.add_part(hf_part);
        }

        Ok(())
    }

    /// Get all paragraphs
    pub fn paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.body.paragraphs()
    }

    /// Get paragraph count
    pub fn paragraph_count(&self) -> usize {
        self.body
            .content
            .iter()
            .filter(|c| matches!(c, BlockContent::Paragraph(_)))
            .count()
    }

    /// Get paragraph by index
    pub fn paragraph(&self, index: usize) -> Option<&Paragraph> {
        self.body.paragraphs().nth(index)
    }

    /// Get all tables
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.body.tables()
    }

    /// Get table count
    pub fn table_count(&self) -> usize {
        self.body
            .content
            .iter()
            .filter(|c| matches!(c, BlockContent::Table(_)))
            .count()
    }

    /// Get table by index
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.body.tables().nth(index)
    }

    /// Get all text in the document
    pub fn text(&self) -> String {
        self.body
            .paragraphs()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get the underlying package
    pub fn package(&self) -> &Package {
        &self.package
    }

    /// Get mutable body
    pub fn body_mut(&mut self) -> &mut Body {
        &mut self.body
    }

    /// Add a paragraph with text
    pub fn add_paragraph(&mut self, text: impl Into<String>) -> &mut Paragraph {
        let para = Paragraph::new(text);
        self.body.add_paragraph(para);
        // Return mutable reference to the last paragraph
        self.body
            .content
            .iter_mut()
            .rev()
            .find_map(|c| {
                if let BlockContent::Paragraph(p) = c {
                    Some(p.as_mut())
                } else {
                    None
                }
            })
            .expect("Just added paragraph")
    }

    /// Add an empty paragraph
    pub fn add_empty_paragraph(&mut self) -> &mut Paragraph {
        self.body.add_paragraph(Paragraph::default());
        self.body
            .content
            .iter_mut()
            .rev()
            .find_map(|c| {
                if let BlockContent::Paragraph(p) = c {
                    Some(p.as_mut())
                } else {
                    None
                }
            })
            .expect("Just added paragraph")
    }

    /// Get numbering definitions
    pub fn numbering(&self) -> Option<&Numbering> {
        self.numbering.as_ref()
    }

    /// Get mutable numbering definitions
    pub fn numbering_mut(&mut self) -> Option<&mut Numbering> {
        self.numbering.as_mut()
    }

    /// Check if a paragraph is a list item
    pub fn is_list_item(&self, para: &Paragraph) -> bool {
        para.properties.as_ref().and_then(|p| p.num_id).is_some()
    }

    /// Check if a paragraph is a bullet list item
    pub fn is_bullet_list_item(&self, para: &Paragraph) -> bool {
        if let Some(num_id) = para.properties.as_ref().and_then(|p| p.num_id) {
            if let Some(ref numbering) = self.numbering {
                return numbering.is_bullet_list(num_id);
            }
        }
        false
    }

    /// Get the list level of a paragraph (0-8), or None if not a list item
    pub fn list_level(&self, para: &Paragraph) -> Option<u32> {
        para.properties.as_ref().and_then(|p| {
            if p.num_id.is_some() {
                Some(p.num_level.unwrap_or(0))
            } else {
                None
            }
        })
    }

    /// Get the number format for a list item
    pub fn list_format(&self, para: &Paragraph) -> Option<&NumberFormat> {
        let props = para.properties.as_ref()?;
        let num_id = props.num_id?;
        let level = props.num_level.unwrap_or(0) as u8;
        self.numbering.as_ref()?.get_format(num_id, level)
    }

    /// Add a table to the document
    pub fn add_table(&mut self, table: Table) -> &mut Table {
        self.body.add_table(table);
        // Return mutable reference to the last table
        self.body
            .content
            .iter_mut()
            .rev()
            .find_map(|c| {
                if let BlockContent::Table(t) = c {
                    Some(t.as_mut())
                } else {
                    None
                }
            })
            .expect("Just added table")
    }

    /// Create and add a table with specified rows and columns
    pub fn add_table_with_size(&mut self, rows: usize, cols: usize) -> &mut Table {
        self.add_table(Table::new(rows, cols))
    }

    /// Get styles
    pub fn styles(&self) -> Option<&Styles> {
        self.styles.as_ref()
    }

    /// Get mutable styles
    pub fn styles_mut(&mut self) -> &mut Styles {
        self.styles.get_or_insert_with(Styles::default)
    }

    /// Get a style by ID
    pub fn style(&self, style_id: &str) -> Option<&Style> {
        self.styles.as_ref()?.get(style_id)
    }

    /// Get core properties
    pub fn core_properties(&self) -> Option<&CoreProperties> {
        self.core_properties.as_ref()
    }

    /// Get mutable core properties (creates default if None)
    pub fn core_properties_mut(&mut self) -> &mut CoreProperties {
        self.core_properties
            .get_or_insert_with(CoreProperties::default)
    }

    /// Get section properties from body
    pub fn section_properties(&self) -> Option<&SectionProperties> {
        self.body.section_properties.as_ref()
    }

    /// Get mutable section properties (creates default if None)
    pub fn section_properties_mut(&mut self) -> &mut SectionProperties {
        self.body
            .section_properties
            .get_or_insert_with(SectionProperties::default)
    }

    /// Get mutable table by index
    pub fn table_mut(&mut self, index: usize) -> Option<&mut Table> {
        self.body
            .content
            .iter_mut()
            .filter_map(|c| {
                if let BlockContent::Table(t) = c {
                    Some(t.as_mut())
                } else {
                    None
                }
            })
            .nth(index)
    }

    /// Get mutable paragraph by index
    pub fn paragraph_mut(&mut self, index: usize) -> Option<&mut Paragraph> {
        self.body
            .content
            .iter_mut()
            .filter_map(|c| {
                if let BlockContent::Paragraph(p) = c {
                    Some(p.as_mut())
                } else {
                    None
                }
            })
            .nth(index)
    }

    /// Get mutable paragraphs iterator
    pub fn paragraphs_mut(&mut self) -> impl Iterator<Item = &mut Paragraph> {
        self.body.paragraphs_mut()
    }

    /// Insert a paragraph at a specific index in the body content
    pub fn insert_paragraph(&mut self, index: usize, para: Paragraph) {
        // Find the position in body.content corresponding to the nth paragraph
        let mut para_count = 0;
        for i in 0..self.body.content.len() {
            if matches!(self.body.content[i], BlockContent::Paragraph(_)) {
                if para_count == index {
                    self.body
                        .content
                        .insert(i, BlockContent::Paragraph(Box::new(para)));
                    return;
                }
                para_count += 1;
            }
        }
        // If index >= paragraph_count, append
        self.body
            .content
            .push(BlockContent::Paragraph(Box::new(para)));
    }

    /// Remove a paragraph by index
    pub fn remove_paragraph(&mut self, index: usize) -> bool {
        let mut para_count = 0;
        for i in 0..self.body.content.len() {
            if matches!(self.body.content[i], BlockContent::Paragraph(_)) {
                if para_count == index {
                    self.body.content.remove(i);
                    return true;
                }
                para_count += 1;
            }
        }
        false
    }

    /// Remove a table by index
    pub fn remove_table(&mut self, index: usize) -> bool {
        let mut table_count = 0;
        for i in 0..self.body.content.len() {
            if matches!(self.body.content[i], BlockContent::Table(_)) {
                if table_count == index {
                    self.body.content.remove(i);
                    return true;
                }
                table_count += 1;
            }
        }
        false
    }

    /// Replace text across all paragraphs. Returns number of replacements.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let mut count = 0;
        for content in &mut self.body.content {
            if let BlockContent::Paragraph(para) = content {
                count += replace_text_in_paragraph(para, find, replace);
            }
        }
        count
    }

    /// Get all headers
    pub fn headers(&self) -> &[(String, HeaderFooter)] {
        &self.headers
    }

    /// Get all footers
    pub fn footers(&self) -> &[(String, HeaderFooter)] {
        &self.footers
    }

    /// Get default header (first one)
    pub fn default_header(&self) -> Option<&HeaderFooter> {
        self.headers.first().map(|(_, hf)| hf)
    }

    /// Get default footer (first one)
    pub fn default_footer(&self) -> Option<&HeaderFooter> {
        self.footers.first().map(|(_, hf)| hf)
    }

    /// Get mutable default header
    pub fn default_header_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.headers.first_mut().map(|(_, hf)| hf)
    }

    /// Get mutable default footer
    pub fn default_footer_mut(&mut self) -> Option<&mut HeaderFooter> {
        self.footers.first_mut().map(|(_, hf)| hf)
    }

    /// Get footnotes
    pub fn footnotes(&self) -> Option<&Notes> {
        self.footnotes.as_ref()
    }

    /// Get mutable footnotes (creates if None)
    pub fn footnotes_mut(&mut self) -> &mut Notes {
        self.footnotes.get_or_insert_with(|| Notes {
            is_footnotes: true,
            ..Default::default()
        })
    }

    /// Get endnotes
    pub fn endnotes(&self) -> Option<&Notes> {
        self.endnotes.as_ref()
    }

    /// Get mutable endnotes (creates if None)
    pub fn endnotes_mut(&mut self) -> &mut Notes {
        self.endnotes.get_or_insert_with(|| Notes {
            is_footnotes: false,
            ..Default::default()
        })
    }

    /// Find text locations across all paragraphs
    pub fn find_text(&self, needle: &str) -> Vec<TextLocation> {
        let mut results = Vec::new();
        for (para_idx, para) in self.body.paragraphs().enumerate() {
            let text = para.text();
            let mut start = 0;
            while let Some(pos) = text[start..].find(needle) {
                results.push(TextLocation {
                    paragraph_index: para_idx,
                    char_offset: start + pos,
                });
                start += pos + needle.len();
            }
        }
        results
    }
}

/// Text location in the document
#[derive(Clone, Debug)]
pub struct TextLocation {
    pub paragraph_index: usize,
    pub char_offset: usize,
}

/// Replace text in a paragraph's runs
fn replace_text_in_paragraph(para: &mut Paragraph, find: &str, replace: &str) -> usize {
    let mut count = 0;
    for content in &mut para.content {
        if let ParagraphContent::Run(run) = content {
            for rc in &mut run.content {
                if let RunContent::Text(ref mut text) = rc {
                    let matches = text.matches(find).count();
                    if matches > 0 {
                        *text = text.replace(find, replace);
                        count += matches;
                    }
                }
            }
        }
    }
    count
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}
