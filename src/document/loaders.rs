//! Package loading helpers for Document parts

use crate::document::comments::Comments;
use crate::document::footnotes::Notes;
use crate::document::header_footer::HeaderFooter;
use crate::document::properties::CoreProperties;
use crate::document::styles::Styles;
use crate::document::HeaderFooterList;
use crate::document::Numbering;
use crate::opc::{Package, PartUri};

/// Load numbering definitions from the package
pub fn load_numbering(package: &Package) -> Option<Numbering> {
    let doc_part = package.main_document_part()?;
    let rels = doc_part.relationships()?;
    let numbering_rel = rels.by_type(crate::opc::rel_types::NUMBERING)?;
    let target = &numbering_rel.target;
    let numbering_uri = if target.starts_with('/') {
        PartUri::new(target).ok()?
    } else {
        PartUri::new(&format!("/word/{}", target)).ok()?
    };
    let numbering_part = package.part(&numbering_uri)?;
    let xml = numbering_part.data_as_str().ok()?;
    Numbering::from_xml(xml).ok()
}

/// Load styles from the package
pub fn load_styles(package: &Package) -> Option<Styles> {
    load_doc_part_by_rel(package, crate::opc::rel_types::STYLES)
        .and_then(|xml| Styles::from_xml(&xml).ok())
}

/// Load core properties from the package (package-level relationship)
pub fn load_core_properties(package: &Package) -> Option<CoreProperties> {
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
    CoreProperties::from_xml(part.data_as_str().ok()?).ok()
}

/// Load headers and footers from the package
pub fn load_headers_footers(package: &Package) -> (HeaderFooterList, HeaderFooterList) {
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

    for rel in rels.all_by_type(crate::opc::rel_types::HEADER) {
        if let Some(hf) = load_hf_part(package, &rel.target, true) {
            headers.push((rel.id.clone(), hf));
        }
    }
    for rel in rels.all_by_type(crate::opc::rel_types::FOOTER) {
        if let Some(hf) = load_hf_part(package, &rel.target, false) {
            footers.push((rel.id.clone(), hf));
        }
    }

    (headers, footers)
}

/// Load footnotes or endnotes
pub fn load_notes(package: &Package, is_footnotes: bool) -> Option<Notes> {
    let rel_type = if is_footnotes {
        crate::opc::rel_types::FOOTNOTES
    } else {
        crate::opc::rel_types::ENDNOTES
    };
    let xml = load_doc_part_by_rel(package, rel_type)?;
    Notes::from_xml(&xml, is_footnotes).ok()
}

/// Load comments
pub fn load_comments(package: &Package) -> Option<Comments> {
    let xml = load_doc_part_by_rel(package, crate::opc::rel_types::COMMENTS)?;
    Comments::from_xml(&xml).ok()
}

/// Helper: load a document-level part by relationship type, returning its XML string
fn load_doc_part_by_rel(package: &Package, rel_type: &str) -> Option<String> {
    let doc_part = package.main_document_part()?;
    let rels = doc_part.relationships()?;
    let rel = rels.by_type(rel_type)?;
    let target = &rel.target;
    let uri = if target.starts_with('/') {
        PartUri::new(target).ok()?
    } else {
        PartUri::new(&format!("/word/{}", target)).ok()?
    };
    let part = package.part(&uri)?;
    Some(part.data_as_str().ok()?.to_string())
}

/// Helper: load a header/footer part
fn load_hf_part(package: &Package, target: &str, is_header: bool) -> Option<HeaderFooter> {
    let uri = if target.starts_with('/') {
        PartUri::new(target).ok()?
    } else {
        PartUri::new(&format!("/word/{}", target)).ok()?
    };
    let part = package.part(&uri)?;
    HeaderFooter::from_xml(part.data_as_str().ok()?, is_header).ok()
}
