//! Integration test: open a real DOCX file

use linch_docx_rs::opc::{rel_types, Package};
use std::path::Path;

#[test]
fn test_open_simple_docx() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        eprintln!("Test file not found: {:?}", path);
        return;
    }

    let pkg = Package::open(path).expect("Failed to open DOCX");

    // Verify content types loaded
    let ct = pkg.content_types();
    let doc_uri = linch_docx_rs::PartUri::new("/word/document.xml").unwrap();
    assert!(ct.get(&doc_uri).is_some(), "Document content type missing");

    // Verify package relationships
    let rels = pkg.relationships();
    let doc_rel = rels.by_type(rel_types::OFFICE_DOCUMENT);
    assert!(doc_rel.is_some(), "Office document relationship missing");
    assert_eq!(doc_rel.unwrap().target, "word/document.xml");

    // Verify main document part exists
    let doc_part = pkg.main_document_part();
    assert!(doc_part.is_some(), "Main document part missing");

    // Verify document content
    let doc_data = doc_part
        .unwrap()
        .data_as_str()
        .expect("Document should be UTF-8");
    assert!(doc_data.contains("Hello, World!"));
    assert!(doc_data.contains("This is a heading"));
    assert!(doc_data.contains("bold italic"));

    println!("Document content ({} bytes):", doc_data.len());
    println!("{}", doc_data);
}

#[test]
fn test_roundtrip_simple_docx() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    // Open
    let pkg = Package::open(path).expect("Failed to open");

    // Save to bytes
    let bytes = pkg.to_bytes().expect("Failed to save");

    // Reopen
    let pkg2 = Package::from_bytes(&bytes).expect("Failed to reopen");

    // Verify same content
    let doc1 = pkg.main_document_part().unwrap().data_as_str().unwrap();
    let doc2 = pkg2.main_document_part().unwrap().data_as_str().unwrap();
    assert_eq!(
        doc1, doc2,
        "Document content should be identical after roundtrip"
    );
}

#[test]
fn test_list_all_parts() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    let pkg = Package::open(path).expect("Failed to open");

    println!("Parts in package:");
    for (uri, part) in pkg.parts() {
        println!(
            "  {} ({}) - {} bytes",
            uri,
            part.content_type(),
            part.data().len()
        );
    }

    // Should have at least document.xml
    assert!(pkg.part_uris().any(|u| u.as_str() == "/word/document.xml"));
}
