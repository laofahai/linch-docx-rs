//! Integration test: Document API

use linch_docx_rs::Document;
use std::path::Path;

#[test]
fn test_document_open_and_read() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        eprintln!("Test file not found: {:?}", path);
        return;
    }

    let doc = Document::open(path).expect("Failed to open document");

    // Test paragraph count
    assert_eq!(doc.paragraph_count(), 3);

    // Test text extraction
    let text = doc.text();
    assert!(text.contains("Hello, World!"));
    assert!(text.contains("This is a heading"));
    assert!(text.contains("bold italic"));

    println!("Document text:\n{}", text);
}

#[test]
fn test_paragraph_iteration() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    let doc = Document::open(path).unwrap();

    println!("\n=== Paragraphs ===");
    for (i, para) in doc.paragraphs().enumerate() {
        println!("Paragraph {}: \"{}\"", i, para.text());
        println!("  Style: {:?}", para.style());
        println!("  Is heading: {}", para.is_heading());

        // Print runs
        for (j, run) in para.runs().enumerate() {
            println!("  Run {}: \"{}\"", j, run.text());
            println!("    Bold: {}, Italic: {}", run.bold(), run.italic());
            if let Some(size) = run.font_size_pt() {
                println!("    Size: {}pt", size);
            }
        }
    }
}

#[test]
fn test_run_formatting() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    let doc = Document::open(path).unwrap();

    // Third paragraph should have "bold italic" run
    let para = doc.paragraph(2).expect("Should have 3rd paragraph");
    let runs: Vec<_> = para.runs().collect();

    // Find the bold+italic run
    let bold_italic_run = runs.iter().find(|r| r.bold() && r.italic());
    assert!(bold_italic_run.is_some(), "Should find bold+italic run");
    assert_eq!(bold_italic_run.unwrap().text(), "bold italic");
}

#[test]
fn test_heading_detection() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    let doc = Document::open(path).unwrap();

    // Second paragraph is a heading
    let para = doc.paragraph(1).expect("Should have 2nd paragraph");
    assert!(para.is_heading(), "Second paragraph should be a heading");
    assert_eq!(para.style(), Some("Heading1"));
}

#[test]
fn test_document_from_bytes() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }

    let bytes = std::fs::read(path).unwrap();
    let doc = Document::from_bytes(&bytes).expect("Failed to open from bytes");

    assert_eq!(doc.paragraph_count(), 3);
    assert!(doc.text().contains("Hello, World!"));
}

#[test]
fn test_create_new_document() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello from Rust!");
    doc.add_paragraph("This is the second paragraph.");

    // Verify paragraphs were added
    assert_eq!(doc.paragraph_count(), 2);
    assert_eq!(doc.text(), "Hello from Rust!\nThis is the second paragraph.");

    // Save to bytes
    let bytes = doc.to_bytes().expect("Should serialize to bytes");
    assert!(!bytes.is_empty(), "Should produce non-empty output");

    // Verify it's a valid ZIP
    assert_eq!(&bytes[0..2], b"PK", "Should be valid ZIP file");

    println!("Created DOCX with {} bytes", bytes.len());
}

#[test]
fn test_create_document_with_formatting() {
    use linch_docx_rs::Run;

    let mut doc = Document::new();

    // Add a heading
    let heading = doc.add_paragraph("My Document Title");
    heading.set_style("Heading1");

    // Add paragraph with formatted run
    let para = doc.add_empty_paragraph();
    let mut run = Run::new("Bold text");
    run.set_bold(true);
    para.add_run(run);

    // Add another run with different formatting
    let mut run2 = Run::new(" and italic text");
    run2.set_italic(true);
    para.add_run(run2);

    assert_eq!(doc.paragraph_count(), 2);

    // Save and reload
    let bytes = doc.to_bytes().expect("Should serialize");
    let doc2 = Document::from_bytes(&bytes).expect("Should deserialize");

    // Verify content is preserved
    assert_eq!(doc2.paragraph_count(), 2);
    assert!(doc2.text().contains("My Document Title"));
    assert!(doc2.text().contains("Bold text"));
    assert!(doc2.text().contains("italic text"));
}

#[test]
fn test_roundtrip_existing_document() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        eprintln!("Test file not found: {:?}", path);
        return;
    }

    // Open existing document
    let mut doc = Document::open(path).expect("Should open");
    let original_text = doc.text();
    let original_count = doc.paragraph_count();

    // Save and reload
    let bytes = doc.to_bytes().expect("Should serialize");
    let doc2 = Document::from_bytes(&bytes).expect("Should deserialize");

    // Verify content is preserved
    assert_eq!(doc2.paragraph_count(), original_count);
    assert_eq!(doc2.text(), original_text);
}

#[test]
fn test_save_to_file() {
    let output_path = Path::new("target/test_output.docx");

    let mut doc = Document::new();
    doc.add_paragraph("Test document created by linch-docx-rs");
    doc.add_paragraph("This is a test paragraph.");

    // Save to file
    doc.save(output_path).expect("Should save to file");

    // Verify file exists and can be reopened
    assert!(output_path.exists(), "Output file should exist");

    let doc2 = Document::open(output_path).expect("Should reopen");
    assert_eq!(doc2.paragraph_count(), 2);
    assert!(doc2.text().contains("Test document"));

    // Clean up
    std::fs::remove_file(output_path).ok();
}
