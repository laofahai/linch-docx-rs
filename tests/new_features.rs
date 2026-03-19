//! Integration tests for new features:
//! styles, properties, section, header/footer, footnotes, text ops, paragraph/run enhancements

use linch_docx_rs::{Alignment, Document, Indentation, LineSpacing, Run, Style, StyleType, Table};
use std::path::Path;

// ============================================================
// Styles
// ============================================================

#[test]
fn test_styles_roundtrip() {
    let mut doc = Document::new();

    // Create styles
    let styles = doc.styles_mut();
    styles.add(Style {
        style_type: Some(StyleType::Paragraph),
        style_id: "MyStyle".into(),
        name: Some("My Custom Style".into()),
        ..Default::default()
    });

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let styles2 = doc2.styles().unwrap();
    let s = styles2.get("MyStyle").unwrap();
    assert_eq!(s.name.as_deref(), Some("My Custom Style"));
    assert_eq!(s.style_type, Some(StyleType::Paragraph));
}

#[test]
fn test_styles_from_existing_doc() {
    let path = Path::new("tests/fixtures/simple.docx");
    if !path.exists() {
        return;
    }
    let doc = Document::open(path).unwrap();

    // simple.docx should have styles
    if let Some(styles) = doc.styles() {
        assert!(styles.iter().count() > 0);
    }
}

// ============================================================
// Core Properties
// ============================================================

#[test]
fn test_core_properties_create() {
    let mut doc = Document::new();

    let props = doc.core_properties_mut();
    props.title = Some("Test Title".into());
    props.creator = Some("Test Author".into());
    props.description = Some("Test Description".into());

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let props2 = doc2.core_properties().unwrap();
    assert_eq!(props2.title.as_deref(), Some("Test Title"));
    assert_eq!(props2.creator.as_deref(), Some("Test Author"));
    assert_eq!(props2.description.as_deref(), Some("Test Description"));
}

// ============================================================
// Section Properties
// ============================================================

#[test]
fn test_section_properties_a4() {
    let mut doc = Document::new();
    doc.section_properties_mut().set_a4_portrait();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let pg = sect.page_size.as_ref().unwrap();
    assert_eq!(pg.width, Some(11906));
    assert_eq!(pg.height, Some(16838));
}

#[test]
fn test_section_properties_letter() {
    let mut doc = Document::new();
    doc.section_properties_mut().set_letter_portrait();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let pg = sect.page_size.as_ref().unwrap();
    assert_eq!(pg.width, Some(12240));
    assert_eq!(pg.height, Some(15840));
}

// ============================================================
// Paragraph Enhancements
// ============================================================

#[test]
fn test_paragraph_alignment() {
    let mut doc = Document::new();

    let p = doc.add_paragraph("Centered text");
    p.set_alignment(Alignment::Center);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let para = doc2.paragraph(0).unwrap();
    assert_eq!(para.alignment(), Some(Alignment::Center));
}

#[test]
fn test_paragraph_indentation() {
    let mut doc = Document::new();

    let p = doc.add_paragraph("Indented");
    p.set_indentation(Indentation {
        left: Some(720),
        first_line: Some(480),
        ..Default::default()
    });

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let para = doc2.paragraph(0).unwrap();
    let ind = para
        .properties
        .as_ref()
        .unwrap()
        .indentation
        .as_ref()
        .unwrap();
    assert_eq!(ind.left, Some(720));
    assert_eq!(ind.first_line, Some(480));
}

#[test]
fn test_paragraph_spacing() {
    let mut doc = Document::new();

    let p = doc.add_paragraph("Spaced");
    p.set_spacing(LineSpacing {
        before: Some(240),
        after: Some(120),
        line: Some(360),
        line_rule: Some("auto".into()),
    });

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sp = doc2
        .paragraph(0)
        .unwrap()
        .properties
        .as_ref()
        .unwrap()
        .spacing
        .as_ref()
        .unwrap();
    assert_eq!(sp.before, Some(240));
    assert_eq!(sp.after, Some(120));
    assert_eq!(sp.line, Some(360));
}

#[test]
fn test_paragraph_heading_level() {
    let mut doc = Document::new();

    let p = doc.add_paragraph("Heading");
    p.set_style("Heading2");

    assert_eq!(doc.paragraph(0).unwrap().heading_level(), Some(2));
}

#[test]
fn test_paragraph_mutations() {
    let mut doc = Document::new();

    doc.add_paragraph("First");
    doc.add_paragraph("Second");
    doc.add_paragraph("Third");
    assert_eq!(doc.paragraph_count(), 3);

    // Insert
    doc.insert_paragraph(1, linch_docx_rs::Paragraph::new("Inserted"));
    assert_eq!(doc.paragraph_count(), 4);
    assert_eq!(doc.paragraph(1).unwrap().text(), "Inserted");

    // Remove
    assert!(doc.remove_paragraph(1));
    assert_eq!(doc.paragraph_count(), 3);
    assert_eq!(doc.paragraph(1).unwrap().text(), "Second");

    // Modify via paragraph_mut
    doc.paragraph_mut(0).unwrap().set_text("Modified First");
    assert_eq!(doc.paragraph(0).unwrap().text(), "Modified First");
}

// ============================================================
// Run Enhancements
// ============================================================

#[test]
fn test_run_full_formatting() {
    let mut doc = Document::new();

    let p = doc.add_empty_paragraph();
    let mut run = Run::new("Fancy text");
    run.set_bold(true);
    run.set_italic(true);
    run.set_underline("single");
    run.set_strike(true);
    run.set_color("0000FF");
    run.set_font_size_pt(16.0);
    run.set_font("Arial");
    run.set_highlight("yellow");
    p.add_run(run);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let r = doc2.paragraph(0).unwrap().runs().next().unwrap();
    assert!(r.bold());
    assert!(r.italic());
    assert_eq!(r.underline(), Some("single"));
    assert!(r.strike());
    assert_eq!(r.color(), Some("0000FF"));
    assert_eq!(r.font_size_pt(), Some(16.0));
    assert_eq!(r.font(), Some("Arial"));
    assert_eq!(r.highlight(), Some("yellow"));
}

#[test]
fn test_run_set_text() {
    let mut doc = Document::new();
    doc.add_paragraph("Original");

    doc.paragraph_mut(0)
        .unwrap()
        .runs_mut()
        .next()
        .unwrap()
        .set_text("Replaced");

    assert_eq!(doc.paragraph(0).unwrap().text(), "Replaced");
}

// ============================================================
// Text Operations
// ============================================================

#[test]
fn test_replace_text() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello World");
    doc.add_paragraph("Hello Again");

    let count = doc.replace_text("Hello", "Hi");
    assert_eq!(count, 2);
    assert_eq!(doc.paragraph(0).unwrap().text(), "Hi World");
    assert_eq!(doc.paragraph(1).unwrap().text(), "Hi Again");
}

#[test]
fn test_find_text() {
    let mut doc = Document::new();
    doc.add_paragraph("foo bar foo");
    doc.add_paragraph("baz");
    doc.add_paragraph("foo");

    let locations = doc.find_text("foo");
    assert_eq!(locations.len(), 3);
    assert_eq!(locations[0].paragraph_index, 0);
    assert_eq!(locations[0].char_offset, 0);
    assert_eq!(locations[1].paragraph_index, 0);
    assert_eq!(locations[1].char_offset, 8);
    assert_eq!(locations[2].paragraph_index, 2);
}

// ============================================================
// Hyperlinks & Bookmarks
// ============================================================

#[test]
fn test_add_hyperlink_and_bookmark() {
    let mut doc = Document::new();

    let p = doc.add_empty_paragraph();
    p.add_hyperlink("rId1", "Click here");
    p.add_internal_link("section1", "Go to section 1");
    p.add_bookmark("1", "section1");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify text is preserved
    let text = doc2.paragraph(0).unwrap().text();
    assert!(text.contains("Click here"));
    assert!(text.contains("Go to section 1"));
}

// ============================================================
// Footnotes & Endnotes
// ============================================================

#[test]
fn test_footnotes_create() {
    let mut doc = Document::new();
    doc.add_paragraph("Main text");

    let id = doc.footnotes_mut().add("This is a footnote.");
    assert_eq!(id, 1);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let fn_notes = doc2.footnotes().unwrap();
    assert_eq!(fn_notes.get(1).unwrap().text(), "This is a footnote.");
}

#[test]
fn test_endnotes_create() {
    let mut doc = Document::new();
    doc.add_paragraph("Main text");

    doc.endnotes_mut().add("Endnote 1");
    doc.endnotes_mut().add("Endnote 2");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let en = doc2.endnotes().unwrap();
    assert_eq!(en.regular_notes().count(), 2);
}

// ============================================================
// Units
// ============================================================

#[test]
fn test_unit_conversions() {
    use linch_docx_rs::{Cm, Inch, Pt, Twip};

    // 1 inch = 72pt = 1440 twips = 2.54cm
    let inch = Inch(1.0);
    assert_eq!(Pt::from(inch), Pt(72.0));
    assert_eq!(Twip::from(inch), Twip(1440));

    let cm = Cm::from(inch);
    assert!((cm.0 - 2.54).abs() < 0.001);

    // Round-trip
    let pt = Pt(12.0);
    let twip = Twip::from(pt);
    let pt2 = Pt::from(twip);
    assert!((pt.0 - pt2.0).abs() < 0.1);
}

// ============================================================
// Full Round-trip with New Features
// ============================================================

#[test]
fn test_full_roundtrip_new_features() {
    let mut doc = Document::new();

    // Section
    doc.section_properties_mut().set_a4_portrait();

    // Properties
    let props = doc.core_properties_mut();
    props.title = Some("Round-trip Test".into());

    // Styles
    doc.styles_mut().add(Style {
        style_type: Some(StyleType::Character),
        style_id: "TestChar".into(),
        name: Some("Test Character".into()),
        ..Default::default()
    });

    // Content
    let h = doc.add_paragraph("Title");
    h.set_style("Heading1");
    h.set_alignment(Alignment::Center);

    let p = doc.add_paragraph("Body text");
    p.set_indentation(Indentation {
        first_line: Some(480),
        ..Default::default()
    });

    // Table
    doc.add_table(Table::new(2, 2));

    // Footnote
    doc.footnotes_mut().add("A footnote");

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify everything survived
    assert_eq!(
        doc2.core_properties().unwrap().title.as_deref(),
        Some("Round-trip Test")
    );
    assert!(doc2.styles().unwrap().get("TestChar").is_some());
    assert_eq!(
        doc2.section_properties()
            .unwrap()
            .page_size
            .as_ref()
            .unwrap()
            .width,
        Some(11906)
    );
    assert_eq!(
        doc2.paragraph(0).unwrap().alignment(),
        Some(Alignment::Center)
    );
    assert_eq!(doc2.table_count(), 1);
    assert_eq!(doc2.footnotes().unwrap().regular_notes().count(), 1);
}
