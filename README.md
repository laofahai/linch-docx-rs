# linch-docx-rs

A reliable DOCX reading and writing library for Rust with **round-trip preservation**.

[![Crates.io](https://img.shields.io/crates/v/linch-docx-rs.svg)](https://crates.io/crates/linch-docx-rs)
[![Documentation](https://docs.rs/linch-docx-rs/badge.svg)](https://docs.rs/linch-docx-rs)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Read & Write DOCX** - Full support for reading and creating Word documents
- **Round-trip Preservation** - Unknown elements are kept intact during read-modify-save operations
- **Simple API** - Pythonic API design inspired by [python-docx](https://python-docx.readthedocs.io/)
- **Type Safe** - Leverages Rust's type system for reliability
- **Zero Unsafe** - Pure safe Rust implementation

## Quick Start

Add to your `Cargo.toml`:

```toml
[dependencies]
linch-docx-rs = "0.1"
```

### Reading a Document

```rust
use linch_docx_rs::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open("example.docx")?;

    // Get all text
    println!("{}", doc.text());

    // Iterate paragraphs
    for para in doc.paragraphs() {
        println!("Style: {:?}, Text: {}", para.style(), para.text());

        for run in para.runs() {
            if run.bold() { print!("[B] "); }
            if run.italic() { print!("[I] "); }
            println!("{}", run.text());
        }
    }

    // Access styles, properties, sections
    if let Some(styles) = doc.styles() {
        println!("Styles: {}", styles.iter().count());
    }
    if let Some(props) = doc.core_properties() {
        println!("Title: {:?}", props.title);
    }

    Ok(())
}
```

### Creating a Document

```rust
use linch_docx_rs::{Alignment, Document, Indentation, Run, Table};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = Document::new();

    // Page setup
    doc.section_properties_mut().set_a4_portrait();

    // Document properties
    let props = doc.core_properties_mut();
    props.title = Some("My Report".into());
    props.creator = Some("linch-docx-rs".into());

    // Add a centered heading
    let heading = doc.add_paragraph("Project Report");
    heading.set_style("Heading1");
    heading.set_alignment(Alignment::Center);

    // Add formatted text
    let para = doc.add_empty_paragraph();
    let mut bold = Run::new("Important: ");
    bold.set_bold(true);
    bold.set_color("FF0000");
    para.add_run(bold);
    para.add_run(Run::new("This is the body text."));

    // Add a table
    let table = Table::new(2, 3);
    let t = doc.add_table(table);
    t.set_cell_text(0, 0, "Name");
    t.set_cell_text(0, 1, "Age");
    t.set_cell_text(0, 2, "City");
    t.set_cell_text(1, 0, "Alice");
    t.set_cell_text(1, 1, "30");
    t.set_cell_text(1, 2, "Beijing");

    doc.save("output.docx")?;
    Ok(())
}
```

### Modifying an Existing Document

```rust
use linch_docx_rs::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = Document::open("template.docx")?;

    // Find and replace text
    let count = doc.replace_text("{{name}}", "Alice");
    println!("Replaced {} occurrences", count);

    // Modify paragraphs
    if let Some(para) = doc.paragraph_mut(0) {
        para.set_text("Updated first paragraph");
    }

    // Add new content
    doc.add_paragraph("Appended paragraph");

    // Save - all original elements are preserved!
    doc.save("filled.docx")?;
    Ok(())
}
```

## API Overview

### Document

| Method | Description |
|--------|-------------|
| `Document::new()` | Create a new empty document |
| `Document::open(path)` | Open a document from file |
| `Document::from_bytes(bytes)` | Open a document from bytes |
| `doc.save(path)` | Save document to file |
| `doc.to_bytes()` | Save document to bytes |
| `doc.paragraphs()` / `paragraphs_mut()` | Iterate over paragraphs |
| `doc.paragraph(i)` / `paragraph_mut(i)` | Get paragraph by index |
| `doc.add_paragraph(text)` | Add a new paragraph |
| `doc.insert_paragraph(i, para)` | Insert at position |
| `doc.remove_paragraph(i)` | Remove a paragraph |
| `doc.tables()` / `table_mut(i)` | Access tables |
| `doc.add_table(table)` | Add a table |
| `doc.text()` | Get all document text |
| `doc.replace_text(find, replace)` | Find and replace text |
| `doc.find_text(needle)` | Find text locations |
| `doc.styles()` / `styles_mut()` | Access style definitions |
| `doc.core_properties()` / `core_properties_mut()` | Document metadata |
| `doc.section_properties()` / `section_properties_mut()` | Page layout |
| `doc.headers()` / `footers()` | Access headers/footers |
| `doc.footnotes()` / `footnotes_mut()` | Access footnotes |
| `doc.endnotes()` / `endnotes_mut()` | Access endnotes |

### Paragraph

| Method | Description |
|--------|-------------|
| `para.text()` | Get paragraph text |
| `para.set_text(text)` | Replace all content |
| `para.style()` / `set_style(name)` | Get/set style |
| `para.alignment()` / `set_alignment(align)` | Get/set alignment |
| `para.set_indentation(indent)` | Set indentation |
| `para.set_spacing(spacing)` | Set line spacing |
| `para.is_heading()` / `heading_level()` | Heading detection |
| `para.runs()` / `runs_mut()` | Access runs |
| `para.add_run(run)` | Add a text run |
| `para.add_hyperlink(r_id, text)` | Add hyperlink |
| `para.add_bookmark(id, name)` | Add bookmark |
| `para.is_list_item()` / `list_level()` | List detection |
| `para.set_numbering(num_id, level)` | Make list item |

### Run (Text with Formatting)

| Method | Description |
|--------|-------------|
| `Run::new(text)` | Create a new run |
| `run.text()` / `set_text(text)` | Get/set text |
| `run.bold()` / `set_bold(bool)` | Bold |
| `run.italic()` / `set_italic(bool)` | Italic |
| `run.underline()` / `set_underline(type)` | Underline |
| `run.strike()` / `set_strike(bool)` | Strikethrough |
| `run.font_size_pt()` / `set_font_size_pt(f32)` | Font size |
| `run.color()` / `set_color(hex)` | Text color |
| `run.font()` / `set_font(name)` | Font family |
| `run.highlight()` / `set_highlight(color)` | Highlight |
| `run.set_superscript()` / `set_subscript()` | Super/subscript |
| `run.clear_formatting()` | Remove all formatting |

### Units

```rust
use linch_docx_rs::{Pt, Twip, Cm, Inch, Emu};

let pt = Pt(12.0);
let twip = Twip::from(pt);    // 240 twips
let cm = Cm::from(pt);        // ~0.42 cm
let inch = Inch::from(pt);    // ~0.17 inches
```

## Architecture

```
┌─────────────────────────────────────────┐
│            Public API Layer             │
│  Document, Paragraph, Run, Table, ...   │
├─────────────────────────────────────────┤
│           Document Layer                │
│  Body, Styles, Properties, Section,     │
│  HeaderFooter, Footnotes                │
├─────────────────────────────────────────┤
│             OPC Layer                   │
│  Package, Part, Relationships           │
├─────────────────────────────────────────┤
│             XML Layer                   │
│  RawXmlNode, Namespace helpers          │
└─────────────────────────────────────────┘
```

## Roadmap

- [x] Basic document reading & writing
- [x] Paragraph and Run support with full formatting
- [x] Table support (read, create, modify)
- [x] Round-trip preservation
- [x] Lists and numbering
- [x] Styles management
- [x] Core properties (title, author, dates)
- [x] Section properties (page size, margins, orientation)
- [x] Headers and footers
- [x] Footnotes and endnotes
- [x] Hyperlinks and bookmarks
- [x] Text find and replace
- [x] Measurement units (Pt, Twip, Emu, Cm, Mm, Inch)
- [ ] Images and drawings
- [ ] Comments and track changes

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see [LICENSE](LICENSE) for details.

## Acknowledgments

- Inspired by [python-docx](https://python-docx.readthedocs.io/)
- Built with [quick-xml](https://github.com/tafia/quick-xml) and [zip](https://github.com/zip-rs/zip)
