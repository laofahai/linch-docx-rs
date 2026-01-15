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
        println!("Style: {:?}", para.style());
        println!("Text: {}", para.text());

        // Access runs (text with formatting)
        for run in para.runs() {
            if run.bold() {
                print!("[BOLD] ");
            }
            println!("{}", run.text());
        }
    }

    Ok(())
}
```

### Creating a Document

```rust
use linch_docx_rs::{Document, Run};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = Document::new();

    // Add a heading
    let heading = doc.add_paragraph("My Document");
    heading.set_style("Heading1");

    // Add a paragraph with formatted text
    let para = doc.add_empty_paragraph();

    let mut bold_run = Run::new("Bold text");
    bold_run.set_bold(true);
    para.add_run(bold_run);

    let mut italic_run = Run::new(" and italic text");
    italic_run.set_italic(true);
    para.add_run(italic_run);

    // Save
    doc.save("output.docx")?;

    Ok(())
}
```

### Modifying an Existing Document

```rust
use linch_docx_rs::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = Document::open("existing.docx")?;

    // Add new content
    doc.add_paragraph("This paragraph was added by Rust!");

    // Save - unknown elements are preserved!
    doc.save("modified.docx")?;

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
| `doc.paragraphs()` | Iterate over paragraphs |
| `doc.tables()` | Iterate over tables |
| `doc.add_paragraph(text)` | Add a new paragraph |
| `doc.text()` | Get all document text |

### Paragraph

| Method | Description |
|--------|-------------|
| `para.text()` | Get paragraph text |
| `para.style()` | Get style name |
| `para.is_heading()` | Check if heading |
| `para.runs()` | Iterate over runs |
| `para.add_run(run)` | Add a text run |
| `para.set_style(name)` | Set paragraph style |

### Run (Text with Formatting)

| Method | Description |
|--------|-------------|
| `Run::new(text)` | Create a new run |
| `run.text()` | Get run text |
| `run.bold()` | Check if bold |
| `run.italic()` | Check if italic |
| `run.set_bold(bool)` | Set bold |
| `run.set_italic(bool)` | Set italic |
| `run.set_font_size_pt(size)` | Set font size |
| `run.set_color(hex)` | Set text color |

### Table

| Method | Description |
|--------|-------------|
| `table.rows()` | Iterate over rows |
| `table.row_count()` | Get row count |
| `table.column_count()` | Get column count |
| `table.cell(row, col)` | Get cell at position |

## Architecture

```
┌─────────────────────────────────────────┐
│            Public API Layer             │
│   Document, Paragraph, Run, Table       │
├─────────────────────────────────────────┤
│           Document Layer                │
│   Body, BlockContent, Properties        │
├─────────────────────────────────────────┤
│             OPC Layer                   │
│   Package, Part, Relationships          │
├─────────────────────────────────────────┤
│             XML Layer                   │
│   RawXmlNode, Namespace helpers         │
└─────────────────────────────────────────┘
```

## Round-trip Preservation

Unlike many DOCX libraries that lose formatting and unknown elements, `linch-docx-rs` preserves everything:

```rust
// Original document has custom XML, tracked changes, comments, etc.
let mut doc = Document::open("complex.docx")?;

// Make your changes
doc.add_paragraph("New content");

// Save - all original elements are preserved!
doc.save("output.docx")?;
```

This is achieved by storing unknown XML elements as `RawXmlNode` during parsing and serializing them back unchanged.

## Roadmap

- [x] Basic document reading
- [x] Basic document writing
- [x] Paragraph and Run support
- [x] Table support (read)
- [x] Round-trip preservation
- [ ] Images and drawings
- [ ] Headers and footers
- [ ] Styles management
- [ ] Table creation/modification
- [ ] Lists and numbering
- [ ] Comments and track changes

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see [LICENSE](LICENSE) for details.

## Acknowledgments

- Inspired by [python-docx](https://python-docx.readthedocs.io/)
- Built with [quick-xml](https://github.com/tafia/quick-xml) and [zip](https://github.com/zip-rs/zip)
