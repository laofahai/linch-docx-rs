//! # linch-docx-rs
//!
//! A reliable DOCX reading and writing library for Rust.
//!
//! ## Features
//!
//! - Read and write DOCX files
//! - Round-trip preservation (unknown elements are kept intact)
//! - Simple, pythonic API inspired by python-docx
//!
//! ## Quick Start
//!
//! ```rust,ignore
//! use linch_docx_rs::Document;
//!
//! // Open a document
//! let doc = Document::open("example.docx")?;
//!
//! // Read paragraphs
//! for para in doc.paragraphs() {
//!     println!("{}", para.text());
//! }
//!
//! // Create a new document
//! let mut doc = Document::new();
//! doc.add_paragraph("Hello World!");
//! doc.save("output.docx")?;
//! ```

pub mod document;
pub mod error;
pub mod opc;
pub mod xml;

pub use document::{Document, Paragraph, Run, Table};
pub use error::{Error, Result};
pub use opc::{Package, Part, PartUri};
