//! Text search and replace operations for Document

use crate::document::{BlockContent, Document, Paragraph, ParagraphContent, RunContent};

/// Text location in the document
#[derive(Clone, Debug)]
pub struct TextLocation {
    pub paragraph_index: usize,
    pub char_offset: usize,
}

impl Document {
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
