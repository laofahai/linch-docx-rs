//! Image support for DOCX documents
//!
//! Handles inline images via DrawingML (w:drawing > wp:inline > a:graphic > pic:pic).

use crate::error::Result;
use crate::xml::RawXmlNode;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Writer;

/// An inline image in the document
#[derive(Clone, Debug)]
pub struct InlineImage {
    /// Relationship ID referencing the image part
    pub r_id: String,
    /// Image width in EMU (English Metric Units, 914400 per inch)
    pub width_emu: i64,
    /// Image height in EMU
    pub height_emu: i64,
    /// Description / alt text
    pub description: String,
    /// Name
    pub name: String,
    /// The full raw XML of the drawing element (for round-trip preservation)
    pub raw_xml: Option<RawXmlNode>,
}

impl InlineImage {
    /// Create a new inline image reference
    pub fn new(r_id: impl Into<String>, width_emu: i64, height_emu: i64) -> Self {
        InlineImage {
            r_id: r_id.into(),
            width_emu,
            height_emu,
            description: String::new(),
            name: String::new(),
            raw_xml: None,
        }
    }

    /// Create with dimensions in centimeters
    pub fn from_cm(r_id: impl Into<String>, width_cm: f64, height_cm: f64) -> Self {
        // 1 cm = 360000 EMU
        Self::new(
            r_id,
            (width_cm * 360000.0) as i64,
            (height_cm * 360000.0) as i64,
        )
    }

    /// Create with dimensions in inches
    pub fn from_inches(r_id: impl Into<String>, width_in: f64, height_in: f64) -> Self {
        // 1 inch = 914400 EMU
        Self::new(
            r_id,
            (width_in * 914400.0) as i64,
            (height_in * 914400.0) as i64,
        )
    }

    /// Set alt text
    pub fn with_description(mut self, desc: impl Into<String>) -> Self {
        self.description = desc.into();
        self
    }

    /// Set name
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = name.into();
        self
    }

    /// Generate the DrawingML XML for this image
    pub fn to_drawing_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // If we have raw XML from parsing, use it for round-trip
        if let Some(ref raw) = self.raw_xml {
            raw.write_to(writer)?;
            return Ok(());
        }

        // Generate minimal inline image XML
        // <w:drawing>
        writer.write_event(Event::Start(BytesStart::new("w:drawing")))?;

        //   <wp:inline distT="0" distB="0" distL="0" distR="0">
        let mut inline = BytesStart::new("wp:inline");
        inline.push_attribute(("distT", "0"));
        inline.push_attribute(("distB", "0"));
        inline.push_attribute(("distL", "0"));
        inline.push_attribute(("distR", "0"));
        writer.write_event(Event::Start(inline))?;

        //     <wp:extent cx="..." cy="..."/>
        let mut extent = BytesStart::new("wp:extent");
        extent.push_attribute(("cx", self.width_emu.to_string().as_str()));
        extent.push_attribute(("cy", self.height_emu.to_string().as_str()));
        writer.write_event(Event::Empty(extent))?;

        //     <wp:docPr id="1" name="..." descr="..."/>
        let mut doc_pr = BytesStart::new("wp:docPr");
        doc_pr.push_attribute(("id", "1"));
        doc_pr.push_attribute(("name", self.name.as_str()));
        doc_pr.push_attribute(("descr", self.description.as_str()));
        writer.write_event(Event::Empty(doc_pr))?;

        //     <a:graphic>
        let mut graphic = BytesStart::new("a:graphic");
        graphic.push_attribute(("xmlns:a", crate::xml::A));
        writer.write_event(Event::Start(graphic))?;

        //       <a:graphicData uri="...">
        let mut gd = BytesStart::new("a:graphicData");
        gd.push_attribute(("uri", crate::xml::PIC));
        writer.write_event(Event::Start(gd))?;

        //         <pic:pic xmlns:pic="...">
        let mut pic = BytesStart::new("pic:pic");
        pic.push_attribute(("xmlns:pic", crate::xml::PIC));
        writer.write_event(Event::Start(pic))?;

        //           <pic:nvPicPr>
        writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
        let mut cnvpr = BytesStart::new("pic:cNvPr");
        cnvpr.push_attribute(("id", "0"));
        cnvpr.push_attribute(("name", self.name.as_str()));
        writer.write_event(Event::Empty(cnvpr))?;
        writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

        //           <pic:blipFill>
        writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
        let mut blip = BytesStart::new("a:blip");
        blip.push_attribute(("r:embed", self.r_id.as_str()));
        writer.write_event(Event::Empty(blip))?;
        writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
        writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

        //           <pic:spPr>
        writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
        writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", "0"));
        off.push_attribute(("y", "0"));
        writer.write_event(Event::Empty(off))?;
        let mut ext = BytesStart::new("a:ext");
        ext.push_attribute(("cx", self.width_emu.to_string().as_str()));
        ext.push_attribute(("cy", self.height_emu.to_string().as_str()));
        writer.write_event(Event::Empty(ext))?;
        writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
        let mut prst = BytesStart::new("a:prstGeom");
        prst.push_attribute(("prst", "rect"));
        writer.write_event(Event::Start(prst))?;
        writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

        //         </pic:pic>
        writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
        //       </a:graphicData>
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        //     </a:graphic>
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
        //   </wp:inline>
        writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;
        // </w:drawing>
        writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;

        Ok(())
    }
}

/// Image data to be embedded in the document
pub struct ImageData {
    /// Raw image bytes
    pub data: Vec<u8>,
    /// Content type (e.g., "image/png", "image/jpeg")
    pub content_type: String,
    /// File extension
    pub extension: String,
}

impl ImageData {
    /// Create from PNG bytes
    pub fn png(data: Vec<u8>) -> Self {
        ImageData {
            data,
            content_type: "image/png".into(),
            extension: "png".into(),
        }
    }

    /// Create from JPEG bytes
    pub fn jpeg(data: Vec<u8>) -> Self {
        ImageData {
            data,
            content_type: "image/jpeg".into(),
            extension: "jpeg".into(),
        }
    }

    /// Create from file path (auto-detects type)
    pub fn from_file(path: &std::path::Path) -> std::io::Result<Self> {
        let data = std::fs::read(path)?;
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png")
            .to_lowercase();
        let content_type = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" | "tif" => "image/tiff",
            _ => "image/png",
        };
        Ok(ImageData {
            data,
            content_type: content_type.into(),
            extension: ext,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_inline_image_new() {
        let img = InlineImage::new("rId5", 914400, 914400);
        assert_eq!(img.r_id, "rId5");
        assert_eq!(img.width_emu, 914400);
        assert_eq!(img.height_emu, 914400);
    }

    #[test]
    fn test_inline_image_from_cm() {
        let img = InlineImage::from_cm("rId1", 10.0, 5.0);
        assert_eq!(img.width_emu, 3600000);
        assert_eq!(img.height_emu, 1800000);
    }

    #[test]
    fn test_inline_image_from_inches() {
        let img = InlineImage::from_inches("rId1", 1.0, 1.0);
        assert_eq!(img.width_emu, 914400);
        assert_eq!(img.height_emu, 914400);
    }

    #[test]
    fn test_image_data_png() {
        let data = ImageData::png(vec![0x89, 0x50, 0x4E, 0x47]);
        assert_eq!(data.content_type, "image/png");
        assert_eq!(data.extension, "png");
    }

    #[test]
    fn test_generate_drawing_xml() {
        let img = InlineImage::new("rId1", 914400, 914400)
            .with_name("test.png")
            .with_description("Test image");

        let mut buf = Vec::new();
        let mut writer = Writer::new(&mut buf);
        img.to_drawing_xml(&mut writer).unwrap();

        let xml = String::from_utf8(buf).unwrap();
        assert!(xml.contains("w:drawing"));
        assert!(xml.contains("wp:inline"));
        assert!(xml.contains("r:embed=\"rId1\""));
        assert!(xml.contains("cx=\"914400\""));
    }
}
