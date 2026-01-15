# linch-docx-rs Round-trip 保真策略

## 1. 什么是 Round-trip 保真

**Round-trip** 指：打开文档 → 修改 → 保存，整个过程中不会丢失或破坏原文档中的任何信息。

**保真要求**：
- 未修改的内容完全保持原样
- 未识别的 XML 元素、属性原样保留
- 未解析的 Part 直接复制
- 保持 XML 的格式、顺序、命名空间前缀

**为什么重要**：
- OOXML 规范庞大（6000+ 页），无法实现所有功能
- 不同软件（Word、WPS、LibreOffice）会添加私有扩展
- 用户期望"打开-保存"不会破坏文档

## 2. 竞品分析

| 库 | Round-trip 支持 | 问题 |
|---|---|---|
| docx-rs | ❌ 不支持 | 只能创建，无法保持原文档结构 |
| docx-rust | ⚠️ 部分 | 丢失未识别元素 |
| python-docx | ⚠️ 部分 | 保留 Part，但丢失元素级细节 |
| Open XML SDK | ✅ 良好 | 微软官方，完整支持 |

## 3. 设计原则

### 3.1 三级保留策略

```
┌─────────────────────────────────────────┐
│         Level 1: Part 级保留             │
│  未访问的 Part 直接保留原始字节           │
├─────────────────────────────────────────┤
│         Level 2: 元素级保留              │
│  已解析元素保留未识别的子元素/属性        │
├─────────────────────────────────────────┤
│         Level 3: 格式级保留              │
│  保持 XML 格式细节（空白、顺序等）        │
└─────────────────────────────────────────┘
```

### 3.2 核心思想：只解析需要的，其余原样保留

```rust
// 每个元素都携带 "未知部分"
struct ParagraphElement {
    // 已解析的部分 - 可以读写
    properties: Option<ParagraphProperties>,
    runs: Vec<RunElement>,

    // 未识别的部分 - 原样保留
    unknown_children: Vec<RawXmlNode>,
    unknown_attributes: Vec<(String, String)>,

    // 原始 XML（可选，用于格式级保真）
    raw_xml: Option<String>,
}
```

## 4. 实现细节

### 4.1 Part 级保留

```rust
/// Part 数据
pub enum PartData {
    /// 未解析的原始字节
    Raw {
        bytes: Vec<u8>,
        /// 是否已被访问过
        accessed: bool,
    },

    /// 已解析的 XML
    Parsed {
        document: XmlDocument,
        /// 原始字节（用于未修改时的快速写入）
        original_bytes: Option<Vec<u8>>,
        /// 是否已修改
        modified: bool,
    },

    /// 二进制数据（图片等）
    Binary(Vec<u8>),
}

impl Part {
    /// 获取数据，标记为已访问
    pub fn data(&mut self) -> &[u8] {
        if let PartData::Raw { accessed, bytes } = &mut self.data {
            *accessed = true;
            bytes
        } else {
            self.data.as_bytes()
        }
    }

    /// 保存时的策略
    pub fn write_to(&self, writer: &mut impl Write) -> Result<()> {
        match &self.data {
            // 未访问过的 Part 直接写原始字节
            PartData::Raw { bytes, accessed: false } => {
                writer.write_all(bytes)?;
            }
            // 已解析但未修改的 Part 写原始字节
            PartData::Parsed { original_bytes: Some(bytes), modified: false, .. } => {
                writer.write_all(bytes)?;
            }
            // 已修改的 Part 重新序列化
            PartData::Parsed { document, .. } => {
                document.write_to(writer)?;
            }
            // 其他情况
            _ => {
                writer.write_all(self.data.as_bytes())?;
            }
        }
        Ok(())
    }
}
```

### 4.2 元素级保留

```rust
/// 原始 XML 节点
#[derive(Clone, Debug)]
pub enum RawXmlNode {
    /// 元素
    Element(RawXmlElement),
    /// 文本
    Text(String),
    /// 注释
    Comment(String),
    /// CDATA
    CData(String),
    /// 处理指令
    ProcessingInstruction { target: String, content: String },
}

/// 原始 XML 元素
#[derive(Clone, Debug)]
pub struct RawXmlElement {
    /// 完整的元素名（包含前缀）
    pub name: String,
    /// 命名空间 URI
    pub namespace: Option<String>,
    /// 属性列表（保持原始顺序）
    pub attributes: Vec<RawXmlAttribute>,
    /// 子节点
    pub children: Vec<RawXmlNode>,
    /// 是否为空元素（<foo/> vs <foo></foo>）
    pub self_closing: bool,
}

/// 原始 XML 属性
#[derive(Clone, Debug)]
pub struct RawXmlAttribute {
    /// 完整的属性名（包含前缀）
    pub name: String,
    /// 属性值
    pub value: String,
    /// 命名空间 URI
    pub namespace: Option<String>,
}
```

### 4.3 解析时保留未知部分

```rust
impl ParagraphElement {
    pub fn from_xml(reader: &mut XmlReader) -> Result<Self> {
        let mut element = Self::default();
        let start_tag = reader.current_start()?;

        // 保留未识别的属性
        for attr in start_tag.attributes() {
            let attr = attr?;
            let name = std::str::from_utf8(attr.key.as_ref())?;
            let value = std::str::from_utf8(&attr.value)?;

            match name {
                "w:rsidR" => element.rsid_r = Some(value.to_string()),
                "w:rsidRDefault" => element.rsid_r_default = Some(value.to_string()),
                "w:rsidP" => element.rsid_p = Some(value.to_string()),
                // ... 其他已知属性
                _ => {
                    // ⭐ 保留未识别的属性
                    element.unknown_attributes.push((name.to_string(), value.to_string()));
                }
            }
        }

        // 解析子元素
        loop {
            match reader.read_event()? {
                Event::Start(e) | Event::Empty(e) => {
                    let name = std::str::from_utf8(e.name().as_ref())?;

                    match name {
                        "w:pPr" => {
                            element.properties = Some(ParagraphProperties::from_xml(reader)?);
                        }
                        "w:r" => {
                            element.content.push(
                                ParagraphContent::Run(RunElement::from_xml(reader)?)
                            );
                        }
                        "w:hyperlink" => {
                            element.content.push(
                                ParagraphContent::Hyperlink(HyperlinkElement::from_xml(reader)?)
                            );
                        }
                        "w:bookmarkStart" => {
                            element.content.push(
                                ParagraphContent::BookmarkStart(BookmarkStart::from_xml(&e)?)
                            );
                        }
                        // ... 其他已知元素
                        _ => {
                            // ⭐ 保留未识别的元素
                            let raw = RawXmlElement::from_reader(reader, &e)?;
                            element.unknown_children.push(RawXmlNode::Element(raw));
                        }
                    }
                }
                Event::Text(t) => {
                    // 段落内的文本通常在 Run 中，这里保留意外的文本
                    let text = t.unescape()?.to_string();
                    if !text.trim().is_empty() {
                        element.unknown_children.push(RawXmlNode::Text(text));
                    }
                }
                Event::End(_) => break,
                Event::Eof => return Err(Error::UnexpectedEof),
                _ => {}
            }
        }

        Ok(element)
    }
}

impl RawXmlElement {
    /// 从 reader 读取完整元素（包括所有子元素）
    pub fn from_reader(reader: &mut XmlReader, start: &BytesStart) -> Result<Self> {
        let name = std::str::from_utf8(start.name().as_ref())?.to_string();

        let attributes = start.attributes()
            .map(|attr| {
                let attr = attr?;
                Ok(RawXmlAttribute {
                    name: std::str::from_utf8(attr.key.as_ref())?.to_string(),
                    value: std::str::from_utf8(&attr.value)?.to_string(),
                    namespace: None, // TODO: 解析命名空间
                })
            })
            .collect::<Result<Vec<_>>>()?;

        let mut children = Vec::new();
        let mut depth = 1;

        loop {
            match reader.read_event()? {
                Event::Start(e) => {
                    depth += 1;
                    children.push(RawXmlNode::Element(
                        Self::from_reader(reader, &e)?
                    ));
                    depth -= 1;
                }
                Event::Empty(e) => {
                    let elem = Self {
                        name: std::str::from_utf8(e.name().as_ref())?.to_string(),
                        attributes: e.attributes()
                            .map(|a| {
                                let a = a?;
                                Ok(RawXmlAttribute {
                                    name: std::str::from_utf8(a.key.as_ref())?.to_string(),
                                    value: std::str::from_utf8(&a.value)?.to_string(),
                                    namespace: None,
                                })
                            })
                            .collect::<Result<Vec<_>>>()?,
                        children: Vec::new(),
                        namespace: None,
                        self_closing: true,
                    };
                    children.push(RawXmlNode::Element(elem));
                }
                Event::Text(t) => {
                    children.push(RawXmlNode::Text(t.unescape()?.to_string()));
                }
                Event::Comment(c) => {
                    children.push(RawXmlNode::Comment(
                        std::str::from_utf8(&c)?.to_string()
                    ));
                }
                Event::CData(c) => {
                    children.push(RawXmlNode::CData(
                        std::str::from_utf8(&c)?.to_string()
                    ));
                }
                Event::End(_) => {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
                Event::Eof => return Err(Error::UnexpectedEof),
                _ => {}
            }
        }

        Ok(Self {
            name,
            namespace: None, // TODO
            attributes,
            children,
            self_closing: false,
        })
    }
}
```

### 4.4 序列化时还原未知部分

```rust
impl ParagraphElement {
    pub fn to_xml(&self, writer: &mut XmlWriter) -> Result<()> {
        // 开始标签
        let mut start = BytesStart::new("w:p");

        // 写入已知属性
        if let Some(ref rsid) = self.rsid_r {
            start.push_attribute(("w:rsidR", rsid.as_str()));
        }
        if let Some(ref rsid) = self.rsid_r_default {
            start.push_attribute(("w:rsidRDefault", rsid.as_str()));
        }

        // ⭐ 写入未识别的属性
        for (name, value) in &self.unknown_attributes {
            start.push_attribute((name.as_str(), value.as_str()));
        }

        writer.write_event(Event::Start(start))?;

        // 写入段落属性
        if let Some(ref props) = self.properties {
            props.to_xml(writer)?;
        }

        // 写入内容（保持顺序）
        let mut unknown_iter = self.unknown_children.iter().peekable();

        for content in &self.content {
            // 写入内容之前，先写入位置靠前的未知元素
            // （这里简化处理，实际可能需要更精确的位置追踪）

            match content {
                ParagraphContent::Run(run) => run.to_xml(writer)?,
                ParagraphContent::Hyperlink(link) => link.to_xml(writer)?,
                ParagraphContent::BookmarkStart(bm) => bm.to_xml(writer)?,
                // ... 其他类型
            }
        }

        // ⭐ 写入未识别的子元素
        for unknown in &self.unknown_children {
            unknown.to_xml(writer)?;
        }

        // 结束标签
        writer.write_event(Event::End(BytesEnd::new("w:p")))?;

        Ok(())
    }
}

impl RawXmlNode {
    pub fn to_xml(&self, writer: &mut XmlWriter) -> Result<()> {
        match self {
            RawXmlNode::Element(elem) => elem.to_xml(writer),
            RawXmlNode::Text(text) => {
                writer.write_event(Event::Text(BytesText::new(text)))?;
                Ok(())
            }
            RawXmlNode::Comment(comment) => {
                writer.write_event(Event::Comment(BytesText::new(comment)))?;
                Ok(())
            }
            RawXmlNode::CData(cdata) => {
                writer.write_event(Event::CData(BytesCData::new(cdata)))?;
                Ok(())
            }
            RawXmlNode::ProcessingInstruction { target, content } => {
                writer.write_event(Event::PI(BytesPI::new(
                    format!("{} {}", target, content)
                )))?;
                Ok(())
            }
        }
    }
}

impl RawXmlElement {
    pub fn to_xml(&self, writer: &mut XmlWriter) -> Result<()> {
        if self.children.is_empty() && self.self_closing {
            // 空元素
            let mut elem = BytesStart::new(&self.name);
            for attr in &self.attributes {
                elem.push_attribute((attr.name.as_str(), attr.value.as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        } else {
            // 有内容的元素
            let mut start = BytesStart::new(&self.name);
            for attr in &self.attributes {
                start.push_attribute((attr.name.as_str(), attr.value.as_str()));
            }
            writer.write_event(Event::Start(start))?;

            for child in &self.children {
                child.to_xml(writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new(&self.name)))?;
        }

        Ok(())
    }
}
```

### 4.5 元素顺序保持

OOXML 对某些元素的顺序有严格要求，需要保持：

```rust
/// 段落内容，使用枚举保持插入顺序
struct ParagraphElement {
    properties: Option<ParagraphProperties>,  // 总是第一个

    /// 内容按原始顺序存储
    /// 包括 Run、书签、批注标记、未知元素等
    content: Vec<ParagraphContentItem>,
}

/// 段落内容项（保持顺序的关键）
enum ParagraphContentItem {
    Run(RunElement),
    Hyperlink(HyperlinkElement),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    CommentRangeStart(CommentRangeStart),
    CommentRangeEnd(CommentRangeEnd),
    /// 未识别的元素也参与排序
    Unknown(RawXmlNode),
}
```

### 4.6 命名空间处理

```rust
/// 命名空间上下文
pub struct NamespaceContext {
    /// 前缀到 URI 的映射
    prefix_to_uri: HashMap<String, String>,
    /// URI 到前缀的映射（用于序列化时查找）
    uri_to_prefix: HashMap<String, String>,
    /// 默认命名空间
    default_namespace: Option<String>,
}

impl NamespaceContext {
    /// 从根元素提取命名空间声明
    pub fn from_element(element: &BytesStart) -> Self {
        let mut ctx = Self::new();

        for attr in element.attributes() {
            if let Ok(attr) = attr {
                let name = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                let value = std::str::from_utf8(&attr.value).unwrap_or("");

                if name == "xmlns" {
                    ctx.default_namespace = Some(value.to_string());
                } else if name.starts_with("xmlns:") {
                    let prefix = &name[6..];
                    ctx.add(prefix, value);
                }
            }
        }

        ctx
    }

    /// 解析带前缀的名称
    pub fn resolve(&self, prefixed_name: &str) -> (Option<&str>, &str) {
        if let Some(pos) = prefixed_name.find(':') {
            let prefix = &prefixed_name[..pos];
            let local = &prefixed_name[pos + 1..];
            (self.prefix_to_uri.get(prefix).map(|s| s.as_str()), local)
        } else {
            (self.default_namespace.as_deref(), prefixed_name)
        }
    }
}
```

## 5. 修改跟踪

### 5.1 修改标记

```rust
/// 可修改的元素包装
pub struct ModifiableElement<T> {
    element: T,
    modified: bool,
}

impl<T> ModifiableElement<T> {
    pub fn new(element: T) -> Self {
        Self { element, modified: false }
    }

    pub fn get(&self) -> &T {
        &self.element
    }

    pub fn get_mut(&mut self) -> &mut T {
        self.modified = true;
        &mut self.element
    }

    pub fn is_modified(&self) -> bool {
        self.modified
    }

    pub fn into_inner(self) -> T {
        self.element
    }
}
```

### 5.2 惰性解析与修改

```rust
/// 惰性解析的段落
pub enum LazyParagraph {
    /// 未解析，保持原始 XML
    Raw(RawXmlElement),
    /// 已解析
    Parsed(ParagraphElement),
}

impl LazyParagraph {
    /// 获取文本（触发解析）
    pub fn text(&mut self) -> String {
        self.ensure_parsed();
        match self {
            LazyParagraph::Parsed(p) => p.text(),
            _ => unreachable!(),
        }
    }

    /// 确保已解析
    fn ensure_parsed(&mut self) {
        if let LazyParagraph::Raw(raw) = self {
            let parsed = ParagraphElement::from_raw(raw);
            *self = LazyParagraph::Parsed(parsed);
        }
    }

    /// 序列化
    pub fn to_xml(&self, writer: &mut XmlWriter) -> Result<()> {
        match self {
            // 未解析的直接写原始 XML
            LazyParagraph::Raw(raw) => raw.to_xml(writer),
            // 已解析的正常序列化
            LazyParagraph::Parsed(p) => p.to_xml(writer),
        }
    }
}
```

## 6. 测试策略

### 6.1 Round-trip 测试

```rust
#[cfg(test)]
mod roundtrip_tests {
    use super::*;

    /// 基本 round-trip 测试
    #[test]
    fn test_roundtrip_simple() {
        let original = std::fs::read("tests/fixtures/simple.docx").unwrap();

        // 打开
        let doc = Document::from_bytes(&original).unwrap();

        // 直接保存（不修改）
        let saved = doc.to_bytes().unwrap();

        // 验证关键结构相同
        let doc2 = Document::from_bytes(&saved).unwrap();
        assert_eq!(doc.paragraph_count(), doc2.paragraph_count());
    }

    /// 字节级 round-trip（未修改的部分应该完全相同）
    #[test]
    fn test_roundtrip_byte_level() {
        let original = std::fs::read("tests/fixtures/simple.docx").unwrap();

        let doc = Document::from_bytes(&original).unwrap();
        let saved = doc.to_bytes().unwrap();

        // 解压两个 ZIP，比较未访问的 Part
        let original_zip = ZipArchive::new(Cursor::new(&original)).unwrap();
        let saved_zip = ZipArchive::new(Cursor::new(&saved)).unwrap();

        // 比较文件列表
        assert_eq!(original_zip.len(), saved_zip.len());

        // 比较每个文件（未访问的应该相同）
        // ...
    }

    /// 修改后 round-trip
    #[test]
    fn test_roundtrip_after_modification() {
        let original = std::fs::read("tests/fixtures/simple.docx").unwrap();

        let mut doc = Document::from_bytes(&original).unwrap();

        // 修改第一段
        if let Some(mut para) = doc.paragraph_mut(0) {
            para.set_text("Modified");
        }

        let saved = doc.to_bytes().unwrap();

        // 重新打开，验证修改保留
        let doc2 = Document::from_bytes(&saved).unwrap();
        assert_eq!(doc2.paragraph(0).unwrap().text(), "Modified");

        // 验证其他段落不变
        // ...
    }

    /// 复杂文档 round-trip
    #[test]
    fn test_roundtrip_complex() {
        let test_files = [
            "tests/fixtures/word_2016.docx",
            "tests/fixtures/word_2019.docx",
            "tests/fixtures/wps.docx",
            "tests/fixtures/libreoffice.docx",
            "tests/fixtures/google_docs.docx",
        ];

        for file in &test_files {
            if !std::path::Path::new(file).exists() {
                continue;
            }

            let original = std::fs::read(file).unwrap();
            let doc = Document::from_bytes(&original).unwrap();
            let saved = doc.to_bytes().unwrap();

            // 验证可以被 Word 打开（手动测试）
            // 这里只验证我们自己能重新打开
            let _ = Document::from_bytes(&saved).unwrap();
        }
    }

    /// 未知元素保留测试
    #[test]
    fn test_unknown_elements_preserved() {
        // 创建包含自定义元素的测试文档
        let xml = r#"
        <w:p xmlns:w="...">
            <w:pPr/>
            <w:customElement foo="bar">
                <w:nested>content</w:nested>
            </w:customElement>
            <w:r><w:t>text</w:t></w:r>
        </w:p>
        "#;

        let para = ParagraphElement::from_xml_str(xml).unwrap();

        // 验证未知元素被保留
        assert_eq!(para.unknown_children.len(), 1);

        // 序列化
        let output = para.to_xml_string().unwrap();

        // 验证自定义元素存在
        assert!(output.contains("customElement"));
        assert!(output.contains("foo=\"bar\""));
        assert!(output.contains("nested"));
    }
}
```

### 6.2 XML diff 工具

```rust
/// XML 差异比较（忽略格式差异）
pub fn xml_diff(a: &str, b: &str) -> Vec<XmlDifference> {
    let doc_a = parse_xml(a);
    let doc_b = parse_xml(b);

    compare_elements(&doc_a.root, &doc_b.root)
}

pub enum XmlDifference {
    /// 元素缺失
    MissingElement { path: String, element: String },
    /// 元素多余
    ExtraElement { path: String, element: String },
    /// 属性不同
    AttributeDiff { path: String, attr: String, expected: String, actual: String },
    /// 文本不同
    TextDiff { path: String, expected: String, actual: String },
    /// 顺序不同
    OrderDiff { path: String },
}
```

## 7. 边界情况处理

### 7.1 损坏的 XML

```rust
impl ParagraphElement {
    pub fn from_xml_lenient(reader: &mut XmlReader) -> Result<Self> {
        match Self::from_xml(reader) {
            Ok(elem) => Ok(elem),
            Err(e) => {
                // 解析失败时，保留原始 XML
                log::warn!("Failed to parse paragraph: {}, preserving raw XML", e);
                Ok(Self {
                    unknown_children: vec![/* 原始 XML */],
                    ..Default::default()
                })
            }
        }
    }
}
```

### 7.2 编码问题

```rust
/// 安全的 XML 文本编码
pub fn encode_xml_text(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            '\'' => result.push_str("&apos;"),
            // 保留其他有效的 XML 字符
            c if is_valid_xml_char(c) => result.push(c),
            // 无效字符替换为空格
            _ => result.push(' '),
        }
    }
    result
}

fn is_valid_xml_char(c: char) -> bool {
    matches!(c,
        '\u{09}' | '\u{0A}' | '\u{0D}' |
        '\u{20}'..='\u{D7FF}' |
        '\u{E000}'..='\u{FFFD}' |
        '\u{10000}'..='\u{10FFFF}'
    )
}
```

## 8. 性能考量

### 8.1 内存优化

```rust
/// 大文件使用 Cow 减少复制
pub struct TextElement<'a> {
    pub text: Cow<'a, str>,
    pub preserve_space: bool,
}

/// 未修改的原始数据使用引用
pub enum PartData<'a> {
    Raw(&'a [u8]),
    Parsed(XmlDocument),
}
```

### 8.2 增量序列化

```rust
impl Document {
    /// 只重新序列化修改过的部分
    pub fn save_incremental(&self, path: &str) -> Result<()> {
        // 复制原始 ZIP
        let mut zip = ZipWriter::new_append(/* 原始文件 */)?;

        // 只更新修改过的 Part
        for (uri, part) in &self.modified_parts {
            zip.start_file(uri.as_str(), /* options */)?;
            part.write_to(&mut zip)?;
        }

        zip.finish()?;
        Ok(())
    }
}
```

## 9. 总结

Round-trip 保真是 linch-docx-rs 的核心竞争力，通过：

1. **三级保留策略**：Part 级、元素级、格式级
2. **未知元素保留**：每个元素携带 `unknown_children` 和 `unknown_attributes`
3. **惰性解析**：未访问的内容保持原始字节
4. **修改跟踪**：只重新序列化修改过的部分
5. **全面测试**：多源文档测试、字节级比较

这确保用户可以放心使用库修改文档，不会意外破坏原有内容。
