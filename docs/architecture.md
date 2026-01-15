# linch-docx-rs 整体架构设计

## 1. 项目定位

**目标**：成为 Rust 生态中最可靠、最完整的 DOCX 读写库。

**核心原则**：
- **可靠性优先**：正确性 > 性能 > 功能丰富度
- **Round-trip 保真**：未识别的元素原样保留，不丢失数据
- **API 简洁**：常用操作一行代码，复杂操作也不繁琐
- **渐进式解析**：只解析需要的部分，按需加载

## 2. 架构总览

```
┌─────────────────────────────────────────────────────────────────┐
│                        Public API Layer                          │
│                                                                   │
│  Document  Paragraph  Run  Table  Style  Image  Section  ...    │
│                                                                   │
│  用户直接交互的高层 API，参考 python-docx 风格                      │
├─────────────────────────────────────────────────────────────────┤
│                        Parts Layer                               │
│                                                                   │
│  DocumentPart  StylesPart  NumberingPart  HeaderPart  ...       │
│                                                                   │
│  对应 OOXML 包中的各个 Part，管理 Part 间的关系                     │
├─────────────────────────────────────────────────────────────────┤
│                        Elements Layer                            │
│                                                                   │
│  ParagraphElement  RunElement  TableElement  PropertiesElement   │
│                                                                   │
│  XML 元素的 Rust 表示，包含原始数据 + 解析后的结构                   │
├─────────────────────────────────────────────────────────────────┤
│                        OPC Layer                                 │
│                                                                   │
│  Package  Part  ContentTypes  Relationships  PartUri            │
│                                                                   │
│  Open Packaging Convention 实现，处理 ZIP 和包结构                 │
├─────────────────────────────────────────────────────────────────┤
│                        XML Layer                                 │
│                                                                   │
│  XmlReader  XmlWriter  Namespace  Serialization                 │
│                                                                   │
│  底层 XML 处理，基于 quick-xml                                     │
└─────────────────────────────────────────────────────────────────┘
```

## 3. 各层职责

### 3.1 XML Layer（XML 层）

**职责**：
- XML 读取与写入的底层封装
- 命名空间管理
- XML 元素的序列化/反序列化基础设施

**关键设计**：
```rust
/// XML 命名空间定义
pub mod namespace {
    pub const WORDPROCESSING_ML: &str =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    pub const RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    pub const CONTENT_TYPES: &str =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    // ... 更多命名空间
}

/// 通用 XML 元素，用于保留未识别的元素
pub struct RawXmlElement {
    pub name: String,
    pub namespace: Option<String>,
    pub attributes: Vec<(String, String)>,
    pub children: Vec<RawXmlNode>,
}

pub enum RawXmlNode {
    Element(RawXmlElement),
    Text(String),
}
```

### 3.2 OPC Layer（OPC 层）

**职责**：
- ZIP 包的读取与写入
- Part（部件）的管理
- Content Types 解析
- Relationships（关系）解析与维护
- Part URI 路径解析

**关键类型**：
```rust
/// OPC 包
pub struct Package {
    parts: HashMap<PartUri, Part>,
    relationships: Relationships,
    content_types: ContentTypes,
}

/// 包中的部件
pub struct Part {
    uri: PartUri,
    content_type: String,
    data: PartData,
    relationships: Option<Relationships>,
}

/// Part 数据可以是原始字节或已解析的内容
pub enum PartData {
    Raw(Vec<u8>),
    Xml(ParsedXml),
    Binary(Vec<u8>),  // 图片等二进制资源
}

/// 关系定义
pub struct Relationship {
    id: String,
    rel_type: String,
    target: String,
    target_mode: TargetMode,
}
```

### 3.3 Elements Layer（元素层）

**职责**：
- OOXML 元素的 Rust 类型定义
- 元素的解析（从 XML 到 Rust 结构）
- 元素的序列化（从 Rust 结构到 XML）
- **保留未识别的子元素**（Round-trip 关键）

**关键设计 - Round-trip 保真**：
```rust
/// 段落元素
pub struct ParagraphElement {
    /// 解析后的属性（可修改）
    pub properties: Option<ParagraphProperties>,

    /// 解析后的 Run 列表
    pub runs: Vec<RunElement>,

    /// 书签、批注等其他已知元素
    pub bookmarks: Vec<BookmarkElement>,

    /// ⭐ 未识别的子元素，原样保留
    /// 这是 Round-trip 保真的关键
    pub unknown_children: Vec<RawXmlNode>,

    /// ⭐ 原始属性中未识别的部分
    pub unknown_attributes: Vec<(String, String)>,
}

/// Run 元素
pub struct RunElement {
    pub properties: Option<RunProperties>,
    pub content: Vec<RunContent>,
    pub unknown_children: Vec<RawXmlNode>,
}

/// Run 内容
pub enum RunContent {
    Text(TextElement),
    Tab,
    Break(BreakElement),
    Image(ImageElement),
    // ... 更多类型
}
```

### 3.4 Parts Layer（部件层）

**职责**：
- 管理文档的各个 Part（document.xml, styles.xml 等）
- 处理 Part 之间的关系和引用
- 提供 Part 级别的操作接口

**关键类型**：
```rust
/// 主文档部件
pub struct DocumentPart {
    pub body: BodyElement,
    pub background: Option<BackgroundElement>,
    relationships: Relationships,
}

/// 样式部件
pub struct StylesPart {
    pub doc_defaults: Option<DocDefaults>,
    pub styles: Vec<StyleElement>,
}

/// 编号部件
pub struct NumberingPart {
    pub abstract_nums: Vec<AbstractNum>,
    pub nums: Vec<Num>,
}
```

### 3.5 Public API Layer（公开 API 层）

**职责**：
- 提供简洁、符合人体工程学的 API
- 隐藏底层复杂性
- 提供迭代器、Builder 模式等便捷接口

**API 风格**：
```rust
// ===== 读取 =====
let doc = Document::open("example.docx")?;

// 遍历段落
for para in doc.paragraphs() {
    println!("{}", para.text());

    // 访问格式
    if para.style() == Some("Heading1") {
        // ...
    }

    // 遍历 runs
    for run in para.runs() {
        if run.bold() {
            println!("Bold: {}", run.text());
        }
    }
}

// 遍历表格
for table in doc.tables() {
    for row in table.rows() {
        for cell in row.cells() {
            println!("{}", cell.text());
        }
    }
}

// ===== 创建 =====
let mut doc = Document::new();

doc.add_paragraph("Title")
    .style("Heading1");

doc.add_paragraph("Normal text.")
    .bold(true)
    .font_size(Pt(12));

let table = doc.add_table(3, 3);
table.cell(0, 0).set_text("Header");

doc.save("output.docx")?;

// ===== 修改 =====
let mut doc = Document::open("template.docx")?;

doc.replace_text("{{name}}", "张三");

for para in doc.paragraphs_mut() {
    if para.text().contains("重要") {
        para.set_style("Heading2");
    }
}

doc.save("filled.docx")?;
```

## 4. 数据流

### 4.1 读取流程

```
DOCX 文件 (ZIP)
    │
    ▼
┌─────────────────┐
│   OPC Layer     │  解压 ZIP，解析 Content Types 和 Relationships
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Parts Layer    │  识别各个 Part，建立 Part 间的关系图
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Elements Layer  │  惰性解析 XML，保留未识别元素
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   API Layer     │  包装为用户友好的 API
└─────────────────┘
```

### 4.2 写入流程

```
API 调用
    │
    ▼
┌─────────────────┐
│   API Layer     │  收集修改，更新内部状态
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Elements Layer  │  序列化为 XML，合并未识别元素
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Parts Layer    │  更新 Part 内容，维护关系
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   OPC Layer     │  打包为 ZIP
└────────┬────────┘
         │
         ▼
DOCX 文件 (ZIP)
```

### 4.3 修改流程（Round-trip）

```
原始 DOCX
    │
    ▼
┌──────────────────────────────────────┐
│  读取：解析已知元素，保留未知元素       │
└────────────────┬─────────────────────┘
                 │
                 ▼
┌──────────────────────────────────────┐
│  修改：只修改用户指定的部分             │
│  未触及的部分保持原样                   │
└────────────────┬─────────────────────┘
                 │
                 ▼
┌──────────────────────────────────────┐
│  保存：已知元素序列化，未知元素原样输出  │
└────────────────┬─────────────────────┘
                 │
                 ▼
修改后的 DOCX（保留了所有未识别的内容）
```

## 5. 模块划分

```
linch-docx-rs/
├── src/
│   ├── lib.rs              # 库入口，公开 API 导出
│   │
│   ├── xml/                # XML Layer
│   │   ├── mod.rs
│   │   ├── namespace.rs    # 命名空间常量
│   │   ├── reader.rs       # XML 读取封装
│   │   ├── writer.rs       # XML 写入封装
│   │   └── raw.rs          # RawXmlElement 定义
│   │
│   ├── opc/                # OPC Layer
│   │   ├── mod.rs
│   │   ├── package.rs      # Package 结构
│   │   ├── part.rs         # Part 结构
│   │   ├── content_types.rs
│   │   ├── relationships.rs
│   │   └── part_uri.rs     # URI 解析
│   │
│   ├── elements/           # Elements Layer
│   │   ├── mod.rs
│   │   ├── document.rs     # w:document
│   │   ├── body.rs         # w:body
│   │   ├── paragraph.rs    # w:p
│   │   ├── run.rs          # w:r
│   │   ├── text.rs         # w:t
│   │   ├── table.rs        # w:tbl, w:tr, w:tc
│   │   ├── properties/     # 各种属性
│   │   │   ├── mod.rs
│   │   │   ├── paragraph.rs
│   │   │   ├── run.rs
│   │   │   └── table.rs
│   │   └── styles.rs       # 样式元素
│   │
│   ├── parts/              # Parts Layer
│   │   ├── mod.rs
│   │   ├── document.rs     # DocumentPart
│   │   ├── styles.rs       # StylesPart
│   │   ├── numbering.rs    # NumberingPart
│   │   ├── header.rs       # HeaderPart
│   │   ├── footer.rs       # FooterPart
│   │   └── media.rs        # 图片等媒体
│   │
│   ├── api/                # Public API Layer
│   │   ├── mod.rs
│   │   ├── document.rs     # Document 主结构
│   │   ├── paragraph.rs    # Paragraph API
│   │   ├── run.rs          # Run API
│   │   ├── table.rs        # Table API
│   │   ├── style.rs        # Style API
│   │   ├── image.rs        # Image API
│   │   └── builder.rs      # Builder 模式
│   │
│   ├── error.rs            # 错误类型定义
│   └── units.rs            # 单位（Pt, Emu, Twip 等）
│
├── tests/
│   ├── fixtures/           # 测试用 DOCX 文件
│   ├── read_tests.rs       # 读取测试
│   ├── write_tests.rs      # 写入测试
│   ├── roundtrip_tests.rs  # Round-trip 测试
│   └── compat_tests.rs     # 兼容性测试
│
└── docs/
    ├── architecture.md     # 本文档
    ├── api-design.md       # API 设计文档
    ├── opc-layer.md        # OPC 层详细设计
    ├── xml-mapping.md      # OOXML 元素映射
    ├── roundtrip.md        # Round-trip 策略
    └── error-handling.md   # 错误处理设计
```

## 6. 关键设计决策

### 6.1 所有权模型

**选择**：`Document` 拥有所有数据，API 返回引用或轻量包装

```rust
// Document 拥有所有数据
pub struct Document {
    package: Package,
    document_part: DocumentPart,
    styles_part: Option<StylesPart>,
    // ...
}

// Paragraph 是轻量引用包装
pub struct Paragraph<'a> {
    doc: &'a Document,
    element: &'a ParagraphElement,
}

// 可变版本
pub struct ParagraphMut<'a> {
    doc: &'a mut Document,
    index: usize,  // 用索引而非引用，避免借用冲突
}
```

**理由**：
- 避免 `Rc<RefCell<>>` 的运行时开销
- 符合 Rust 的所有权语义
- 生命周期清晰

### 6.2 惰性解析 vs 预解析

**选择**：两阶段解析

1. **打开文件时**：解析 OPC 结构和 Part 目录，不解析 XML 内容
2. **访问时**：按需解析具体 Part 的 XML 内容

```rust
impl Document {
    pub fn open(path: &str) -> Result<Self> {
        // 阶段 1：只解析 OPC 结构
        let package = Package::open(path)?;
        Ok(Self { package, parsed_parts: HashMap::new() })
    }

    pub fn paragraphs(&self) -> impl Iterator<Item = Paragraph> {
        // 阶段 2：首次访问时解析 document.xml
        self.ensure_document_parsed();
        // ...
    }
}
```

**理由**：
- 打开大文件速度快
- 只访问部分内容时节省内存
- 用户感知的性能更好

### 6.3 错误处理策略

**选择**：分层错误 + 详细上下文

```rust
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error at {location}: {message}")]
    Xml { location: String, message: String },

    #[error("Invalid DOCX: {0}")]
    InvalidDocx(String),

    #[error("Missing part: {0}")]
    MissingPart(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),
}
```

### 6.4 线程安全

**选择**：`Document` 不实现 `Sync`，但支持 `Send`

```rust
// Document 可以在线程间转移
// 但不能同时被多个线程访问
impl Send for Document {}
// 不实现 Sync
```

**理由**：
- 简化内部实现
- 大多数用例不需要并发访问同一文档
- 用户可以用 `Mutex` 包装实现并发（如果需要）

## 7. 性能考量

### 7.1 内存使用

- 大文件使用流式解析，避免一次性加载全部 XML
- 图片等二进制资源延迟加载
- 未修改的 Part 保持原始字节，不解析

### 7.2 解析性能

- 使用 `quick-xml` 的零拷贝解析
- 字符串使用 `Cow<str>` 减少分配
- 常用字符串（如命名空间）使用静态常量

### 7.3 写入性能

- 未修改的 Part 直接从原始字节写入
- 使用 `BufWriter` 减少 IO 次数
- 批量操作合并为单次写入

## 8. 扩展性设计

### 8.1 自定义元素处理

```rust
// 允许用户注册自定义元素解析器
doc.register_element_handler("w:customXml", |element| {
    // 自定义处理逻辑
});
```

### 8.2 插件机制（未来）

```rust
// 未来可能支持插件
trait DocxPlugin {
    fn on_load(&self, doc: &mut Document);
    fn on_save(&self, doc: &Document);
}
```

## 9. 与 python-docx 的对比

| 方面 | python-docx | linch-docx-rs |
|------|-------------|---------------|
| 内存模型 | GC + 弱引用 | 所有权 + 生命周期 |
| Part 关系 | 运行时动态查找 | 编译期类型保证 |
| Round-trip | 部分支持 | 完整支持 |
| 类型安全 | 动态类型 | 静态类型 |
| API 风格 | 相似 | 相似，但更 Rust 风格 |

## 10. 下一步

1. 完成 [API 设计文档](./api-design.md)
2. 完成 [OPC 层设计文档](./opc-layer.md)
3. 实现 OPC 层核心代码
4. 实现 Elements 层核心代码
5. 实现 Public API 层
