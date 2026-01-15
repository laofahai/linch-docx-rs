# linch-docx-rs OPC 层设计文档

## 1. 概述

OPC (Open Packaging Conventions) 是 Office Open XML 的基础，定义了如何将文档打包为 ZIP 格式，以及各部件之间的关系。

本文档详细说明 linch-docx-rs 中 OPC 层的设计。

## 2. OPC 规范要点

### 2.1 包结构

一个 DOCX 文件本质上是一个 ZIP 压缩包，包含以下结构：

```
example.docx (ZIP)
│
├── [Content_Types].xml          # 必需：内容类型声明
│
├── _rels/
│   └── .rels                    # 必需：包级关系
│
├── docProps/
│   ├── app.xml                  # 可选：应用属性
│   └── core.xml                 # 可选：核心属性 (DC metadata)
│
└── word/
    ├── document.xml             # 必需：主文档
    ├── styles.xml               # 可选：样式定义
    ├── settings.xml             # 可选：文档设置
    ├── fontTable.xml            # 可选：字体表
    ├── numbering.xml            # 可选：编号定义
    ├── footnotes.xml            # 可选：脚注
    ├── endnotes.xml             # 可选：尾注
    ├── comments.xml             # 可选：批注
    ├── header1.xml              # 可选：页眉
    ├── footer1.xml              # 可选：页脚
    │
    ├── _rels/
    │   └── document.xml.rels    # document.xml 的关系
    │
    └── media/                   # 可选：媒体文件
        ├── image1.png
        └── image2.jpeg
```

### 2.2 内容类型 ([Content_Types].xml)

声明包中每个部件的内容类型（MIME type）。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- 默认扩展名映射 -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>

  <!-- 特定部件类型覆盖 -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <!-- ... -->
</Types>
```

### 2.3 关系文件 (.rels)

定义部件之间的关系。

**包级关系 (`/_rels/.rels`)**：
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml"/>
</Relationships>
```

**部件级关系 (`/word/_rels/document.xml.rels`)**：
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.com"
                TargetMode="External"/>
</Relationships>
```

## 3. 核心类型设计

### 3.1 Package（包）

```rust
/// OPC 包，表示一个 DOCX 文件
pub struct Package {
    /// 包中的所有部件
    parts: HashMap<PartUri, Part>,

    /// 包级关系
    relationships: Relationships,

    /// 内容类型定义
    content_types: ContentTypes,

    /// 原始 ZIP 数据（用于未修改部件的快速写入）
    source: Option<PackageSource>,
}

/// 包的来源
enum PackageSource {
    /// 从文件读取
    File(PathBuf),
    /// 从内存读取
    Memory(Vec<u8>),
}

impl Package {
    /// 从文件打开包
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self>;

    /// 从字节打开包
    pub fn from_bytes(bytes: &[u8]) -> Result<Self>;

    /// 从 Read + Seek 打开包
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self>;

    /// 创建空包
    pub fn new() -> Self;

    /// 保存到文件
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()>;

    /// 保存到字节
    pub fn to_bytes(&self) -> Result<Vec<u8>>;

    /// 保存到 Write + Seek
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()>;

    /// 获取部件
    pub fn part(&self, uri: &PartUri) -> Option<&Part>;

    /// 获取可变部件
    pub fn part_mut(&mut self, uri: &PartUri) -> Option<&mut Part>;

    /// 添加部件
    pub fn add_part(&mut self, part: Part) -> Result<()>;

    /// 移除部件
    pub fn remove_part(&mut self, uri: &PartUri) -> Option<Part>;

    /// 获取所有部件 URI
    pub fn part_uris(&self) -> impl Iterator<Item = &PartUri>;

    /// 获取包级关系
    pub fn relationships(&self) -> &Relationships;

    /// 获取可变包级关系
    pub fn relationships_mut(&mut self) -> &mut Relationships;

    /// 获取内容类型
    pub fn content_types(&self) -> &ContentTypes;

    /// 获取可变内容类型
    pub fn content_types_mut(&mut self) -> &mut ContentTypes;

    /// 通过关系类型查找部件
    pub fn part_by_relationship_type(&self, rel_type: &str) -> Option<&Part>;

    /// 获取主文档部件
    pub fn main_document_part(&self) -> Option<&Part> {
        self.part_by_relationship_type(rel_types::OFFICE_DOCUMENT)
    }
}
```

### 3.2 Part（部件）

```rust
/// OPC 部件
pub struct Part {
    /// 部件 URI
    uri: PartUri,

    /// 内容类型
    content_type: String,

    /// 部件数据
    data: PartData,

    /// 部件级关系（如果有）
    relationships: Option<Relationships>,

    /// 是否已修改
    modified: bool,
}

/// 部件数据
pub enum PartData {
    /// 原始字节（未解析）
    Raw(Vec<u8>),

    /// 已解析的 XML
    Xml(XmlData),

    /// 二进制数据（图片等）
    Binary(Vec<u8>),
}

/// XML 数据
pub struct XmlData {
    /// 解析后的文档
    pub document: XmlDocument,

    /// 原始字节（用于未修改时的快速写入）
    pub raw: Option<Vec<u8>>,
}

impl Part {
    /// 创建新部件
    pub fn new(uri: PartUri, content_type: &str, data: Vec<u8>) -> Self;

    /// 获取 URI
    pub fn uri(&self) -> &PartUri;

    /// 获取内容类型
    pub fn content_type(&self) -> &str;

    /// 获取数据（作为字节）
    pub fn data(&self) -> &[u8];

    /// 获取数据（作为字符串，如果是 XML）
    pub fn data_as_str(&self) -> Result<&str>;

    /// 设置数据
    pub fn set_data(&mut self, data: Vec<u8>);

    /// 获取关系
    pub fn relationships(&self) -> Option<&Relationships>;

    /// 获取可变关系
    pub fn relationships_mut(&mut self) -> Option<&mut Relationships>;

    /// 确保有关系对象
    pub fn ensure_relationships(&mut self) -> &mut Relationships;

    /// 是否已修改
    pub fn is_modified(&self) -> bool;

    /// 标记为已修改
    pub fn mark_modified(&mut self);

    /// 获取关系文件的 URI
    pub fn relationships_uri(&self) -> PartUri;
}
```

### 3.3 PartUri（部件 URI）

```rust
/// 部件 URI
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct PartUri {
    /// 规范化后的路径（总是以 / 开头）
    path: String,
}

impl PartUri {
    /// 从字符串创建
    pub fn new(path: &str) -> Result<Self>;

    /// 从路径组件创建
    pub fn from_parts(parts: &[&str]) -> Self;

    /// 获取路径字符串
    pub fn as_str(&self) -> &str;

    /// 获取文件名
    pub fn file_name(&self) -> Option<&str>;

    /// 获取扩展名
    pub fn extension(&self) -> Option<&str>;

    /// 获取父目录
    pub fn parent(&self) -> Option<PartUri>;

    /// 获取关系文件路径
    /// 例如：/word/document.xml -> /word/_rels/document.xml.rels
    pub fn relationships_uri(&self) -> PartUri;

    /// 解析相对路径
    /// 例如：/word/document.xml + "../media/image1.png" -> /word/media/image1.png
    pub fn resolve(&self, relative: &str) -> Result<PartUri>;

    /// 获取相对于另一个 URI 的相对路径
    pub fn relative_to(&self, base: &PartUri) -> String;

    /// 是否是关系文件
    pub fn is_relationships(&self) -> bool;
}

impl std::fmt::Display for PartUri {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.path)
    }
}

impl std::str::FromStr for PartUri {
    type Err = Error;

    fn from_str(s: &str) -> Result<Self> {
        Self::new(s)
    }
}

// 常用 URI 常量
pub mod part_uris {
    use super::PartUri;

    pub fn content_types() -> PartUri {
        PartUri::new("/[Content_Types].xml").unwrap()
    }

    pub fn package_relationships() -> PartUri {
        PartUri::new("/_rels/.rels").unwrap()
    }

    pub fn document() -> PartUri {
        PartUri::new("/word/document.xml").unwrap()
    }

    pub fn styles() -> PartUri {
        PartUri::new("/word/styles.xml").unwrap()
    }

    pub fn numbering() -> PartUri {
        PartUri::new("/word/numbering.xml").unwrap()
    }

    pub fn settings() -> PartUri {
        PartUri::new("/word/settings.xml").unwrap()
    }

    pub fn core_properties() -> PartUri {
        PartUri::new("/docProps/core.xml").unwrap()
    }

    pub fn app_properties() -> PartUri {
        PartUri::new("/docProps/app.xml").unwrap()
    }
}
```

### 3.4 ContentTypes（内容类型）

```rust
/// 内容类型定义
pub struct ContentTypes {
    /// 默认扩展名映射
    defaults: HashMap<String, String>,

    /// 特定部件覆盖
    overrides: HashMap<PartUri, String>,
}

impl ContentTypes {
    /// 创建默认内容类型
    pub fn new() -> Self {
        let mut ct = Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
        };

        // 添加默认扩展名映射
        ct.add_default("rels", content_types::RELATIONSHIPS);
        ct.add_default("xml", content_types::XML);
        ct.add_default("png", "image/png");
        ct.add_default("jpeg", "image/jpeg");
        ct.add_default("jpg", "image/jpeg");
        ct.add_default("gif", "image/gif");
        ct.add_default("bmp", "image/bmp");
        ct.add_default("tiff", "image/tiff");

        ct
    }

    /// 从 XML 解析
    pub fn from_xml(xml: &str) -> Result<Self>;

    /// 转换为 XML
    pub fn to_xml(&self) -> String;

    /// 添加默认扩展名映射
    pub fn add_default(&mut self, extension: &str, content_type: &str);

    /// 添加特定部件覆盖
    pub fn add_override(&mut self, uri: &PartUri, content_type: &str);

    /// 获取部件的内容类型
    pub fn get_content_type(&self, uri: &PartUri) -> Option<&str>;

    /// 移除覆盖
    pub fn remove_override(&mut self, uri: &PartUri);
}

/// 常用内容类型常量
pub mod content_types {
    pub const RELATIONSHIPS: &str =
        "application/vnd.openxmlformats-package.relationships+xml";
    pub const XML: &str = "application/xml";

    pub const MAIN_DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub const STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    pub const SETTINGS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    pub const NUMBERING: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    pub const FONT_TABLE: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    pub const FOOTNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    pub const ENDNOTES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    pub const HEADER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    pub const FOOTER: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    pub const COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

    pub const CORE_PROPERTIES: &str =
        "application/vnd.openxmlformats-package.core-properties+xml";
    pub const EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
}
```

### 3.5 Relationships（关系）

```rust
/// 关系集合
pub struct Relationships {
    /// 所有关系，按 ID 索引
    items: HashMap<String, Relationship>,

    /// 下一个自动生成的 ID 序号
    next_id: u32,
}

/// 单个关系
#[derive(Clone, Debug)]
pub struct Relationship {
    /// 关系 ID（如 "rId1"）
    pub id: String,

    /// 关系类型（URI）
    pub rel_type: String,

    /// 目标（相对或绝对路径）
    pub target: String,

    /// 目标模式
    pub target_mode: TargetMode,
}

/// 目标模式
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TargetMode {
    /// 内部目标（包内的其他部件）
    #[default]
    Internal,

    /// 外部目标（超链接等）
    External,
}

impl Relationships {
    /// 创建空关系集合
    pub fn new() -> Self;

    /// 从 XML 解析
    pub fn from_xml(xml: &str) -> Result<Self>;

    /// 转换为 XML
    pub fn to_xml(&self) -> String;

    /// 获取关系
    pub fn get(&self, id: &str) -> Option<&Relationship>;

    /// 按类型获取关系
    pub fn by_type(&self, rel_type: &str) -> Option<&Relationship>;

    /// 按类型获取所有关系
    pub fn all_by_type(&self, rel_type: &str) -> Vec<&Relationship>;

    /// 添加关系（自动生成 ID）
    pub fn add(&mut self, rel_type: &str, target: &str) -> String;

    /// 添加关系（指定 ID）
    pub fn add_with_id(&mut self, id: &str, rel_type: &str, target: &str);

    /// 添加外部关系
    pub fn add_external(&mut self, rel_type: &str, target: &str) -> String;

    /// 移除关系
    pub fn remove(&mut self, id: &str) -> Option<Relationship>;

    /// 按类型移除关系
    pub fn remove_by_type(&mut self, rel_type: &str) -> Vec<Relationship>;

    /// 迭代所有关系
    pub fn iter(&self) -> impl Iterator<Item = &Relationship>;

    /// 关系数量
    pub fn len(&self) -> usize;

    /// 是否为空
    pub fn is_empty(&self) -> bool;

    /// 生成新的唯一 ID
    fn generate_id(&mut self) -> String;
}

/// 常用关系类型常量
pub mod rel_types {
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
}
```

## 4. 读取流程

### 4.1 打开包

```rust
impl Package {
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        // 1. 打开 ZIP 文件
        let mut archive = ZipArchive::new(reader)?;

        // 2. 解析 [Content_Types].xml（必需）
        let content_types = Self::read_content_types(&mut archive)?;

        // 3. 解析包级关系 /_rels/.rels（必需）
        let relationships = Self::read_package_relationships(&mut archive)?;

        // 4. 读取所有部件（惰性或即时，取决于策略）
        let parts = Self::read_parts(&mut archive, &content_types)?;

        Ok(Self {
            parts,
            relationships,
            content_types,
            source: None,
        })
    }

    fn read_content_types<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<ContentTypes> {
        let mut file = archive.by_name("[Content_Types].xml")
            .map_err(|_| Error::MissingPart("[Content_Types].xml".into()))?;

        let mut content = String::new();
        file.read_to_string(&mut content)?;

        ContentTypes::from_xml(&content)
    }

    fn read_package_relationships<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<Relationships> {
        match archive.by_name("_rels/.rels") {
            Ok(mut file) => {
                let mut content = String::new();
                file.read_to_string(&mut content)?;
                Relationships::from_xml(&content)
            }
            Err(_) => {
                // 关系文件可以不存在
                Ok(Relationships::new())
            }
        }
    }

    fn read_parts<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        content_types: &ContentTypes,
    ) -> Result<HashMap<PartUri, Part>> {
        let mut parts = HashMap::new();

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();

            // 跳过目录和特殊文件
            if name.ends_with('/') || name == "[Content_Types].xml" {
                continue;
            }

            // 跳过关系文件（稍后处理）
            if name.contains("_rels/") && name.ends_with(".rels") {
                continue;
            }

            let uri = PartUri::new(&format!("/{}", name))?;

            // 获取内容类型
            let content_type = content_types
                .get_content_type(&uri)
                .unwrap_or("application/octet-stream")
                .to_string();

            // 读取数据
            let mut data = Vec::new();
            file.read_to_end(&mut data)?;

            let part = Part::new(uri.clone(), &content_type, data);
            parts.insert(uri, part);
        }

        // 读取每个部件的关系文件
        Self::read_part_relationships(archive, &mut parts)?;

        Ok(parts)
    }

    fn read_part_relationships<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        parts: &mut HashMap<PartUri, Part>,
    ) -> Result<()> {
        // 收集所有部件的关系文件路径
        let rel_paths: Vec<(PartUri, String)> = parts
            .keys()
            .map(|uri| (uri.clone(), uri.relationships_uri().as_str()[1..].to_string()))
            .collect();

        for (part_uri, rel_path) in rel_paths {
            if let Ok(mut file) = archive.by_name(&rel_path) {
                let mut content = String::new();
                file.read_to_string(&mut content)?;
                let rels = Relationships::from_xml(&content)?;

                if let Some(part) = parts.get_mut(&part_uri) {
                    part.relationships = Some(rels);
                }
            }
        }

        Ok(())
    }
}
```

### 4.2 惰性读取策略

对于大文件，可以采用惰性读取策略：

```rust
/// 惰性部件数据
pub enum LazyPartData {
    /// 未加载
    NotLoaded {
        /// ZIP 文件中的索引
        zip_index: usize,
        /// 数据大小
        size: u64,
    },

    /// 已加载
    Loaded(PartData),
}

impl Part {
    /// 确保数据已加载
    pub fn ensure_loaded(&mut self, archive: &mut ZipArchive<impl Read + Seek>) -> Result<()> {
        if let LazyPartData::NotLoaded { zip_index, .. } = &self.lazy_data {
            let mut file = archive.by_index(*zip_index)?;
            let mut data = Vec::new();
            file.read_to_end(&mut data)?;
            self.lazy_data = LazyPartData::Loaded(PartData::Raw(data));
        }
        Ok(())
    }
}
```

## 5. 写入流程

### 5.1 保存包

```rust
impl Package {
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = FileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        // 1. 写入 [Content_Types].xml
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(self.content_types.to_xml().as_bytes())?;

        // 2. 写入包级关系
        if !self.relationships.is_empty() {
            zip.start_file("_rels/.rels", options)?;
            zip.write_all(self.relationships.to_xml().as_bytes())?;
        }

        // 3. 写入所有部件
        for (uri, part) in &self.parts {
            // 写入部件数据
            let path = &uri.as_str()[1..]; // 去掉开头的 /
            zip.start_file(path, options)?;
            zip.write_all(part.data())?;

            // 写入部件关系（如果有）
            if let Some(rels) = part.relationships() {
                if !rels.is_empty() {
                    let rels_path = &uri.relationships_uri().as_str()[1..];
                    zip.start_file(rels_path, options)?;
                    zip.write_all(rels.to_xml().as_bytes())?;
                }
            }
        }

        zip.finish()?;
        Ok(())
    }
}
```

### 5.2 增量写入优化

对于修改场景，只重写修改过的部件：

```rust
impl Package {
    pub fn save_incremental<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options = FileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        // 写入 [Content_Types].xml（总是重写）
        zip.start_file("[Content_Types].xml", options)?;
        zip.write_all(self.content_types.to_xml().as_bytes())?;

        // 写入包级关系
        if !self.relationships.is_empty() {
            zip.start_file("_rels/.rels", options)?;
            zip.write_all(self.relationships.to_xml().as_bytes())?;
        }

        // 写入部件
        for (uri, part) in &self.parts {
            let path = &uri.as_str()[1..];

            if part.is_modified() {
                // 部件已修改，写入新数据
                zip.start_file(path, options)?;
                zip.write_all(part.data())?;
            } else if let Some(source) = &self.source {
                // 部件未修改，直接从原始 ZIP 复制
                self.copy_from_source(&mut zip, source, path)?;
            }

            // 处理关系文件
            if let Some(rels) = part.relationships() {
                if !rels.is_empty() {
                    let rels_path = &uri.relationships_uri().as_str()[1..];
                    zip.start_file(rels_path, options)?;
                    zip.write_all(rels.to_xml().as_bytes())?;
                }
            }
        }

        zip.finish()?;
        Ok(())
    }
}
```

## 6. XML 序列化

### 6.1 ContentTypes 序列化

```rust
impl ContentTypes {
    pub fn to_xml(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);

        // 写入默认扩展名映射
        for (ext, ct) in &self.defaults {
            xml.push_str(&format!(
                r#"<Default Extension="{}" ContentType="{}"/>"#,
                ext, ct
            ));
        }

        // 写入覆盖
        for (uri, ct) in &self.overrides {
            xml.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                uri, ct
            ));
        }

        xml.push_str("</Types>");
        xml
    }

    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut ct = ContentTypes::new();
        ct.defaults.clear(); // 清除默认值，从 XML 读取

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) => {
                    match e.name().as_ref() {
                        b"Default" => {
                            let ext = get_attr(&e, "Extension")?;
                            let content_type = get_attr(&e, "ContentType")?;
                            ct.defaults.insert(ext, content_type);
                        }
                        b"Override" => {
                            let part_name = get_attr(&e, "PartName")?;
                            let content_type = get_attr(&e, "ContentType")?;
                            ct.overrides.insert(PartUri::new(&part_name)?, content_type);
                        }
                        _ => {}
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(ct)
    }
}
```

### 6.2 Relationships 序列化

```rust
impl Relationships {
    pub fn to_xml(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);

        for rel in self.items.values() {
            xml.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}""#,
                rel.id, rel.rel_type, rel.target
            ));

            if rel.target_mode == TargetMode::External {
                xml.push_str(r#" TargetMode="External""#);
            }

            xml.push_str("/>");
        }

        xml.push_str("</Relationships>");
        xml
    }

    pub fn from_xml(xml: &str) -> Result<Self> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut rels = Relationships::new();

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    let id = get_attr(&e, "Id")?;
                    let rel_type = get_attr(&e, "Type")?;
                    let target = get_attr(&e, "Target")?;
                    let target_mode = get_attr_opt(&e, "TargetMode")
                        .map(|s| if s == "External" { TargetMode::External } else { TargetMode::Internal })
                        .unwrap_or_default();

                    rels.items.insert(id.clone(), Relationship {
                        id,
                        rel_type,
                        target,
                        target_mode,
                    });
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // 更新 next_id
        rels.update_next_id();

        Ok(rels)
    }

    fn update_next_id(&mut self) {
        let max_id = self.items.keys()
            .filter_map(|id| {
                if id.starts_with("rId") {
                    id[3..].parse::<u32>().ok()
                } else {
                    None
                }
            })
            .max()
            .unwrap_or(0);

        self.next_id = max_id + 1;
    }
}
```

## 7. 错误处理

```rust
/// OPC 层错误
#[derive(Debug, thiserror::Error)]
pub enum OpcError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Missing required part: {0}")]
    MissingPart(String),

    #[error("Invalid part URI: {0}")]
    InvalidPartUri(String),

    #[error("Invalid content type: {0}")]
    InvalidContentType(String),

    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    #[error("Duplicate part: {0}")]
    DuplicatePart(String),

    #[error("Part not found: {0}")]
    PartNotFound(String),
}
```

## 8. 测试策略

### 8.1 单元测试

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_part_uri_parsing() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        assert_eq!(uri.file_name(), Some("document.xml"));
        assert_eq!(uri.extension(), Some("xml"));
        assert_eq!(uri.parent().unwrap().as_str(), "/word");
    }

    #[test]
    fn test_part_uri_relationships() {
        let uri = PartUri::new("/word/document.xml").unwrap();
        let rels_uri = uri.relationships_uri();
        assert_eq!(rels_uri.as_str(), "/word/_rels/document.xml.rels");
    }

    #[test]
    fn test_content_types_roundtrip() {
        let mut ct = ContentTypes::new();
        ct.add_override(&PartUri::new("/word/document.xml").unwrap(), content_types::MAIN_DOCUMENT);

        let xml = ct.to_xml();
        let ct2 = ContentTypes::from_xml(&xml).unwrap();

        assert_eq!(
            ct2.get_content_type(&PartUri::new("/word/document.xml").unwrap()),
            Some(content_types::MAIN_DOCUMENT)
        );
    }

    #[test]
    fn test_relationships_roundtrip() {
        let mut rels = Relationships::new();
        let id = rels.add(rel_types::STYLES, "styles.xml");

        let xml = rels.to_xml();
        let rels2 = Relationships::from_xml(&xml).unwrap();

        assert!(rels2.get(&id).is_some());
    }
}
```

### 8.2 集成测试

```rust
#[cfg(test)]
mod integration_tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_open_and_save_package() {
        // 创建测试包
        let mut pkg = Package::new();
        // ... 添加部件

        // 保存到内存
        let mut buf = Cursor::new(Vec::new());
        pkg.write_to(&mut buf).unwrap();

        // 重新打开
        buf.set_position(0);
        let pkg2 = Package::from_reader(buf).unwrap();

        // 验证内容一致
        // ...
    }

    #[test]
    fn test_read_real_docx() {
        let pkg = Package::open("tests/fixtures/simple.docx").unwrap();

        // 验证主文档存在
        assert!(pkg.main_document_part().is_some());

        // 验证关系正确
        let doc_rel = pkg.relationships().by_type(rel_types::OFFICE_DOCUMENT);
        assert!(doc_rel.is_some());
    }
}
```

## 9. 性能考量

1. **惰性加载**：大文件不一次性加载所有部件
2. **增量保存**：只重写修改过的部件
3. **内存映射**：对于超大文件可考虑 mmap
4. **并行解析**：多个独立部件可以并行解析
5. **缓冲 IO**：使用 BufReader/BufWriter

## 10. 下一步

1. 实现 OPC 层核心代码
2. 编写单元测试
3. 与 Elements 层集成
