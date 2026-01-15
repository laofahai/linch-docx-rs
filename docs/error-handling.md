# linch-docx-rs 错误处理设计

## 1. 设计原则

- **精确定位**：错误信息应该帮助用户快速定位问题
- **分层错误**：不同层有不同的错误类型，可以向上转换
- **可恢复性**：区分可恢复错误和不可恢复错误
- **友好提示**：错误信息对人类可读
- **上下文丰富**：包含文件路径、行号、元素位置等

## 2. 错误类型层次

```
Error (顶层错误)
├── IoError          # IO 错误
├── ZipError         # ZIP 错误
├── OpcError         # OPC 层错误
│   ├── MissingPart
│   ├── InvalidPartUri
│   ├── InvalidContentType
│   └── InvalidRelationship
├── XmlError         # XML 解析错误
│   ├── ParseError
│   ├── InvalidStructure
│   └── EncodingError
├── DocumentError    # 文档结构错误
│   ├── InvalidDocument
│   ├── MissingElement
│   └── InvalidReference
└── ValidationError  # 验证错误
    ├── InvalidValue
    ├── OutOfRange
    └── ConstraintViolation
```

## 3. 错误定义

### 3.1 顶层错误

```rust
use thiserror::Error;

/// linch-docx-rs 顶层错误类型
#[derive(Debug, Error)]
pub enum Error {
    /// IO 错误
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP 处理错误
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// OPC 包错误
    #[error(transparent)]
    Opc(#[from] OpcError),

    /// XML 解析错误
    #[error(transparent)]
    Xml(#[from] XmlError),

    /// 文档结构错误
    #[error(transparent)]
    Document(#[from] DocumentError),

    /// 验证错误
    #[error(transparent)]
    Validation(#[from] ValidationError),
}

/// Result 类型别名
pub type Result<T> = std::result::Result<T, Error>;
```

### 3.2 OPC 层错误

```rust
/// OPC 层错误
#[derive(Debug, Error)]
pub enum OpcError {
    /// 缺少必需的部件
    #[error("Missing required part: {part_uri}")]
    MissingPart {
        part_uri: String,
        #[source]
        source: Option<Box<dyn std::error::Error + Send + Sync>>,
    },

    /// 无效的部件 URI
    #[error("Invalid part URI '{uri}': {reason}")]
    InvalidPartUri {
        uri: String,
        reason: String,
    },

    /// 无效的内容类型
    #[error("Invalid content type for '{part_uri}': expected '{expected}', got '{actual}'")]
    InvalidContentType {
        part_uri: String,
        expected: String,
        actual: String,
    },

    /// 无效的关系
    #[error("Invalid relationship '{rel_id}' in '{source_part}': {reason}")]
    InvalidRelationship {
        rel_id: String,
        source_part: String,
        reason: String,
    },

    /// 重复的部件
    #[error("Duplicate part: {part_uri}")]
    DuplicatePart {
        part_uri: String,
    },

    /// 部件未找到
    #[error("Part not found: {part_uri}")]
    PartNotFound {
        part_uri: String,
    },

    /// 循环引用
    #[error("Circular reference detected: {path}")]
    CircularReference {
        path: String,
    },
}

impl OpcError {
    /// 创建缺少部件错误
    pub fn missing_part(uri: impl Into<String>) -> Self {
        Self::MissingPart {
            part_uri: uri.into(),
            source: None,
        }
    }

    /// 创建无效 URI 错误
    pub fn invalid_uri(uri: impl Into<String>, reason: impl Into<String>) -> Self {
        Self::InvalidPartUri {
            uri: uri.into(),
            reason: reason.into(),
        }
    }
}
```

### 3.3 XML 错误

```rust
/// XML 解析错误
#[derive(Debug, Error)]
pub enum XmlError {
    /// 解析错误
    #[error("XML parse error at line {line}, column {column}: {message}")]
    Parse {
        line: usize,
        column: usize,
        message: String,
        #[source]
        source: Option<quick_xml::Error>,
    },

    /// 无效的 XML 结构
    #[error("Invalid XML structure at '{path}': {message}")]
    InvalidStructure {
        path: String,
        message: String,
    },

    /// 编码错误
    #[error("Encoding error: {message}")]
    Encoding {
        message: String,
        #[source]
        source: Option<std::str::Utf8Error>,
    },

    /// 意外的 EOF
    #[error("Unexpected end of file while parsing '{context}'")]
    UnexpectedEof {
        context: String,
    },

    /// 缺少必需属性
    #[error("Missing required attribute '{attribute}' on element '{element}'")]
    MissingAttribute {
        element: String,
        attribute: String,
    },

    /// 无效的属性值
    #[error("Invalid value '{value}' for attribute '{attribute}' on element '{element}': {reason}")]
    InvalidAttributeValue {
        element: String,
        attribute: String,
        value: String,
        reason: String,
    },

    /// 意外的元素
    #[error("Unexpected element '{found}' at '{path}', expected one of: {expected:?}")]
    UnexpectedElement {
        path: String,
        found: String,
        expected: Vec<String>,
    },

    /// 缺少必需元素
    #[error("Missing required element '{element}' in '{parent}'")]
    MissingElement {
        parent: String,
        element: String,
    },
}

impl XmlError {
    /// 从 quick_xml 错误创建
    pub fn from_quick_xml(err: quick_xml::Error, context: &str) -> Self {
        Self::Parse {
            line: 0,  // quick_xml 不提供行号
            column: 0,
            message: format!("{} while parsing {}", err, context),
            source: Some(err),
        }
    }

    /// 创建缺少属性错误
    pub fn missing_attr(element: impl Into<String>, attribute: impl Into<String>) -> Self {
        Self::MissingAttribute {
            element: element.into(),
            attribute: attribute.into(),
        }
    }

    /// 创建无效属性值错误
    pub fn invalid_attr(
        element: impl Into<String>,
        attribute: impl Into<String>,
        value: impl Into<String>,
        reason: impl Into<String>,
    ) -> Self {
        Self::InvalidAttributeValue {
            element: element.into(),
            attribute: attribute.into(),
            value: value.into(),
            reason: reason.into(),
        }
    }
}
```

### 3.4 文档错误

```rust
/// 文档结构错误
#[derive(Debug, Error)]
pub enum DocumentError {
    /// 无效的文档
    #[error("Invalid DOCX document: {reason}")]
    InvalidDocument {
        reason: String,
        #[source]
        source: Option<Box<dyn std::error::Error + Send + Sync>>,
    },

    /// 缺少元素
    #[error("Missing element: {element} at {location}")]
    MissingElement {
        element: String,
        location: String,
    },

    /// 无效的引用
    #[error("Invalid reference '{ref_id}' of type '{ref_type}': {reason}")]
    InvalidReference {
        ref_id: String,
        ref_type: String,
        reason: String,
    },

    /// 索引越界
    #[error("Index {index} out of bounds (max: {max}) for {context}")]
    IndexOutOfBounds {
        index: usize,
        max: usize,
        context: String,
    },

    /// 样式未找到
    #[error("Style not found: '{style_id}'")]
    StyleNotFound {
        style_id: String,
    },

    /// 编号定义未找到
    #[error("Numbering definition not found: numId={num_id}")]
    NumberingNotFound {
        num_id: u32,
    },

    /// 图片未找到
    #[error("Image not found: relationship '{rel_id}'")]
    ImageNotFound {
        rel_id: String,
    },

    /// 不支持的功能
    #[error("Unsupported feature: {feature}")]
    Unsupported {
        feature: String,
    },
}

impl DocumentError {
    pub fn invalid(reason: impl Into<String>) -> Self {
        Self::InvalidDocument {
            reason: reason.into(),
            source: None,
        }
    }

    pub fn index_out_of_bounds(index: usize, max: usize, context: impl Into<String>) -> Self {
        Self::IndexOutOfBounds {
            index,
            max,
            context: context.into(),
        }
    }

    pub fn unsupported(feature: impl Into<String>) -> Self {
        Self::Unsupported {
            feature: feature.into(),
        }
    }
}
```

### 3.5 验证错误

```rust
/// 验证错误
#[derive(Debug, Error)]
pub enum ValidationError {
    /// 无效的值
    #[error("Invalid value '{value}' for '{field}': {reason}")]
    InvalidValue {
        field: String,
        value: String,
        reason: String,
    },

    /// 超出范围
    #[error("Value {value} out of range [{min}, {max}] for '{field}'")]
    OutOfRange {
        field: String,
        value: String,
        min: String,
        max: String,
    },

    /// 约束违反
    #[error("Constraint violation for '{field}': {constraint}")]
    ConstraintViolation {
        field: String,
        constraint: String,
    },

    /// 格式错误
    #[error("Invalid format for '{field}': expected {expected}, got '{actual}'")]
    InvalidFormat {
        field: String,
        expected: String,
        actual: String,
    },

    /// 必需字段缺失
    #[error("Required field '{field}' is missing")]
    RequiredFieldMissing {
        field: String,
    },
}

impl ValidationError {
    pub fn invalid_value(
        field: impl Into<String>,
        value: impl Into<String>,
        reason: impl Into<String>,
    ) -> Self {
        Self::InvalidValue {
            field: field.into(),
            value: value.into(),
            reason: reason.into(),
        }
    }

    pub fn out_of_range<T: std::fmt::Display>(
        field: impl Into<String>,
        value: T,
        min: T,
        max: T,
    ) -> Self {
        Self::OutOfRange {
            field: field.into(),
            value: value.to_string(),
            min: min.to_string(),
            max: max.to_string(),
        }
    }
}
```

## 4. 错误上下文

### 4.1 位置信息

```rust
/// 文档位置
#[derive(Clone, Debug, Default)]
pub struct DocumentLocation {
    /// 部件 URI
    pub part: Option<String>,
    /// XPath 路径
    pub path: Option<String>,
    /// 段落索引
    pub paragraph_index: Option<usize>,
    /// Run 索引
    pub run_index: Option<usize>,
    /// 表格索引
    pub table_index: Option<usize>,
    /// 行索引
    pub row_index: Option<usize>,
    /// 列索引
    pub column_index: Option<usize>,
}

impl std::fmt::Display for DocumentLocation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut parts = Vec::new();

        if let Some(ref part) = self.part {
            parts.push(format!("part={}", part));
        }
        if let Some(ref path) = self.path {
            parts.push(format!("path={}", path));
        }
        if let Some(idx) = self.paragraph_index {
            parts.push(format!("paragraph[{}]", idx));
        }
        if let Some(idx) = self.run_index {
            parts.push(format!("run[{}]", idx));
        }
        if let Some(idx) = self.table_index {
            parts.push(format!("table[{}]", idx));
        }

        write!(f, "{}", parts.join(", "))
    }
}
```

### 4.2 带上下文的错误

```rust
/// 带上下文的错误包装
#[derive(Debug, Error)]
#[error("{message}")]
pub struct ContextualError {
    message: String,
    location: DocumentLocation,
    #[source]
    source: Box<dyn std::error::Error + Send + Sync>,
}

impl ContextualError {
    pub fn new(
        error: impl std::error::Error + Send + Sync + 'static,
        message: impl Into<String>,
        location: DocumentLocation,
    ) -> Self {
        Self {
            message: message.into(),
            location,
            source: Box::new(error),
        }
    }

    pub fn location(&self) -> &DocumentLocation {
        &self.location
    }
}

/// 错误上下文扩展 trait
pub trait ErrorContext<T> {
    /// 添加错误上下文
    fn with_context(self, message: impl Into<String>) -> Result<T>;

    /// 添加位置上下文
    fn with_location(self, location: DocumentLocation) -> Result<T>;

    /// 添加部件上下文
    fn in_part(self, part: impl Into<String>) -> Result<T>;

    /// 添加段落上下文
    fn in_paragraph(self, index: usize) -> Result<T>;
}

impl<T, E: std::error::Error + Send + Sync + 'static> ErrorContext<T> for std::result::Result<T, E> {
    fn with_context(self, message: impl Into<String>) -> Result<T> {
        self.map_err(|e| {
            Error::Document(DocumentError::InvalidDocument {
                reason: format!("{}: {}", message.into(), e),
                source: Some(Box::new(e)),
            })
        })
    }

    fn with_location(self, location: DocumentLocation) -> Result<T> {
        self.map_err(|e| {
            Error::Document(DocumentError::InvalidDocument {
                reason: format!("at {}: {}", location, e),
                source: Some(Box::new(e)),
            })
        })
    }

    fn in_part(self, part: impl Into<String>) -> Result<T> {
        let part = part.into();
        self.map_err(|e| {
            Error::Document(DocumentError::InvalidDocument {
                reason: format!("in part '{}': {}", part, e),
                source: Some(Box::new(e)),
            })
        })
    }

    fn in_paragraph(self, index: usize) -> Result<T> {
        self.map_err(|e| {
            Error::Document(DocumentError::InvalidDocument {
                reason: format!("in paragraph {}: {}", index, e),
                source: Some(Box::new(e)),
            })
        })
    }
}
```

## 5. 错误恢复

### 5.1 宽松解析模式

```rust
/// 解析配置
#[derive(Clone, Debug, Default)]
pub struct ParseOptions {
    /// 是否使用宽松模式
    pub lenient: bool,
    /// 是否忽略未知元素
    pub ignore_unknown: bool,
    /// 是否在错误时继续
    pub continue_on_error: bool,
    /// 警告回调
    pub on_warning: Option<fn(Warning)>,
}

/// 警告信息
#[derive(Clone, Debug)]
pub struct Warning {
    pub message: String,
    pub location: DocumentLocation,
    pub severity: WarningSeverity,
}

#[derive(Clone, Copy, Debug)]
pub enum WarningSeverity {
    Info,
    Warning,
    Error,  // 被降级为警告的错误
}

impl ParseOptions {
    /// 严格模式
    pub fn strict() -> Self {
        Self::default()
    }

    /// 宽松模式
    pub fn lenient() -> Self {
        Self {
            lenient: true,
            ignore_unknown: true,
            continue_on_error: true,
            ..Default::default()
        }
    }
}
```

### 5.2 错误收集

```rust
/// 解析结果（带警告）
pub struct ParseResult<T> {
    pub value: T,
    pub warnings: Vec<Warning>,
}

impl<T> ParseResult<T> {
    pub fn new(value: T) -> Self {
        Self { value, warnings: Vec::new() }
    }

    pub fn with_warning(mut self, warning: Warning) -> Self {
        self.warnings.push(warning);
        self
    }

    pub fn has_warnings(&self) -> bool {
        !self.warnings.is_empty()
    }

    pub fn map<U>(self, f: impl FnOnce(T) -> U) -> ParseResult<U> {
        ParseResult {
            value: f(self.value),
            warnings: self.warnings,
        }
    }
}

/// 错误收集器
pub struct ErrorCollector {
    warnings: Vec<Warning>,
    errors: Vec<Error>,
    options: ParseOptions,
}

impl ErrorCollector {
    pub fn new(options: ParseOptions) -> Self {
        Self {
            warnings: Vec::new(),
            errors: Vec::new(),
            options,
        }
    }

    /// 处理错误（根据配置决定是报错还是记录为警告）
    pub fn handle_error(&mut self, error: Error, location: DocumentLocation) -> Result<()> {
        if self.options.continue_on_error {
            self.warnings.push(Warning {
                message: error.to_string(),
                location,
                severity: WarningSeverity::Error,
            });
            Ok(())
        } else {
            Err(error)
        }
    }

    /// 添加警告
    pub fn warn(&mut self, message: impl Into<String>, location: DocumentLocation) {
        self.warnings.push(Warning {
            message: message.into(),
            location,
            severity: WarningSeverity::Warning,
        });

        if let Some(callback) = self.options.on_warning {
            callback(self.warnings.last().unwrap().clone());
        }
    }

    /// 获取所有警告
    pub fn warnings(&self) -> &[Warning] {
        &self.warnings
    }

    /// 是否有错误
    pub fn has_errors(&self) -> bool {
        !self.errors.is_empty()
    }

    /// 完成并返回结果
    pub fn finish<T>(self, value: T) -> ParseResult<T> {
        ParseResult {
            value,
            warnings: self.warnings,
        }
    }
}
```

## 6. 用户友好的错误消息

### 6.1 错误格式化

```rust
impl Error {
    /// 获取用户友好的错误消息
    pub fn user_message(&self) -> String {
        match self {
            Error::Io(e) => format!("无法读取或写入文件: {}", e),
            Error::Zip(e) => format!("无效的 DOCX 文件（ZIP 格式错误）: {}", e),
            Error::Opc(e) => e.user_message(),
            Error::Xml(e) => e.user_message(),
            Error::Document(e) => e.user_message(),
            Error::Validation(e) => e.user_message(),
        }
    }

    /// 获取建议的解决方案
    pub fn suggestion(&self) -> Option<String> {
        match self {
            Error::Zip(_) => Some(
                "请确保文件是有效的 .docx 文件，而不是 .doc（旧版 Word 格式）".into()
            ),
            Error::Opc(OpcError::MissingPart { part_uri, .. }) => Some(
                format!("文件可能已损坏，缺少必需的组件 '{}'", part_uri)
            ),
            Error::Document(DocumentError::StyleNotFound { style_id }) => Some(
                format!("样式 '{}' 不存在，请检查样式名称是否正确", style_id)
            ),
            _ => None,
        }
    }
}

impl OpcError {
    pub fn user_message(&self) -> String {
        match self {
            Self::MissingPart { part_uri, .. } => {
                format!("文档缺少必需的组件: {}", part_uri)
            }
            Self::InvalidPartUri { uri, reason } => {
                format!("无效的文档路径 '{}': {}", uri, reason)
            }
            Self::InvalidContentType { part_uri, expected, actual } => {
                format!(
                    "组件 '{}' 的类型不正确（预期 {}，实际 {}）",
                    part_uri, expected, actual
                )
            }
            Self::InvalidRelationship { rel_id, source_part, reason } => {
                format!(
                    "组件 '{}' 中的引用 '{}' 无效: {}",
                    source_part, rel_id, reason
                )
            }
            Self::DuplicatePart { part_uri } => {
                format!("文档中存在重复的组件: {}", part_uri)
            }
            Self::PartNotFound { part_uri } => {
                format!("找不到文档组件: {}", part_uri)
            }
            Self::CircularReference { path } => {
                format!("检测到循环引用: {}", path)
            }
        }
    }
}

impl XmlError {
    pub fn user_message(&self) -> String {
        match self {
            Self::Parse { line, column, message, .. } => {
                if *line > 0 {
                    format!("XML 解析错误（第 {} 行，第 {} 列）: {}", line, column, message)
                } else {
                    format!("XML 解析错误: {}", message)
                }
            }
            Self::InvalidStructure { path, message } => {
                format!("文档结构错误（位置 {}）: {}", path, message)
            }
            Self::Encoding { message, .. } => {
                format!("编码错误: {}", message)
            }
            Self::UnexpectedEof { context } => {
                format!("文件意外结束，正在解析: {}", context)
            }
            Self::MissingAttribute { element, attribute } => {
                format!("元素 '{}' 缺少必需的属性 '{}'", element, attribute)
            }
            Self::InvalidAttributeValue { element, attribute, value, reason } => {
                format!(
                    "元素 '{}' 的属性 '{}' 值 '{}' 无效: {}",
                    element, attribute, value, reason
                )
            }
            Self::UnexpectedElement { path, found, expected } => {
                format!(
                    "在 '{}' 位置发现意外元素 '{}'，预期: {:?}",
                    path, found, expected
                )
            }
            Self::MissingElement { parent, element } => {
                format!("'{}' 中缺少必需的元素 '{}'", parent, element)
            }
        }
    }
}

impl DocumentError {
    pub fn user_message(&self) -> String {
        match self {
            Self::InvalidDocument { reason, .. } => {
                format!("无效的文档: {}", reason)
            }
            Self::MissingElement { element, location } => {
                format!("缺少元素 '{}' 于 {}", element, location)
            }
            Self::InvalidReference { ref_id, ref_type, reason } => {
                format!("无效的{}引用 '{}': {}", ref_type, ref_id, reason)
            }
            Self::IndexOutOfBounds { index, max, context } => {
                format!("{} 的索引 {} 超出范围（最大: {}）", context, index, max)
            }
            Self::StyleNotFound { style_id } => {
                format!("找不到样式: '{}'", style_id)
            }
            Self::NumberingNotFound { num_id } => {
                format!("找不到编号定义: numId={}", num_id)
            }
            Self::ImageNotFound { rel_id } => {
                format!("找不到图片: 关系 '{}'", rel_id)
            }
            Self::Unsupported { feature } => {
                format!("不支持的功能: {}", feature)
            }
        }
    }
}

impl ValidationError {
    pub fn user_message(&self) -> String {
        match self {
            Self::InvalidValue { field, value, reason } => {
                format!("字段 '{}' 的值 '{}' 无效: {}", field, value, reason)
            }
            Self::OutOfRange { field, value, min, max } => {
                format!(
                    "字段 '{}' 的值 {} 超出允许范围 [{}, {}]",
                    field, value, min, max
                )
            }
            Self::ConstraintViolation { field, constraint } => {
                format!("字段 '{}' 违反约束: {}", field, constraint)
            }
            Self::InvalidFormat { field, expected, actual } => {
                format!(
                    "字段 '{}' 格式错误: 预期 {}，实际 '{}'",
                    field, expected, actual
                )
            }
            Self::RequiredFieldMissing { field } => {
                format!("必需字段 '{}' 缺失", field)
            }
        }
    }
}
```

### 6.2 错误报告

```rust
/// 错误报告生成器
pub struct ErrorReport {
    error: Error,
    location: Option<DocumentLocation>,
    context: Vec<String>,
}

impl ErrorReport {
    pub fn new(error: Error) -> Self {
        Self {
            error,
            location: None,
            context: Vec::new(),
        }
    }

    pub fn with_location(mut self, location: DocumentLocation) -> Self {
        self.location = Some(location);
        self
    }

    pub fn with_context(mut self, ctx: impl Into<String>) -> Self {
        self.context.push(ctx.into());
        self
    }

    /// 生成详细报告
    pub fn detailed_report(&self) -> String {
        let mut report = String::new();

        // 错误消息
        report.push_str(&format!("错误: {}\n", self.error.user_message()));

        // 位置信息
        if let Some(ref loc) = self.location {
            report.push_str(&format!("位置: {}\n", loc));
        }

        // 上下文
        if !self.context.is_empty() {
            report.push_str("上下文:\n");
            for (i, ctx) in self.context.iter().enumerate() {
                report.push_str(&format!("  {}. {}\n", i + 1, ctx));
            }
        }

        // 建议
        if let Some(suggestion) = self.error.suggestion() {
            report.push_str(&format!("建议: {}\n", suggestion));
        }

        // 技术细节
        report.push_str(&format!("\n技术细节: {:?}\n", self.error));

        report
    }

    /// 生成简短报告
    pub fn brief_report(&self) -> String {
        self.error.user_message()
    }
}
```

## 7. 测试

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_display() {
        let err = Error::Opc(OpcError::missing_part("/word/document.xml"));
        assert!(err.to_string().contains("document.xml"));
    }

    #[test]
    fn test_error_user_message() {
        let err = Error::Zip(zip::result::ZipError::InvalidArchive("test"));
        let msg = err.user_message();
        assert!(msg.contains("DOCX"));
    }

    #[test]
    fn test_error_context() {
        let result: std::result::Result<(), std::io::Error> =
            Err(std::io::Error::new(std::io::ErrorKind::NotFound, "test"));

        let err = result
            .with_context("while reading document")
            .unwrap_err();

        assert!(err.to_string().contains("reading document"));
    }

    #[test]
    fn test_error_collector() {
        let options = ParseOptions::lenient();
        let mut collector = ErrorCollector::new(options);

        collector.warn("test warning", DocumentLocation::default());

        assert_eq!(collector.warnings().len(), 1);
    }

    #[test]
    fn test_parse_result() {
        let result = ParseResult::new(42)
            .with_warning(Warning {
                message: "test".into(),
                location: DocumentLocation::default(),
                severity: WarningSeverity::Warning,
            });

        assert_eq!(result.value, 42);
        assert!(result.has_warnings());
    }
}
```

## 8. 总结

linch-docx-rs 的错误处理设计特点：

1. **分层错误**：每层有专门的错误类型，清晰的错误来源
2. **丰富上下文**：包含位置、路径、索引等信息
3. **用户友好**：提供中文错误消息和修复建议
4. **可恢复性**：支持宽松模式和错误收集
5. **符合 Rust 惯例**：使用 `thiserror`，实现标准 trait
