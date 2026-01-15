# linch-docx-rs API 设计文档

## 1. 设计目标

- **简洁直观**：常用操作一行代码完成
- **类型安全**：编译期捕获错误
- **符合 Rust 惯例**：所有权清晰，错误处理规范
- **渐进式复杂度**：简单用法简单，复杂用法也不繁琐

## 2. 核心类型概览

```rust
// 主要类型
pub struct Document { ... }         // 文档主结构
pub struct Paragraph<'a> { ... }    // 段落（不可变）
pub struct ParagraphMut<'a> { ... } // 段落（可变）
pub struct Run<'a> { ... }          // 文本片段
pub struct RunMut<'a> { ... }       // 文本片段（可变）
pub struct Table<'a> { ... }        // 表格
pub struct TableMut<'a> { ... }     // 表格（可变）
pub struct Row<'a> { ... }          // 表格行
pub struct Cell<'a> { ... }         // 表格单元格
pub struct Style<'a> { ... }        // 样式

// 属性类型
pub struct ParagraphProperties { ... }
pub struct RunProperties { ... }
pub struct TableProperties { ... }

// 单位类型
pub struct Pt(pub f64);             // 磅
pub struct Emu(pub i64);            // EMU (English Metric Units)
pub struct Twip(pub i32);           // Twip (1/20 点)
pub struct Cm(pub f64);             // 厘米
pub struct Inch(pub f64);           // 英寸

// 错误类型
pub enum Error { ... }
pub type Result<T> = std::result::Result<T, Error>;
```

## 3. Document API

### 3.1 打开与创建

```rust
impl Document {
    /// 从文件路径打开文档
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self>;

    /// 从字节数组打开文档
    pub fn from_bytes(bytes: &[u8]) -> Result<Self>;

    /// 从 Read 实现打开文档
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self>;

    /// 创建空白文档
    pub fn new() -> Self;

    /// 从模板创建文档
    pub fn from_template<P: AsRef<Path>>(path: P) -> Result<Self>;
}
```

### 3.2 保存

```rust
impl Document {
    /// 保存到文件
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()>;

    /// 保存到字节数组
    pub fn to_bytes(&self) -> Result<Vec<u8>>;

    /// 保存到 Write 实现
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()>;
}
```

### 3.3 段落访问

```rust
impl Document {
    /// 获取所有段落的迭代器
    pub fn paragraphs(&self) -> Paragraphs<'_>;

    /// 获取可变段落迭代器
    pub fn paragraphs_mut(&mut self) -> ParagraphsMut<'_>;

    /// 按索引获取段落
    pub fn paragraph(&self, index: usize) -> Option<Paragraph<'_>>;

    /// 按索引获取可变段落
    pub fn paragraph_mut(&mut self, index: usize) -> Option<ParagraphMut<'_>>;

    /// 段落数量
    pub fn paragraph_count(&self) -> usize;

    /// 添加段落（返回 Builder）
    pub fn add_paragraph(&mut self, text: &str) -> ParagraphBuilder<'_>;

    /// 在指定位置插入段落
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> ParagraphBuilder<'_>;

    /// 删除段落
    pub fn remove_paragraph(&mut self, index: usize) -> bool;
}
```

### 3.4 表格访问

```rust
impl Document {
    /// 获取所有表格的迭代器
    pub fn tables(&self) -> Tables<'_>;

    /// 获取可变表格迭代器
    pub fn tables_mut(&mut self) -> TablesMut<'_>;

    /// 按索引获取表格
    pub fn table(&self, index: usize) -> Option<Table<'_>>;

    /// 表格数量
    pub fn table_count(&self) -> usize;

    /// 添加表格
    pub fn add_table(&mut self, rows: usize, cols: usize) -> TableBuilder<'_>;
}
```

### 3.5 样式访问

```rust
impl Document {
    /// 获取所有样式
    pub fn styles(&self) -> Styles<'_>;

    /// 按 ID 获取样式
    pub fn style(&self, style_id: &str) -> Option<Style<'_>>;

    /// 按名称获取样式
    pub fn style_by_name(&self, name: &str) -> Option<Style<'_>>;
}
```

### 3.6 文本操作

```rust
impl Document {
    /// 获取文档全部文本（纯文本）
    pub fn text(&self) -> String;

    /// 查找替换文本
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize;

    /// 使用正则表达式替换
    pub fn replace_regex(&mut self, pattern: &str, replace: &str) -> Result<usize>;

    /// 查找文本位置
    pub fn find_text(&self, text: &str) -> Vec<TextLocation>;
}

/// 文本位置
pub struct TextLocation {
    pub paragraph_index: usize,
    pub run_index: usize,
    pub char_offset: usize,
}
```

### 3.7 文档属性

```rust
impl Document {
    /// 获取核心属性（标题、作者等）
    pub fn core_properties(&self) -> &CoreProperties;

    /// 获取可变核心属性
    pub fn core_properties_mut(&mut self) -> &mut CoreProperties;
}

pub struct CoreProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub revision: Option<u32>,
    pub created: Option<DateTime>,
    pub modified: Option<DateTime>,
}
```

## 4. Paragraph API

### 4.1 不可变访问

```rust
impl<'a> Paragraph<'a> {
    /// 获取段落全部文本
    pub fn text(&self) -> String;

    /// 获取样式 ID
    pub fn style(&self) -> Option<&str>;

    /// 获取样式名称
    pub fn style_name(&self) -> Option<&str>;

    /// 获取对齐方式
    pub fn alignment(&self) -> Option<Alignment>;

    /// 获取缩进
    pub fn indentation(&self) -> Option<&Indentation>;

    /// 获取行距
    pub fn line_spacing(&self) -> Option<&LineSpacing>;

    /// 获取所有 Run
    pub fn runs(&self) -> Runs<'a>;

    /// Run 数量
    pub fn run_count(&self) -> usize;

    /// 是否为空段落
    pub fn is_empty(&self) -> bool;

    /// 获取编号信息（如果是列表项）
    pub fn numbering(&self) -> Option<&NumberingInfo>;

    /// 是否是标题
    pub fn is_heading(&self) -> bool;

    /// 获取标题级别（1-9，非标题返回 None）
    pub fn heading_level(&self) -> Option<u8>;
}

/// 对齐方式
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

/// 缩进
pub struct Indentation {
    pub left: Option<Twip>,
    pub right: Option<Twip>,
    pub first_line: Option<Twip>,
    pub hanging: Option<Twip>,
}

/// 行距
pub struct LineSpacing {
    pub before: Option<Twip>,
    pub after: Option<Twip>,
    pub line: Option<LineSpacingValue>,
}

pub enum LineSpacingValue {
    Auto,                    // 自动
    Exact(Twip),            // 固定值
    AtLeast(Twip),          // 最小值
    Multiple(f64),          // 倍数（1.0 = 单倍行距）
}
```

### 4.2 可变访问

```rust
impl<'a> ParagraphMut<'a> {
    /// 设置样式
    pub fn set_style(&mut self, style_id: &str) -> &mut Self;

    /// 设置对齐方式
    pub fn set_alignment(&mut self, alignment: Alignment) -> &mut Self;

    /// 设置缩进
    pub fn set_indentation(&mut self, indentation: Indentation) -> &mut Self;

    /// 设置行距
    pub fn set_line_spacing(&mut self, spacing: LineSpacing) -> &mut Self;

    /// 清除段落内容
    pub fn clear(&mut self) -> &mut Self;

    /// 设置文本（清除现有内容）
    pub fn set_text(&mut self, text: &str) -> &mut Self;

    /// 追加文本
    pub fn append_text(&mut self, text: &str) -> RunBuilder<'_>;

    /// 添加 Run
    pub fn add_run(&mut self, text: &str) -> RunBuilder<'_>;

    /// 获取可变 Run
    pub fn run_mut(&mut self, index: usize) -> Option<RunMut<'_>>;

    /// 删除 Run
    pub fn remove_run(&mut self, index: usize) -> bool;

    // 便捷方法（应用到所有 Run）
    /// 设置粗体
    pub fn set_bold(&mut self, bold: bool) -> &mut Self;

    /// 设置斜体
    pub fn set_italic(&mut self, italic: bool) -> &mut Self;

    /// 设置字号
    pub fn set_font_size(&mut self, size: Pt) -> &mut Self;

    /// 设置字体
    pub fn set_font(&mut self, font: &str) -> &mut Self;

    /// 设置颜色
    pub fn set_color(&mut self, color: Color) -> &mut Self;
}
```

### 4.3 ParagraphBuilder（链式构建）

```rust
pub struct ParagraphBuilder<'a> { ... }

impl<'a> ParagraphBuilder<'a> {
    pub fn style(self, style_id: &str) -> Self;
    pub fn alignment(self, alignment: Alignment) -> Self;
    pub fn indentation(self, indentation: Indentation) -> Self;
    pub fn line_spacing(self, spacing: LineSpacing) -> Self;
    pub fn bold(self, bold: bool) -> Self;
    pub fn italic(self, italic: bool) -> Self;
    pub fn font_size(self, size: Pt) -> Self;
    pub fn font(self, font: &str) -> Self;
    pub fn color(self, color: Color) -> Self;

    /// 完成构建，返回段落索引
    pub fn build(self) -> usize;
}
```

## 5. Run API

### 5.1 不可变访问

```rust
impl<'a> Run<'a> {
    /// 获取文本
    pub fn text(&self) -> &str;

    /// 是否粗体
    pub fn bold(&self) -> bool;

    /// 是否斜体
    pub fn italic(&self) -> bool;

    /// 是否有下划线
    pub fn underline(&self) -> Option<UnderlineType>;

    /// 是否有删除线
    pub fn strike(&self) -> bool;

    /// 字号
    pub fn font_size(&self) -> Option<Pt>;

    /// 字体
    pub fn font(&self) -> Option<&str>;

    /// 字体（东亚字体）
    pub fn font_east_asia(&self) -> Option<&str>;

    /// 颜色
    pub fn color(&self) -> Option<Color>;

    /// 高亮颜色
    pub fn highlight(&self) -> Option<HighlightColor>;

    /// 是否上标
    pub fn superscript(&self) -> bool;

    /// 是否下标
    pub fn subscript(&self) -> bool;

    /// 是否隐藏
    pub fn hidden(&self) -> bool;
}

/// 下划线类型
pub enum UnderlineType {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    Wave,
    // ... 更多类型
}

/// 颜色
#[derive(Clone, Debug, PartialEq)]
pub enum Color {
    /// RGB 颜色 (如 "FF0000" 表示红色)
    Rgb(String),
    /// 主题颜色
    Theme(ThemeColor),
    /// 自动（通常为黑色）
    Auto,
}

/// 主题颜色
pub enum ThemeColor {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

/// 高亮颜色
pub enum HighlightColor {
    Yellow,
    Green,
    Cyan,
    Magenta,
    Blue,
    Red,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
    Black,
    White,
}
```

### 5.2 可变访问

```rust
impl<'a> RunMut<'a> {
    pub fn set_text(&mut self, text: &str) -> &mut Self;
    pub fn set_bold(&mut self, bold: bool) -> &mut Self;
    pub fn set_italic(&mut self, italic: bool) -> &mut Self;
    pub fn set_underline(&mut self, underline: Option<UnderlineType>) -> &mut Self;
    pub fn set_strike(&mut self, strike: bool) -> &mut Self;
    pub fn set_font_size(&mut self, size: Pt) -> &mut Self;
    pub fn set_font(&mut self, font: &str) -> &mut Self;
    pub fn set_font_east_asia(&mut self, font: &str) -> &mut Self;
    pub fn set_color(&mut self, color: Color) -> &mut Self;
    pub fn set_highlight(&mut self, color: Option<HighlightColor>) -> &mut Self;
    pub fn set_superscript(&mut self, superscript: bool) -> &mut Self;
    pub fn set_subscript(&mut self, subscript: bool) -> &mut Self;

    /// 清除所有格式
    pub fn clear_formatting(&mut self) -> &mut Self;
}
```

### 5.3 RunBuilder

```rust
pub struct RunBuilder<'a> { ... }

impl<'a> RunBuilder<'a> {
    pub fn bold(self, bold: bool) -> Self;
    pub fn italic(self, italic: bool) -> Self;
    pub fn underline(self, underline: UnderlineType) -> Self;
    pub fn font_size(self, size: Pt) -> Self;
    pub fn font(self, font: &str) -> Self;
    pub fn color(self, color: Color) -> Self;
    pub fn highlight(self, color: HighlightColor) -> Self;
    pub fn build(self) -> usize;
}
```

## 6. Table API

### 6.1 不可变访问

```rust
impl<'a> Table<'a> {
    /// 行数
    pub fn row_count(&self) -> usize;

    /// 列数（基于第一行）
    pub fn column_count(&self) -> usize;

    /// 获取所有行
    pub fn rows(&self) -> Rows<'a>;

    /// 获取指定行
    pub fn row(&self, index: usize) -> Option<Row<'a>>;

    /// 获取指定单元格
    pub fn cell(&self, row: usize, col: usize) -> Option<Cell<'a>>;

    /// 获取表格宽度
    pub fn width(&self) -> Option<TableWidth>;

    /// 获取表格对齐方式
    pub fn alignment(&self) -> Option<TableAlignment>;
}

impl<'a> Row<'a> {
    /// 单元格数量
    pub fn cell_count(&self) -> usize;

    /// 获取所有单元格
    pub fn cells(&self) -> Cells<'a>;

    /// 获取指定单元格
    pub fn cell(&self, index: usize) -> Option<Cell<'a>>;

    /// 行高
    pub fn height(&self) -> Option<Twip>;
}

impl<'a> Cell<'a> {
    /// 获取单元格文本
    pub fn text(&self) -> String;

    /// 获取单元格内的段落
    pub fn paragraphs(&self) -> Paragraphs<'a>;

    /// 单元格宽度
    pub fn width(&self) -> Option<Twip>;

    /// 垂直对齐
    pub fn vertical_alignment(&self) -> Option<VerticalAlignment>;

    /// 合并信息
    pub fn merge_info(&self) -> CellMergeInfo;
}

/// 表格宽度
pub enum TableWidth {
    Auto,
    Percent(f64),
    Twips(Twip),
}

/// 表格对齐
pub enum TableAlignment {
    Left,
    Center,
    Right,
}

/// 垂直对齐
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

/// 单元格合并信息
pub struct CellMergeInfo {
    pub horizontal: Option<HorizontalMerge>,
    pub vertical: Option<VerticalMerge>,
}

pub enum HorizontalMerge {
    Restart,    // 合并起始
    Continue,   // 合并延续
}

pub enum VerticalMerge {
    Restart,
    Continue,
}
```

### 6.2 可变访问

```rust
impl<'a> TableMut<'a> {
    pub fn cell_mut(&mut self, row: usize, col: usize) -> Option<CellMut<'_>>;
    pub fn row_mut(&mut self, index: usize) -> Option<RowMut<'_>>;
    pub fn add_row(&mut self) -> RowBuilder<'_>;
    pub fn insert_row(&mut self, index: usize) -> RowBuilder<'_>;
    pub fn remove_row(&mut self, index: usize) -> bool;
    pub fn add_column(&mut self);
    pub fn remove_column(&mut self, index: usize) -> bool;
    pub fn set_width(&mut self, width: TableWidth) -> &mut Self;
    pub fn set_alignment(&mut self, alignment: TableAlignment) -> &mut Self;
}

impl<'a> CellMut<'a> {
    pub fn set_text(&mut self, text: &str) -> &mut Self;
    pub fn add_paragraph(&mut self, text: &str) -> ParagraphBuilder<'_>;
    pub fn clear(&mut self) -> &mut Self;
    pub fn set_width(&mut self, width: Twip) -> &mut Self;
    pub fn set_vertical_alignment(&mut self, alignment: VerticalAlignment) -> &mut Self;
    pub fn set_background(&mut self, color: Color) -> &mut Self;
}
```

### 6.3 TableBuilder

```rust
pub struct TableBuilder<'a> { ... }

impl<'a> TableBuilder<'a> {
    pub fn width(self, width: TableWidth) -> Self;
    pub fn alignment(self, alignment: TableAlignment) -> Self;
    pub fn style(self, style_id: &str) -> Self;
    pub fn borders(self, borders: TableBorders) -> Self;

    /// 设置表头行数
    pub fn header_rows(self, count: usize) -> Self;

    /// 填充数据
    pub fn data(self, data: &[&[&str]]) -> Self;

    pub fn build(self) -> usize;
}
```

## 7. Style API

```rust
impl<'a> Style<'a> {
    /// 样式 ID
    pub fn id(&self) -> &str;

    /// 样式名称
    pub fn name(&self) -> Option<&str>;

    /// 样式类型
    pub fn style_type(&self) -> StyleType;

    /// 基于的样式
    pub fn based_on(&self) -> Option<&str>;

    /// 下一个段落的样式
    pub fn next_style(&self) -> Option<&str>;

    /// 是否为默认样式
    pub fn is_default(&self) -> bool;

    /// 获取段落属性
    pub fn paragraph_properties(&self) -> Option<&ParagraphProperties>;

    /// 获取 Run 属性
    pub fn run_properties(&self) -> Option<&RunProperties>;
}

pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}
```

## 8. 单位转换

```rust
// 单位类型定义
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Pt(pub f64);      // 磅 (1/72 inch)

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Emu(pub i64);     // EMU (914400 per inch)

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Twip(pub i32);    // Twip (1/20 point, 1440 per inch)

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Cm(pub f64);      // 厘米

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Mm(pub f64);      // 毫米

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Inch(pub f64);    // 英寸

// 单位转换
impl Pt {
    pub fn to_twip(self) -> Twip;
    pub fn to_emu(self) -> Emu;
    pub fn to_cm(self) -> Cm;
    pub fn to_inch(self) -> Inch;
}

impl From<Twip> for Pt { ... }
impl From<Emu> for Pt { ... }
impl From<Cm> for Pt { ... }
impl From<Inch> for Pt { ... }

// 便捷构造
impl Pt {
    pub const fn new(value: f64) -> Self { Self(value) }
}

impl Cm {
    pub const fn new(value: f64) -> Self { Self(value) }
}

// 使用示例
let size = Pt(12.0);
let width = Cm(2.5).to_twip();
```

## 9. 错误类型

```rust
#[derive(Debug, thiserror::Error)]
pub enum Error {
    /// IO 错误
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP 错误
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML 解析错误
    #[error("XML parse error at line {line}, column {column}: {message}")]
    XmlParse {
        line: usize,
        column: usize,
        message: String,
    },

    /// 无效的 DOCX 文件
    #[error("Invalid DOCX file: {0}")]
    InvalidDocx(String),

    /// 缺少必需的部件
    #[error("Missing required part: {0}")]
    MissingPart(String),

    /// 无效的关系
    #[error("Invalid relationship: {0}")]
    InvalidRelationship(String),

    /// 无效的内容类型
    #[error("Invalid content type: {0}")]
    InvalidContentType(String),

    /// 元素未找到
    #[error("Element not found: {0}")]
    NotFound(String),

    /// 无效的索引
    #[error("Index out of bounds: {index} (max: {max})")]
    IndexOutOfBounds { index: usize, max: usize },

    /// 不支持的功能
    #[error("Unsupported feature: {0}")]
    Unsupported(String),

    /// 编码错误
    #[error("Encoding error: {0}")]
    Encoding(String),
}

pub type Result<T> = std::result::Result<T, Error>;
```

## 10. 使用示例

### 10.1 基本读取

```rust
use linch_docx_rs::{Document, Result};

fn main() -> Result<()> {
    let doc = Document::open("example.docx")?;

    // 读取所有段落文本
    for para in doc.paragraphs() {
        println!("Style: {:?}", para.style());
        println!("Text: {}", para.text());

        // 读取每个 Run
        for run in para.runs() {
            print!("{}", run.text());
            if run.bold() {
                print!(" [BOLD]");
            }
        }
        println!();
    }

    // 读取表格
    for (i, table) in doc.tables().enumerate() {
        println!("Table {}: {}x{}", i, table.row_count(), table.column_count());
        for row in table.rows() {
            for cell in row.cells() {
                print!("| {} ", cell.text());
            }
            println!("|");
        }
    }

    Ok(())
}
```

### 10.2 创建文档

```rust
use linch_docx_rs::{Document, Pt, Alignment, Color, Result};

fn main() -> Result<()> {
    let mut doc = Document::new();

    // 添加标题
    doc.add_paragraph("项目报告")
        .style("Heading1")
        .alignment(Alignment::Center)
        .build();

    // 添加正文段落
    doc.add_paragraph("这是第一段正文内容。")
        .font_size(Pt(12.0))
        .build();

    // 添加带格式的段落
    let para_idx = doc.add_paragraph("").build();
    let para = doc.paragraph_mut(para_idx).unwrap();

    para.add_run("普通文本，").build();
    para.add_run("粗体文本，").bold(true).build();
    para.add_run("红色文本。").color(Color::Rgb("FF0000".into())).build();

    // 添加表格
    doc.add_table(3, 3)
        .header_rows(1)
        .data(&[
            &["姓名", "年龄", "城市"],
            &["张三", "25", "北京"],
            &["李四", "30", "上海"],
        ])
        .build();

    doc.save("output.docx")?;
    Ok(())
}
```

### 10.3 修改文档（模板填充）

```rust
use linch_docx_rs::{Document, Result};

fn main() -> Result<()> {
    let mut doc = Document::open("template.docx")?;

    // 查找替换
    doc.replace_text("{{name}}", "张三");
    doc.replace_text("{{date}}", "2024-01-15");
    doc.replace_text("{{company}}", "某某公司");

    // 修改特定段落
    for para in doc.paragraphs_mut() {
        if para.text().contains("重要提示") {
            para.set_style("Heading2");
            para.set_bold(true);
            para.set_color(Color::Rgb("FF0000".into()));
        }
    }

    // 修改表格
    if let Some(mut table) = doc.tables_mut().next() {
        if let Some(mut cell) = table.cell_mut(1, 0) {
            cell.set_text("更新后的内容");
        }
    }

    doc.save("filled.docx")?;
    Ok(())
}
```

### 10.4 高级用法

```rust
use linch_docx_rs::{Document, Result};

fn main() -> Result<()> {
    // 从字节创建
    let bytes = std::fs::read("template.docx")?;
    let mut doc = Document::from_bytes(&bytes)?;

    // 访问文档属性
    let props = doc.core_properties_mut();
    props.title = Some("生成的报告".into());
    props.creator = Some("linch-docx-rs".into());

    // 批量处理
    let locations = doc.find_text("TODO");
    println!("Found {} TODOs", locations.len());

    // 正则替换
    doc.replace_regex(r"\d{4}-\d{2}-\d{2}", "DATE_PLACEHOLDER")?;

    // 导出为字节
    let output = doc.to_bytes()?;
    std::fs::write("output.docx", output)?;

    Ok(())
}
```

## 11. 迭代器类型

```rust
/// 段落迭代器
pub struct Paragraphs<'a> { ... }
impl<'a> Iterator for Paragraphs<'a> {
    type Item = Paragraph<'a>;
}

/// 可变段落迭代器
pub struct ParagraphsMut<'a> { ... }
impl<'a> Iterator for ParagraphsMut<'a> {
    type Item = ParagraphMut<'a>;
}

/// Run 迭代器
pub struct Runs<'a> { ... }
impl<'a> Iterator for Runs<'a> {
    type Item = Run<'a>;
}

/// 表格迭代器
pub struct Tables<'a> { ... }
impl<'a> Iterator for Tables<'a> {
    type Item = Table<'a>;
}

/// 表格行迭代器
pub struct Rows<'a> { ... }
impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;
}

/// 单元格迭代器
pub struct Cells<'a> { ... }
impl<'a> Iterator for Cells<'a> {
    type Item = Cell<'a>;
}

/// 样式迭代器
pub struct Styles<'a> { ... }
impl<'a> Iterator for Styles<'a> {
    type Item = Style<'a>;
}
```

## 12. 下一步

1. 完成 [OPC 层设计](./opc-layer.md)
2. 完成 [OOXML 元素映射](./xml-mapping.md)
3. 完成 [Round-trip 策略](./roundtrip.md)
4. 完成 [错误处理设计](./error-handling.md)
