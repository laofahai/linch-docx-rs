# linch-docx-rs OOXML 元素映射文档

## 1. 概述

本文档定义 WordprocessingML (WML) XML 元素与 Rust 类型之间的映射关系。

**命名空间**：
- `w:` = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- `r:` = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
- `wp:` = `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`
- `a:` = `http://schemas.openxmlformats.org/drawingml/2006/main`

## 2. 文档结构元素

### 2.1 w:document（文档根元素）

```xml
<w:document xmlns:w="...">
  <w:body>
    <!-- 文档内容 -->
  </w:body>
</w:document>
```

```rust
/// 文档元素
pub struct DocumentElement {
    /// 文档主体
    pub body: BodyElement,

    /// 文档背景（可选）
    pub background: Option<BackgroundElement>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}
```

### 2.2 w:body（文档主体）

```xml
<w:body>
  <w:p>...</w:p>           <!-- 段落 -->
  <w:tbl>...</w:tbl>       <!-- 表格 -->
  <w:sdt>...</w:sdt>       <!-- 结构化文档标签 -->
  <w:sectPr>...</w:sectPr> <!-- 节属性（最后一个） -->
</w:body>
```

```rust
/// 文档主体元素
pub struct BodyElement {
    /// 块级内容（段落、表格等）
    pub content: Vec<BlockContent>,

    /// 最终节属性
    pub section_properties: Option<SectionProperties>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 块级内容
pub enum BlockContent {
    Paragraph(ParagraphElement),
    Table(TableElement),
    Sdt(SdtBlockElement),
    // 未来可扩展
}
```

## 3. 段落元素

### 3.1 w:p（段落）

```xml
<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3">
  <w:pPr>...</w:pPr>       <!-- 段落属性 -->
  <w:r>...</w:r>           <!-- Run -->
  <w:bookmarkStart/>       <!-- 书签开始 -->
  <w:bookmarkEnd/>         <!-- 书签结束 -->
  <w:commentRangeStart/>   <!-- 批注范围开始 -->
  <w:commentRangeEnd/>     <!-- 批注范围结束 -->
  <w:hyperlink>...</w:hyperlink>  <!-- 超链接 -->
</w:p>
```

```rust
/// 段落元素
pub struct ParagraphElement {
    /// 段落属性
    pub properties: Option<ParagraphProperties>,

    /// 段落内容
    pub content: Vec<ParagraphContent>,

    /// 段落 ID 属性
    pub rsid_r: Option<String>,
    pub rsid_r_default: Option<String>,
    pub rsid_p: Option<String>,
    pub rsid_r_pr: Option<String>,

    /// 未知属性
    pub unknown_attributes: Vec<(String, String)>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 段落内容
pub enum ParagraphContent {
    Run(RunElement),
    Hyperlink(HyperlinkElement),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    CommentRangeStart(CommentRangeStart),
    CommentRangeEnd(CommentRangeEnd),
    PermStart(PermStart),
    PermEnd(PermEnd),
    ProofError(ProofError),
    // 更多类型...
}
```

### 3.2 w:pPr（段落属性）

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>     <!-- 样式引用 -->
  <w:jc w:val="center"/>           <!-- 对齐方式 -->
  <w:ind w:left="720" w:firstLine="720"/>  <!-- 缩进 -->
  <w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>
  <w:numPr>...</w:numPr>           <!-- 编号属性 -->
  <w:tabs>...</w:tabs>             <!-- 制表位 -->
  <w:outlineLvl w:val="0"/>        <!-- 大纲级别 -->
  <w:rPr>...</w:rPr>               <!-- 运行属性（段落标记格式） -->
</w:pPr>
```

```rust
/// 段落属性
pub struct ParagraphProperties {
    /// 样式 ID
    pub style: Option<String>,

    /// 对齐方式
    pub justification: Option<Justification>,

    /// 缩进
    pub indentation: Option<Indentation>,

    /// 间距
    pub spacing: Option<Spacing>,

    /// 编号属性
    pub numbering: Option<NumberingProperties>,

    /// 制表位
    pub tabs: Option<Vec<TabStop>>,

    /// 大纲级别
    pub outline_level: Option<u8>,

    /// 保持下一段
    pub keep_next: bool,

    /// 保持段落完整
    pub keep_lines: bool,

    /// 段前分页
    pub page_break_before: bool,

    /// 孤行控制
    pub widow_control: Option<bool>,

    /// 边框
    pub borders: Option<ParagraphBorders>,

    /// 底纹
    pub shading: Option<Shading>,

    /// 文本方向
    pub text_direction: Option<TextDirection>,

    /// 段落标记的运行属性
    pub run_properties: Option<RunProperties>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 对齐方式
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Justification {
    Left,           // w:val="left"
    Center,         // w:val="center"
    Right,          // w:val="right"
    Both,           // w:val="both" (两端对齐)
    Distribute,     // w:val="distribute" (分散对齐)
}

/// 缩进
pub struct Indentation {
    /// 左缩进（twips）
    pub left: Option<i32>,
    /// 右缩进（twips）
    pub right: Option<i32>,
    /// 首行缩进（twips）
    pub first_line: Option<i32>,
    /// 悬挂缩进（twips）
    pub hanging: Option<i32>,
    /// 左缩进（字符）
    pub left_chars: Option<i32>,
    /// 右缩进（字符）
    pub right_chars: Option<i32>,
    /// 首行缩进（字符）
    pub first_line_chars: Option<i32>,
    /// 悬挂缩进（字符）
    pub hanging_chars: Option<i32>,
}

/// 间距
pub struct Spacing {
    /// 段前间距（twips）
    pub before: Option<u32>,
    /// 段后间距（twips）
    pub after: Option<u32>,
    /// 行距值
    pub line: Option<i32>,
    /// 行距规则
    pub line_rule: Option<LineRule>,
    /// 段前间距（行）
    pub before_lines: Option<i32>,
    /// 段后间距（行）
    pub after_lines: Option<i32>,
    /// 自动调整段前间距
    pub before_auto_spacing: Option<bool>,
    /// 自动调整段后间距
    pub after_auto_spacing: Option<bool>,
}

/// 行距规则
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineRule {
    Auto,       // 自动（line 值为 240 的倍数）
    Exact,      // 固定值
    AtLeast,    // 最小值
}

/// 编号属性
pub struct NumberingProperties {
    /// 编号级别
    pub level: u8,
    /// 编号定义 ID
    pub num_id: u32,
}
```

## 4. Run 元素

### 4.1 w:r（Run）

```xml
<w:r w:rsidR="00A1B2C3">
  <w:rPr>...</w:rPr>       <!-- Run 属性 -->
  <w:t>文本</w:t>          <!-- 文本 -->
  <w:t xml:space="preserve"> </w:t>  <!-- 保留空格的文本 -->
  <w:tab/>                 <!-- 制表符 -->
  <w:br/>                  <!-- 换行 -->
  <w:br w:type="page"/>    <!-- 分页符 -->
  <w:sym/>                 <!-- 符号 -->
  <w:drawing>...</w:drawing>  <!-- 图片 -->
</w:r>
```

```rust
/// Run 元素
pub struct RunElement {
    /// Run 属性
    pub properties: Option<RunProperties>,

    /// Run 内容
    pub content: Vec<RunContent>,

    /// RSID 属性
    pub rsid_r: Option<String>,
    pub rsid_r_pr: Option<String>,

    /// 未知属性
    pub unknown_attributes: Vec<(String, String)>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// Run 内容
pub enum RunContent {
    /// 文本 (w:t)
    Text(TextElement),

    /// 制表符 (w:tab)
    Tab,

    /// 换行/分页 (w:br)
    Break(BreakElement),

    /// 回车 (w:cr)
    CarriageReturn,

    /// 符号 (w:sym)
    Symbol(SymbolElement),

    /// 图片 (w:drawing)
    Drawing(DrawingElement),

    /// 字段开始 (w:fldChar type="begin")
    FieldCharBegin(FieldCharBegin),

    /// 字段分隔 (w:fldChar type="separate")
    FieldCharSeparate,

    /// 字段结束 (w:fldChar type="end")
    FieldCharEnd,

    /// 字段代码 (w:instrText)
    InstrText(String),

    /// 脚注引用 (w:footnoteReference)
    FootnoteReference(FootnoteReference),

    /// 尾注引用 (w:endnoteReference)
    EndnoteReference(EndnoteReference),

    /// 批注引用 (w:commentReference)
    CommentReference(CommentReference),

    /// 软换行 (w:softHyphen)
    SoftHyphen,

    /// 不换行连字符 (w:noBreakHyphen)
    NoBreakHyphen,
}

/// 文本元素
pub struct TextElement {
    pub text: String,
    /// xml:space="preserve"
    pub preserve_space: bool,
}

/// 换行元素
pub struct BreakElement {
    pub break_type: BreakType,
    /// 文本环绕的分栏换行 clear 属性
    pub clear: Option<BreakClear>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum BreakType {
    #[default]
    TextWrapping,  // 默认，软换行
    Page,          // w:type="page"
    Column,        // w:type="column"
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakClear {
    None,
    Left,
    Right,
    All,
}
```

### 4.2 w:rPr（Run 属性）

```xml
<w:rPr>
  <w:rStyle w:val="Strong"/>      <!-- 样式引用 -->
  <w:b/>                          <!-- 粗体 -->
  <w:b w:val="false"/>           <!-- 非粗体 -->
  <w:i/>                          <!-- 斜体 -->
  <w:u w:val="single"/>          <!-- 下划线 -->
  <w:strike/>                     <!-- 删除线 -->
  <w:dstrike/>                    <!-- 双删除线 -->
  <w:sz w:val="24"/>             <!-- 字号（半点） -->
  <w:szCs w:val="24"/>           <!-- 复杂脚本字号 -->
  <w:color w:val="FF0000"/>      <!-- 颜色 -->
  <w:highlight w:val="yellow"/>  <!-- 高亮 -->
  <w:rFonts w:ascii="Arial" w:eastAsia="宋体"/>  <!-- 字体 -->
  <w:vertAlign w:val="superscript"/>  <!-- 上下标 -->
  <w:lang w:val="en-US"/>        <!-- 语言 -->
</w:rPr>
```

```rust
/// Run 属性
pub struct RunProperties {
    /// 样式 ID
    pub style: Option<String>,

    /// 粗体
    pub bold: Option<bool>,
    /// 复杂脚本粗体
    pub bold_cs: Option<bool>,

    /// 斜体
    pub italic: Option<bool>,
    /// 复杂脚本斜体
    pub italic_cs: Option<bool>,

    /// 下划线
    pub underline: Option<UnderlineType>,

    /// 删除线
    pub strike: Option<bool>,

    /// 双删除线
    pub double_strike: Option<bool>,

    /// 字号（半点单位，如 24 = 12pt）
    pub size: Option<u32>,
    /// 复杂脚本字号
    pub size_cs: Option<u32>,

    /// 颜色
    pub color: Option<Color>,

    /// 高亮
    pub highlight: Option<HighlightColor>,

    /// 字体
    pub fonts: Option<Fonts>,

    /// 垂直对齐（上下标）
    pub vertical_align: Option<VerticalAlign>,

    /// 字符间距
    pub spacing: Option<i32>,

    /// 位置（升降）
    pub position: Option<i32>,

    /// 字符缩放
    pub w: Option<u32>,

    /// 字距调整
    pub kern: Option<u32>,

    /// 隐藏
    pub vanish: Option<bool>,

    /// 小型大写
    pub small_caps: Option<bool>,

    /// 全部大写
    pub caps: Option<bool>,

    /// 阴影
    pub shadow: Option<bool>,

    /// 浮雕
    pub emboss: Option<bool>,

    /// 印刻
    pub imprint: Option<bool>,

    /// 轮廓
    pub outline: Option<bool>,

    /// 语言
    pub lang: Option<Language>,

    /// 边框
    pub border: Option<Border>,

    /// 底纹
    pub shading: Option<Shading>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 下划线类型
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum UnderlineType {
    Single,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dash,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
    Words,
    None,
}

/// 颜色
#[derive(Clone, Debug, PartialEq)]
pub struct Color {
    /// RGB 值（如 "FF0000"）
    pub val: Option<String>,
    /// 主题颜色
    pub theme_color: Option<ThemeColor>,
    /// 主题色调
    pub theme_tint: Option<String>,
    /// 主题阴影
    pub theme_shade: Option<String>,
}

/// 主题颜色
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
    Background1,
    Text1,
    Background2,
    Text2,
}

/// 高亮颜色
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HighlightColor {
    Black,
    Blue,
    Cyan,
    Green,
    Magenta,
    Red,
    Yellow,
    White,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
    None,
}

/// 字体
pub struct Fonts {
    /// ASCII 字体
    pub ascii: Option<String>,
    /// 高 ANSI 字体
    pub h_ansi: Option<String>,
    /// 东亚字体
    pub east_asia: Option<String>,
    /// 复杂脚本字体
    pub cs: Option<String>,
    /// ASCII 主题字体
    pub ascii_theme: Option<FontTheme>,
    /// 东亚主题字体
    pub east_asia_theme: Option<FontTheme>,
}

/// 字体主题
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FontTheme {
    MajorAscii,
    MajorEastAsia,
    MajorHAnsi,
    MajorBidi,
    MinorAscii,
    MinorEastAsia,
    MinorHAnsi,
    MinorBidi,
}

/// 垂直对齐
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

/// 语言
pub struct Language {
    pub val: Option<String>,
    pub east_asia: Option<String>,
    pub bidi: Option<String>,
}
```

## 5. 表格元素

### 5.1 w:tbl（表格）

```xml
<w:tbl>
  <w:tblPr>...</w:tblPr>     <!-- 表格属性 -->
  <w:tblGrid>                 <!-- 表格网格 -->
    <w:gridCol w:w="2880"/>
    <w:gridCol w:w="2880"/>
  </w:tblGrid>
  <w:tr>...</w:tr>           <!-- 表格行 -->
</w:tbl>
```

```rust
/// 表格元素
pub struct TableElement {
    /// 表格属性
    pub properties: Option<TableProperties>,

    /// 表格网格
    pub grid: Vec<GridColumn>,

    /// 表格行
    pub rows: Vec<TableRowElement>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 网格列
pub struct GridColumn {
    /// 列宽（twips）
    pub width: Option<i32>,
}

/// 表格属性
pub struct TableProperties {
    /// 样式 ID
    pub style: Option<String>,

    /// 表格宽度
    pub width: Option<TableWidth>,

    /// 表格对齐
    pub justification: Option<TableJustification>,

    /// 表格缩进
    pub indent: Option<i32>,

    /// 单元格边距（默认）
    pub cell_margin: Option<TableCellMargin>,

    /// 边框
    pub borders: Option<TableBorders>,

    /// 底纹
    pub shading: Option<Shading>,

    /// 表格布局
    pub layout: Option<TableLayout>,

    /// 外观选项（条带等）
    pub look: Option<TableLook>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 表格宽度
pub struct TableWidth {
    /// 宽度值
    pub width: i32,
    /// 宽度类型
    pub width_type: TableWidthType,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableWidthType {
    Auto,     // w:type="auto"
    Dxa,      // w:type="dxa" (twips)
    Nil,      // w:type="nil"
    Pct,      // w:type="pct" (百分比 * 50)
}

/// 表格对齐
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableJustification {
    Left,
    Center,
    Right,
}

/// 表格布局
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableLayout {
    Fixed,
    AutoFit,
}
```

### 5.2 w:tr（表格行）

```xml
<w:tr>
  <w:trPr>...</w:trPr>       <!-- 行属性 -->
  <w:tc>...</w:tc>           <!-- 单元格 -->
</w:tr>
```

```rust
/// 表格行元素
pub struct TableRowElement {
    /// 行属性
    pub properties: Option<TableRowProperties>,

    /// 单元格
    pub cells: Vec<TableCellElement>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 表格行属性
pub struct TableRowProperties {
    /// 行高
    pub height: Option<RowHeight>,

    /// 表头行
    pub header: bool,

    /// 允许跨页断行
    pub cant_split: Option<bool>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 行高
pub struct RowHeight {
    /// 高度值（twips）
    pub val: u32,
    /// 高度规则
    pub rule: Option<HeightRule>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum HeightRule {
    #[default]
    Auto,
    Exact,
    AtLeast,
}
```

### 5.3 w:tc（表格单元格）

```xml
<w:tc>
  <w:tcPr>
    <w:tcW w:w="2880" w:type="dxa"/>  <!-- 单元格宽度 -->
    <w:gridSpan w:val="2"/>           <!-- 跨列 -->
    <w:vMerge w:val="restart"/>       <!-- 垂直合并开始 -->
    <w:vMerge/>                       <!-- 垂直合并继续 -->
    <w:vAlign w:val="center"/>        <!-- 垂直对齐 -->
    <w:shd w:val="clear" w:fill="FFFF00"/>  <!-- 底纹 -->
  </w:tcPr>
  <w:p>...</w:p>                      <!-- 单元格内容（段落） -->
</w:tc>
```

```rust
/// 表格单元格元素
pub struct TableCellElement {
    /// 单元格属性
    pub properties: Option<TableCellProperties>,

    /// 单元格内容（块级元素）
    pub content: Vec<BlockContent>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 表格单元格属性
pub struct TableCellProperties {
    /// 单元格宽度
    pub width: Option<TableWidth>,

    /// 跨列数
    pub grid_span: Option<u32>,

    /// 垂直合并
    pub v_merge: Option<VMerge>,

    /// 水平合并（已废弃，但仍需支持）
    pub h_merge: Option<HMerge>,

    /// 垂直对齐
    pub vertical_align: Option<VerticalAlignment>,

    /// 底纹
    pub shading: Option<Shading>,

    /// 边框
    pub borders: Option<CellBorders>,

    /// 单元格边距
    pub margins: Option<CellMargins>,

    /// 文本方向
    pub text_direction: Option<TextDirection>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 垂直合并
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VMerge {
    Restart,    // 合并开始
    Continue,   // 合并继续（默认，无 val 属性）
}

/// 水平合并（已废弃）
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HMerge {
    Restart,
    Continue,
}

/// 垂直对齐
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
}

/// 文本方向
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextDirection {
    LrTb,       // 从左到右，从上到下
    TbRl,       // 从上到下，从右到左
    BtLr,       // 从下到上，从左到右
    LrTbV,
    TbRlV,
    TbLrV,
}
```

## 6. 样式元素 (styles.xml)

### 6.1 w:styles（样式根元素）

```xml
<w:styles>
  <w:docDefaults>...</w:docDefaults>  <!-- 文档默认值 -->
  <w:style>...</w:style>              <!-- 样式定义 -->
</w:styles>
```

```rust
/// 样式文档
pub struct StylesDocument {
    /// 文档默认值
    pub doc_defaults: Option<DocDefaults>,

    /// 潜在样式
    pub latent_styles: Option<LatentStyles>,

    /// 样式列表
    pub styles: Vec<StyleElement>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 文档默认值
pub struct DocDefaults {
    /// 默认运行属性
    pub run_properties_default: Option<RunPropertiesDefault>,

    /// 默认段落属性
    pub paragraph_properties_default: Option<ParagraphPropertiesDefault>,
}
```

### 6.2 w:style（样式定义）

```xml
<w:style w:type="paragraph" w:styleId="Heading1" w:default="1">
  <w:name w:val="heading 1"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="Normal"/>
  <w:link w:val="Heading1Char"/>
  <w:uiPriority w:val="9"/>
  <w:qFormat/>
  <w:pPr>...</w:pPr>
  <w:rPr>...</w:rPr>
</w:style>
```

```rust
/// 样式元素
pub struct StyleElement {
    /// 样式类型
    pub style_type: StyleType,

    /// 样式 ID
    pub style_id: String,

    /// 是否为默认样式
    pub default: bool,

    /// 样式名称
    pub name: Option<String>,

    /// 基于的样式 ID
    pub based_on: Option<String>,

    /// 下一段落样式 ID
    pub next: Option<String>,

    /// 链接的样式 ID
    pub link: Option<String>,

    /// UI 优先级
    pub ui_priority: Option<i32>,

    /// 是否在快速样式库中显示
    pub qformat: bool,

    /// 是否隐藏
    pub hidden: bool,

    /// 段落属性
    pub paragraph_properties: Option<ParagraphProperties>,

    /// 运行属性
    pub run_properties: Option<RunProperties>,

    /// 表格属性（表格样式）
    pub table_properties: Option<TableProperties>,

    /// 表格行属性（表格样式）
    pub table_row_properties: Option<TableRowProperties>,

    /// 表格单元格属性（表格样式）
    pub table_cell_properties: Option<TableCellProperties>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 样式类型
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}
```

## 7. 编号元素 (numbering.xml)

### 7.1 编号定义

```xml
<w:numbering>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>...</w:pPr>
      <w:rPr>...</w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>
```

```rust
/// 编号文档
pub struct NumberingDocument {
    /// 抽象编号定义
    pub abstract_nums: Vec<AbstractNum>,

    /// 编号实例
    pub nums: Vec<Num>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 抽象编号
pub struct AbstractNum {
    /// 抽象编号 ID
    pub abstract_num_id: u32,

    /// 多级列表样式
    pub multi_level_type: Option<MultiLevelType>,

    /// 级别定义
    pub levels: Vec<Level>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 级别定义
pub struct Level {
    /// 级别索引（0-8）
    pub ilvl: u8,

    /// 起始值
    pub start: Option<u32>,

    /// 编号格式
    pub num_fmt: Option<NumberFormat>,

    /// 级别文本
    pub level_text: Option<String>,

    /// 对齐方式
    pub justification: Option<Justification>,

    /// 段落属性
    pub paragraph_properties: Option<ParagraphProperties>,

    /// 运行属性
    pub run_properties: Option<RunProperties>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 编号格式
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NumberFormat {
    Decimal,           // 1, 2, 3
    UpperRoman,        // I, II, III
    LowerRoman,        // i, ii, iii
    UpperLetter,       // A, B, C
    LowerLetter,       // a, b, c
    Bullet,            // •
    ChineseCountingThousand,  // 一, 二, 三
    // ... 更多格式
    None,
}

/// 编号实例
pub struct Num {
    /// 编号 ID
    pub num_id: u32,

    /// 引用的抽象编号 ID
    pub abstract_num_id: u32,

    /// 级别覆盖
    pub level_overrides: Vec<LevelOverride>,
}

/// 级别覆盖
pub struct LevelOverride {
    /// 级别索引
    pub ilvl: u8,

    /// 起始值覆盖
    pub start_override: Option<u32>,

    /// 级别定义覆盖
    pub level: Option<Level>,
}
```

## 8. 图片与绘图元素

### 8.1 w:drawing（绘图容器）

```xml
<w:drawing>
  <wp:inline>                    <!-- 行内图片 -->
    <wp:extent cx="914400" cy="914400"/>
    <wp:docPr id="1" name="Picture 1"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:blipFill>
            <a:blip r:embed="rId4"/>  <!-- 引用关系 ID -->
          </pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

```rust
/// 绘图元素
pub struct DrawingElement {
    /// 绘图内容
    pub content: DrawingContent,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 绘图内容
pub enum DrawingContent {
    /// 行内图片
    Inline(InlineDrawing),
    /// 浮动图片
    Anchor(AnchorDrawing),
}

/// 行内绘图
pub struct InlineDrawing {
    /// 图片尺寸
    pub extent: Extent,

    /// 文档属性
    pub doc_properties: DocProperties,

    /// 图形数据
    pub graphic: Graphic,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 尺寸
pub struct Extent {
    /// 宽度（EMU）
    pub cx: i64,
    /// 高度（EMU）
    pub cy: i64,
}

/// 文档属性
pub struct DocProperties {
    pub id: u32,
    pub name: String,
    pub description: Option<String>,
}

/// 图形
pub struct Graphic {
    /// 图形数据
    pub data: GraphicData,
}

/// 图形数据
pub enum GraphicData {
    /// 图片
    Picture(Picture),
    /// 其他类型（保留原始 XML）
    Other(RawXmlElement),
}

/// 图片
pub struct Picture {
    /// 非视觉属性
    pub non_visual_properties: Option<PictureNonVisualProperties>,

    /// 填充
    pub blip_fill: BlipFill,

    /// 形状属性
    pub shape_properties: Option<ShapeProperties>,
}

/// Blip 填充
pub struct BlipFill {
    /// Blip（图片数据引用）
    pub blip: Blip,

    /// 拉伸
    pub stretch: Option<Stretch>,
}

/// Blip
pub struct Blip {
    /// 嵌入的关系 ID
    pub embed: Option<String>,

    /// 链接的关系 ID
    pub link: Option<String>,
}
```

## 9. 超链接元素

```xml
<w:hyperlink r:id="rId5" w:history="1">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
    </w:rPr>
    <w:t>链接文本</w:t>
  </w:r>
</w:hyperlink>
```

```rust
/// 超链接元素
pub struct HyperlinkElement {
    /// 关系 ID（外部链接）
    pub r_id: Option<String>,

    /// 锚点（文档内部链接）
    pub anchor: Option<String>,

    /// 工具提示
    pub tooltip: Option<String>,

    /// 历史记录
    pub history: Option<bool>,

    /// 内容（Run 列表）
    pub content: Vec<RunElement>,

    /// 未知属性
    pub unknown_attributes: Vec<(String, String)>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}
```

## 10. 节属性元素

```xml
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>        <!-- 页面大小 -->
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>  <!-- 页边距 -->
  <w:cols w:space="720"/>                   <!-- 分栏 -->
  <w:docGrid w:linePitch="360"/>           <!-- 文档网格 -->
  <w:headerReference w:type="default" r:id="rId6"/>
  <w:footerReference w:type="default" r:id="rId7"/>
</w:sectPr>
```

```rust
/// 节属性
pub struct SectionProperties {
    /// 页面大小
    pub page_size: Option<PageSize>,

    /// 页边距
    pub page_margin: Option<PageMargin>,

    /// 分栏
    pub columns: Option<Columns>,

    /// 文档网格
    pub doc_grid: Option<DocGrid>,

    /// 页眉引用
    pub header_references: Vec<HeaderFooterReference>,

    /// 页脚引用
    pub footer_references: Vec<HeaderFooterReference>,

    /// 节类型
    pub section_type: Option<SectionType>,

    /// 页码设置
    pub page_numbering: Option<PageNumbering>,

    /// 未知子元素
    pub unknown_children: Vec<RawXmlNode>,
}

/// 页面大小
pub struct PageSize {
    /// 宽度（twips）
    pub width: u32,
    /// 高度（twips）
    pub height: u32,
    /// 方向
    pub orient: Option<PageOrientation>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// 页边距
pub struct PageMargin {
    pub top: i32,
    pub right: u32,
    pub bottom: i32,
    pub left: u32,
    pub header: u32,
    pub footer: u32,
    pub gutter: u32,
}

/// 节类型
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SectionType {
    Continuous,
    NextPage,
    EvenPage,
    OddPage,
    NextColumn,
}

/// 页眉/页脚引用
pub struct HeaderFooterReference {
    /// 类型
    pub header_footer_type: HeaderFooterType,
    /// 关系 ID
    pub r_id: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HeaderFooterType {
    Default,
    First,
    Even,
}
```

## 11. 解析与序列化 trait

```rust
/// XML 解析 trait
pub trait FromXml: Sized {
    /// 从 XML 事件解析
    fn from_xml(reader: &mut XmlReader) -> Result<Self>;
}

/// XML 序列化 trait
pub trait ToXml {
    /// 序列化为 XML
    fn to_xml(&self, writer: &mut XmlWriter) -> Result<()>;
}

/// 通用解析辅助宏
macro_rules! impl_from_xml {
    // ... 实现细节
}

/// 属性解析辅助
pub fn parse_bool_attr(value: &str) -> bool {
    match value {
        "1" | "true" | "on" => true,
        _ => false,
    }
}

pub fn parse_on_off(element: &BytesStart) -> Option<bool> {
    match get_attr_opt(element, "val") {
        None => Some(true),  // 无 val 属性表示 true
        Some(v) => Some(parse_bool_attr(&v)),
    }
}
```

## 12. 下一步

1. 实现核心元素的解析代码
2. 实现 Round-trip 保真机制
3. 编写测试用例
