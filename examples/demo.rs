//! 功能演示：展示 linch-docx-rs 的主要功能
//!
//! 运行方法: cargo run --example demo

use linch_docx_rs::{Alignment, Document, Indentation, LineSpacing, Run, Table};

fn main() -> linch_docx_rs::Result<()> {
    println!("=== linch-docx-rs 功能演示 ===\n");

    // -------------------------------------------------------
    // 1. 创建新文档
    // -------------------------------------------------------
    println!("1. 创建新文档并写入内容...");
    let mut doc = Document::new();

    // 设置页面为 A4
    doc.section_properties_mut().set_a4_portrait();

    // 设置文档属性
    let props = doc.core_properties_mut();
    props.title = Some("linch-docx-rs 演示文档".into());
    props.creator = Some("linch-docx-rs".into());

    // 添加标题
    let heading = doc.add_paragraph("linch-docx-rs 功能演示");
    heading.set_style("Heading1");
    heading.set_alignment(Alignment::Center);

    // 添加正文
    let p1 = doc.add_paragraph("这是一个由 Rust 生成的 DOCX 文档。");
    p1.set_spacing(LineSpacing {
        after: Some(200),
        ..Default::default()
    });

    // 添加带格式的段落
    let p2 = doc.add_paragraph("");
    p2.clear();
    let mut bold_run = Run::new("粗体文本、");
    bold_run.set_bold(true);
    p2.add_run(bold_run);

    let mut italic_run = Run::new("斜体文本、");
    italic_run.set_italic(true);
    p2.add_run(italic_run);

    let mut colored_run = Run::new("红色文本。");
    colored_run.set_color("FF0000");
    p2.add_run(colored_run);

    // 添加带缩进的段落
    let p3 = doc.add_paragraph("这是一个有首行缩进的段落。用于展示段落格式设置功能。");
    p3.set_indentation(Indentation {
        first_line: Some(480), // 首行缩进 2 字符 (约480 twips)
        ..Default::default()
    });

    // 添加超链接
    let p4 = doc.add_paragraph("");
    p4.clear();
    p4.add_run(Run::new("访问项目："));
    p4.add_internal_link("_top", "返回顶部");

    // 添加表格
    let table = Table::new(3, 3);
    let table_ref = doc.add_table(table);
    if let Some(cell) = table_ref.cell_mut(0, 0) {
        cell.set_text("姓名");
    }
    if let Some(cell) = table_ref.cell_mut(0, 1) {
        cell.set_text("年龄");
    }
    if let Some(cell) = table_ref.cell_mut(0, 2) {
        cell.set_text("城市");
    }
    if let Some(cell) = table_ref.cell_mut(1, 0) {
        cell.set_text("张三");
    }
    if let Some(cell) = table_ref.cell_mut(1, 1) {
        cell.set_text("25");
    }
    if let Some(cell) = table_ref.cell_mut(1, 2) {
        cell.set_text("北京");
    }
    if let Some(cell) = table_ref.cell_mut(2, 0) {
        cell.set_text("李四");
    }
    if let Some(cell) = table_ref.cell_mut(2, 1) {
        cell.set_text("30");
    }
    if let Some(cell) = table_ref.cell_mut(2, 2) {
        cell.set_text("上海");
    }

    // 保存
    doc.save("demo_output.docx")?;
    println!("  -> 已保存: demo_output.docx\n");

    // -------------------------------------------------------
    // 2. 读取已有文档
    // -------------------------------------------------------
    println!("2. 读取测试文档...");
    let doc = Document::open("tests/fixtures/simple.docx")?;

    println!("  段落数: {}", doc.paragraph_count());
    println!("  表格数: {}", doc.table_count());

    for (i, para) in doc.paragraphs().enumerate() {
        let style = para.style().unwrap_or("(无样式)");
        let heading = if para.is_heading() { " [标题]" } else { "" };
        println!("  段落{}: [{}]{} \"{}\"", i, style, heading, para.text());

        for run in para.runs() {
            let mut fmt = Vec::new();
            if run.bold() {
                fmt.push("粗体");
            }
            if run.italic() {
                fmt.push("斜体");
            }
            if let Some(size) = run.font_size_pt() {
                fmt.push(if size == 14.0 { "14pt" } else { "其他字号" });
            }
            if !fmt.is_empty() {
                println!("    Run: \"{}\" [{}]", run.text(), fmt.join(", "));
            }
        }
    }

    // 文档属性
    if let Some(props) = doc.core_properties() {
        println!("\n  文档属性:");
        if let Some(ref title) = props.title {
            println!("    标题: {}", title);
        }
        if let Some(ref creator) = props.creator {
            println!("    作者: {}", creator);
        }
    }

    // 样式
    if let Some(styles) = doc.styles() {
        println!("\n  样式数: {}", styles.iter().count());
        for s in styles.iter().take(5) {
            println!("    - {} ({:?})", s.style_id, s.style_type);
        }
    }

    // 节属性
    if let Some(sect) = doc.section_properties() {
        if let Some(ref pg) = sect.page_size {
            println!(
                "\n  页面大小: {}x{} twips ({:?})",
                pg.width.unwrap_or(0),
                pg.height.unwrap_or(0),
                pg.orient
            );
        }
    }

    println!();

    // -------------------------------------------------------
    // 3. 修改文档（模板填充）
    // -------------------------------------------------------
    println!("3. 修改文档（文本替换）...");
    let mut doc = Document::open("tests/fixtures/simple.docx")?;
    let count = doc.replace_text("Hello, World!", "你好，世界！");
    println!("  替换了 {} 处文本", count);

    // 查找文本
    let locations = doc.find_text("heading");
    println!("  找到 {} 处 'heading'", locations.len());

    // 添加段落
    let new_para = doc.add_paragraph("这是新追加的段落。");
    new_para.set_alignment(Alignment::Center);

    doc.save("demo_modified.docx")?;
    println!("  -> 已保存: demo_modified.docx\n");

    // -------------------------------------------------------
    // 4. 单位换算
    // -------------------------------------------------------
    println!("4. 单位换算演示:");
    use linch_docx_rs::{Cm, Inch, Pt, Twip};

    let pt = Pt(12.0);
    let twip = Twip::from(pt);
    let inch = Inch::from(pt);
    let cm = Cm::from(pt);
    println!("  {} = {} = {} = {}", pt, twip, inch, cm);

    let a4_width = Cm(21.0);
    let a4_twip = Twip::from(a4_width);
    println!("  A4 宽度: {} = {}", a4_width, a4_twip);

    println!("\n=== 演示完成 ===");

    // 保留生成的文件，方便查看
    println!("生成的文件:");
    println!("  demo_output.docx   - 全新创建的文档");
    println!("  demo_modified.docx - 修改后的文档");

    Ok(())
}
