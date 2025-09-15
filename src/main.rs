use kimi::*;

// 这是一份从互联网上下载的《中华人民共和国职业分类大典》，不难看出，
// 这是对某个 PDF 文件使用 OCR 技术转换而成的 Excel 文件。
//
// 请写一份代码，将上述 Excel 文件中的职业分类按照大类、中类、小类、细类的层级整理成树状结构化数据，
// 并以 JSON 格式（或你喜欢的其他格式）输出。
//

// run: cargo run -- ./202306151255033.xlsx
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <excel_file.xlsx>", args[0]);
        std::process::exit(1);
    }
    let input_file = &args[1];

    let mut tree = CategoryTree::new();
    tree.build_from(input_file)?;
    tree.pretty_print();
    // json format
    // tree.pretty_print_json();

    Ok(())
}
