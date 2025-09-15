use indexmap::IndexMap;
use regex::Regex;
use serde::{Deserialize, Serialize};
use serde_json;
use std::error::Error;
use umya_spreadsheet::reader::xlsx;

/// Define Category
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Category {
    pub id: String,           // e.g. "1-01"
    pub code: Option<String>, // e.g. "GBM10100"
    pub desc: Option<String>, // e.g. "中国共产党机关和基层组织负责人"
}

/// Hierarchical tree structure
#[derive(Debug, Serialize, Deserialize)]
pub struct CategoryTree {
    children: IndexMap<String, CategoryTree>,
    categories: Vec<Category>,
}

impl CategoryTree {
    pub fn new() -> Self {
        Self {
            children: IndexMap::new(),
            categories: Vec::new(),
        }
    }

    pub fn insert(&mut self, id: String, category: Category) {
        let mut node = self;
        let path = id.split('-').collect::<Vec<&str>>();
        for level in path {
            node = node
                .children
                .entry(level.to_string())
                .or_insert_with(CategoryTree::new);
        }
        node.categories.push(category);
    }

    pub fn pretty_print_json(&self) {
        let json = serde_json::to_string_pretty(&self).unwrap();
        println!("{}", json);
    }

    pub fn pretty_print(&self) {
        self.pretty_print_inner("", true);
    }

    fn pretty_print_inner(&self, prefix: &str, is_last: bool) {
        let branch = if is_last { "└── " } else { "├── " };

        for (i, cat) in self.categories.iter().enumerate() {
            let connector = if i == self.categories.len() - 1 && self.children.is_empty() {
                "└── "
            } else {
                "├── "
            };
            if let Some(code) = &cat.code {
                println!(
                    "{}{}{} [{}, {}]",
                    prefix,
                    connector,
                    cat.id,
                    code,
                    cat.desc.as_ref().map_or("", |v| v)
                );
            } else {
                println!(
                    "{}{}{} [{}]",
                    prefix,
                    connector,
                    cat.id,
                    cat.desc.as_ref().map_or("", |v| v)
                );
            }
        }

        let child_count = self.children.len();
        for (i, (key, child)) in self.children.iter().enumerate() {
            let is_last_child = i == child_count - 1;
            println!("{}{}{}", prefix, branch, key);

            let new_prefix = if is_last {
                format!("{}    ", prefix)
            } else {
                format!("{}│   ", prefix)
            };
            child.pretty_print_inner(&new_prefix, is_last_child);
        }
    }

    pub fn parse_one_column(&mut self, cell_text: &str) -> Result<(), Box<dyn Error>> {
        let chunks = construct_lines(cell_text);
        let parsed = parse_categories(&chunks)?;
        for cat in parsed {
            let cat_clone = cat.clone();
            self.insert(cat.id, cat_clone);
        }
        Ok(())
    }

    pub fn parse_two_columns(
        &mut self,
        cell_first: &str,
        cell_second: &str,
    ) -> Result<(), Box<dyn Error>> {
        let lines_first = construct_lines(cell_first);
        let lines_second = construct_lines(cell_second);

        // if lines_first.len() != lines_second.len() {
        //     println!(
        //         "{}, {}, {:?}, {:?}",
        //         lines_first.len(),
        //         lines_second.len(),
        //         lines_first,
        //         lines_second
        //     );
        //     return Ok(());
        // }

        // Zip the lines together and concatenate each pair
        let mut concatenated_lines = Vec::new();
        for (a, b) in lines_first.iter().zip(lines_second.iter()) {
            let combined = format!("{} {}", a, b).trim().to_string();
            if !combined.is_empty() {
                concatenated_lines.push(combined);
            }
        }

        let final_text = concatenated_lines.join("\n\n\n");
        let chunks = construct_lines(&final_text);
        let parsed = parse_categories(&chunks)?;
        for cat in parsed {
            let cat_clone = cat.clone();
            self.insert(cat.id, cat_clone);
        }

        Ok(())
    }

    pub fn build_from(&mut self, input_file: &str) -> Result<(), Box<dyn std::error::Error>> {
        let mut book = xlsx::read(input_file)?;
        let sheet = book.get_sheet_mut(&0).unwrap();

        let max_row = sheet.get_highest_row();
        let max_col = sheet.get_highest_column();

        for row in 1..=max_row {
            for col in 1..=max_col {
                let cell_value = sheet.get_cell_value((col, row));
                let cell_text = cell_value.get_value().to_string();

                if col == 1 {
                    if let Some(first_text) = normalize_first_category(&cell_text) {
                        self.parse_one_column(&first_text)?;
                    } else {
                        self.parse_one_column(&cell_text)?;
                    }
                } else if col == 3 {
                    self.parse_one_column(&cell_text)?;
                } else if col == 5 {
                    let cell_first = sheet.get_cell_value((col, row));
                    let cell_first = cell_first.get_value().to_string();

                    let cell_second = sheet.get_cell_value((col + 1, row));
                    let cell_second = cell_second.get_value().to_string();
                    self.parse_two_columns(cell_first.trim(), cell_second.trim())?;
                }
            }
        }

        Ok(())
    }
}

/// Parse categories
pub fn parse_categories(chunks: &Vec<String>) -> Result<Vec<Category>, Box<dyn Error>> {
    let mut categories = Vec::new();

    // Regex:
    // - id: one or more numbers separated by '-' at the start
    // - optional code: (GBM digits)
    // - description: rest of the string
    let re = Regex::new(
        r"(?x)
        ^\s*
        (?P<id>(?:\d+-?)+)          # id: 1-01 or 1-01-01-01
        (?:\s*\(\s*(?P<code>GBM\s*\d+)\s*\))?  # optional code
        \s*(?P<desc>.*)?$            # description
    ",
    )?;

    for chunk in chunks {
        if let Some(cap) = re.captures(chunk) {
            let id = cap
                .name("id")
                .map(|m| m.as_str().trim().to_string())
                .unwrap_or_default();
            let code = cap.name("code").map(|m| m.as_str().replace(' ', ""));
            let desc = cap.name("desc").map(|m| m.as_str().replace(' ', ""));
            // println!("{}\n", chunk);
            categories.push(Category { id, code, desc });
        }
    }

    Ok(categories)
}

/// Construct lines
/// TODO: this function is tricy to make it robust
pub fn construct_lines(text: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut buffer = String::new();

    let lines: Vec<String> = text
        .lines()
        .map(|l| l.chars().filter(|c| !c.is_whitespace()).collect())
        .collect();

    let mut i = 0;
    while i < lines.len() {
        let line = &lines[i].replace("L", "").replace("S", "").replace("/", "");
        let line = line.trim();
        // println!("{} {}", line, line.len());
        if line.is_empty()
            || line.ends_with("责人")
            || line.ends_with("员")
            || line.ends_with("护士")
            || line.ends_with("制片人")
            || line.ends_with("师")
            || line.ends_with("官")
            || line.ends_with("律师")
            || line.ends_with("医生")
            || line.ends_with("顾问")
            || line.ends_with("教师")
            || line.ends_with("警察")
            || line.ends_with("经理")
            || line.ends_with("董事")
            || (line.ends_with("工") && i + 1 < lines.len() && !lines[i + 1].contains("技术人员"))
            || line.matches('-').count() == 3
        {
            buffer.push_str(line);
            if !buffer.is_empty() {
                // println!("{}", buffer);
                result.push(buffer.clone());
            }
            buffer.clear();
            i += 1;
            continue;
        } else {
            buffer.push_str(line);
            i += 1;
        }
    }

    if !buffer.is_empty() {
        result.push(buffer);
    }

    result
}

pub fn normalize_first_category(text: &str) -> Option<String> {
    // Find first digit
    let first_digit_idx = text
        .char_indices()
        .find(|(_, c)| c.is_ascii_digit())
        .map(|(i, _)| i)?;

    let name = text.get(..first_digit_idx)?.trim();
    if name.len() == 0 {
        return None;
    }

    // Find first ')'
    let first_paren_idx = text.find(')')?;
    let desc = text.get(first_paren_idx + 1..)?.trim();
    let number = text.get(first_digit_idx..=first_paren_idx)?;
    Some(format!("{} {}{}", number, name, desc))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_merge() {
        let text = "中国共产党机关\n负责人\n中国共产党基层\n组织负责人";
        let result = construct_lines(text);
        let expected = vec![
            "中国共产党机关负责人".to_string(),
            "中国共产党基层组织负责人".to_string(),
        ];
        assert_eq!(result, expected);
    }

    #[test]
    fn test_empty_lines() {
        let text = "国家权力机关负\n责人\n\n国家行政机关负\n责人";
        let result = construct_lines(text);
        let expected = vec![
            "国家权力机关负责人".to_string(),
            "国家行政机关负责人".to_string(),
        ];
        assert_eq!(result, expected);
    }

    #[test]
    fn test_mixed_lines() {
        let text = "中国共产党机关\n负责人\n国家权力机关负\n责人\n民主党派负责人";
        let result = construct_lines(text);
        let expected = vec![
            "中国共产党机关负责人".to_string(),
            "国家权力机关负责人".to_string(),
            "民主党派负责人".to_string(),
        ];
        assert_eq!(result, expected);
    }

    #[test]
    fn test_with_code() {
        let chunks = vec!["1-01(GBM10100)中国共产党机关负责人".to_string()];
        let result = parse_categories(&chunks).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].id, "1-01");
        assert_eq!(result[0].code.as_deref(), Some("GBM10100"));
        assert_eq!(result[0].desc.as_deref(), Some("中国共产党机关负责人"));
    }

    #[test]
    fn test_without_code() {
        let chunks = vec!["1-03 民主党派负责人".to_string()];
        let result = parse_categories(&chunks).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].id, "1-03");
        assert!(result[0].code.is_none());
        assert_eq!(result[0].desc.as_deref(), Some("民主党派负责人"));
    }

    #[test]
    fn test_multiple_level_id() {
        let chunks = vec!["1-02-01-00 国家权力机关负责人".to_string()];
        let result = parse_categories(&chunks).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].id, "1-02-01-00");
        assert!(result[0].code.is_none());
        assert_eq!(result[0].desc.as_deref(), Some("国家权力机关负责人"));
    }

    #[test]
    fn test_only_id() {
        let chunks = vec!["1-04-00-00".to_string()];
        let result = parse_categories(&chunks).unwrap();
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].id, "1-04-00-00");
        assert!(result[0].code.is_none());
        // todo: check this
        // assert!(result[0].desc.is_none());
    }

    #[test]
    fn test_mixed_chunks() {
        let chunks = vec![
            "1-01(GBM10100)中国共产党机关负责人".to_string(),
            "1-02-01-00国家权力机关负责人".to_string(),
            "1-03民主党派负责人".to_string(),
        ];
        let result = parse_categories(&chunks).unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0].id, "1-01");
        assert_eq!(result[1].id, "1-02-01-00");
        assert_eq!(result[2].id, "1-03");
    }

    #[test]
    fn test_normalize_basic() {
        let input = "第一大类1(GBM10)党的机关、国家机关、群众团体和社会组织、企事业单位负责人";
        let expected = "1(GBM10) 第一大类党的机关、国家机关、群众团体和社会组织、企事业单位负责人";
        assert_eq!(normalize_first_category(input).unwrap(), expected);
    }

    #[test]
    fn test_normalize_basic_two() {
        let input = "第一大类 1(GBM10) 党的机关、国家机关、群众团体和社会组织、企事业单位负责人";
        let expected = "1(GBM10) 第一大类党的机关、国家机关、群众团体和社会组织、企事业单位负责人";
        assert_eq!(normalize_first_category(input).unwrap(), expected);
    }
}
