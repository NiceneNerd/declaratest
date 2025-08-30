use clap::Parser;
use anyhow::{Result, Context};
use regex::Regex;
use rand::seq::SliceRandom;
use std::fs;
use std::path::Path;
use docx_rs::*;

#[derive(Debug, Clone)]
pub enum Question {
    Matching { left: String, right: String },
    Text { text: String, lines: Option<u32> },
    Blank { text: String },
}

#[derive(Debug, Clone)]
pub enum SectionType {
    Short,
    Long,
    MatchingV,
    MatchingH,
    Blanks,
}

impl SectionType {
    fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "short" => Some(Self::Short),
            "long" => Some(Self::Long),
            "matching_v" => Some(Self::MatchingV),
            "matching_h" => Some(Self::MatchingH),
            "blanks" => Some(Self::Blanks),
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct Section {
    pub name: String,
    pub questions: Vec<Question>,
    pub section_type: Option<SectionType>,
    pub separate_sheet: bool,
}

#[derive(Debug, Clone)]
pub struct TestData {
    pub subject: String,
    pub title: String,
    pub sections: Vec<Section>,
}

#[derive(Parser)]
#[command(name = "declaratest")]
#[command(about = "Generate a DOCX test file from a template")]
struct Args {
    /// Path to the input template file (.txt or .md)
    input: String,
    
    /// Path to the output DOCX file
    #[arg(short, long, default_value = "test.docx")]
    output: String,
    
    /// Path to a DOCX template file (optional)
    #[arg(short, long)]
    template: Option<String>,
}

fn parse_template<P: AsRef<Path>>(file_path: P) -> Result<TestData> {
    let content = fs::read_to_string(file_path).context("Failed to read template file")?;
    let lines: Vec<&str> = content.lines().collect();
    
    let mut test_data = TestData {
        subject: String::new(),
        title: String::new(),
        sections: Vec::new(),
    };
    
    let mut current_section: Option<Section> = None;
    let mut sections = Vec::new();
    
    for line in lines {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        
        if line.starts_with("# Test") {
            continue;
        } else if line.starts_with("Subject:") {
            test_data.subject = line.split(':').nth(1).unwrap_or("").trim().to_string();
        } else if line.starts_with("Title:") {
            test_data.title = line.split(':').nth(1).unwrap_or("").trim().to_string();
        } else if line.starts_with("## Section:") {
            if let Some(section) = current_section.take() {
                sections.push(section);
            }
            let name = line.split(':').nth(1).unwrap_or("").trim().to_string();
            current_section = Some(Section {
                name,
                questions: Vec::new(),
                section_type: None,
                separate_sheet: false,
            });
        } else if line.starts_with("Type:") {
            if let Some(ref mut section) = current_section {
                let type_str = line.split(':').nth(1).unwrap_or("").trim();
                section.section_type = SectionType::from_str(type_str);
            }
        } else if line.starts_with("Separate Sheet:") {
            if let Some(ref mut section) = current_section {
                let value = line.split(':').nth(1).unwrap_or("").trim().to_lowercase();
                section.separate_sheet = value == "yes";
            }
        } else if line.starts_with("- ") {
            if let Some(ref mut section) = current_section {
                let q = &line[2..].trim();
                if let Some(question) = parse_question(q, &section.section_type) {
                    section.questions.push(question);
                }
            }
        }
    }
    
    if let Some(section) = current_section {
        sections.push(section);
    }
    
    test_data.sections = sections;
    Ok(test_data)
}

fn parse_question(q: &str, section_type: &Option<SectionType>) -> Option<Question> {
    match section_type {
        Some(SectionType::MatchingV) | Some(SectionType::MatchingH) => {
            if q.contains("->") {
                let parts: Vec<&str> = q.split("->").collect();
                if parts.len() == 2 {
                    return Some(Question::Matching {
                        left: parts[0].trim().to_string(),
                        right: parts[1].trim().to_string(),
                    });
                }
            }
        }
        Some(SectionType::Blanks) => {
            return Some(Question::Blank {
                text: q.to_string(),
            });
        }
        _ => {
            // Parse lines count from pattern like "(3 lines)"
            let lines_regex = Regex::new(r"\((\d+)\s+lines?\)").unwrap();
            let lines_count = if let Some(captures) = lines_regex.captures(q) {
                captures.get(1)?.as_str().parse().ok()
            } else {
                None
            };
            
            let text = lines_regex.replace(q, "").trim().to_string();
            
            return Some(Question::Text {
                text,
                lines: lines_count,
            });
        }
    }
    None
}

fn generate_docx(test_data: &TestData, output_path: &str, _template_path: Option<&str>) -> Result<()> {
    let mut docx = Docx::new();
    
    // Add subject and title
    docx = docx
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&format!("{} Test", test_data.subject)).style("Subtitle")))
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&test_data.title).style("Title")));
    
    // Add sections
    for section in &test_data.sections {
        docx = add_section(docx, section)?;
    }
    
    // Save the document
    let file = std::fs::File::create(output_path).context("Failed to create output file")?;
    docx.build().pack(file).context("Failed to write DOCX file")?;
    
    Ok(())
}

fn add_section(mut docx: Docx, section: &Section) -> Result<Docx> {
    // Add section header
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text(&section.name))
            .style("Heading2")
    );
    
    // Add separate sheet notice if needed
    if matches!(section.section_type, Some(SectionType::Long)) && section.separate_sheet {
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Use a separate sheet of paper").italic())
        );
    }
    
    // Add questions based on section type
    match section.section_type {
        Some(SectionType::Short) => {
            docx = add_short_questions(docx, section)?;
        }
        Some(SectionType::Long) => {
            docx = add_long_questions(docx, section)?;
        }
        Some(SectionType::MatchingV) => {
            docx = add_matching_v_questions(docx, section)?;
        }
        Some(SectionType::MatchingH) => {
            docx = add_matching_h_questions(docx, section)?;
        }
        Some(SectionType::Blanks) => {
            docx = add_blanks_questions(docx, section)?;
        }
        None => {
            // Default to short questions
            docx = add_short_questions(docx, section)?;
        }
    }
    
    Ok(docx)
}

fn add_short_questions(mut docx: Docx, section: &Section) -> Result<Docx> {
    for (i, question) in section.questions.iter().enumerate() {
        match question {
            Question::Text { text, lines } => {
                // Add numbered question
                docx = docx.add_paragraph(
                    Paragraph::new()
                        .add_run(Run::new().add_text(&format!("{}. {}", i + 1, text)))
                );
                
                // Add blank lines for answers
                let line_count = lines.unwrap_or(1);
                for _ in 0..line_count {
                    docx = docx.add_paragraph(
                        Paragraph::new()
                            .add_run(Run::new().add_text("_".repeat(50)))
                    );
                }
            }
            _ => {} // Skip non-text questions in short sections
        }
    }
    Ok(docx)
}

fn add_long_questions(mut docx: Docx, section: &Section) -> Result<Docx> {
    for (i, question) in section.questions.iter().enumerate() {
        match question {
            Question::Text { text, lines } => {
                // Add numbered question
                docx = docx.add_paragraph(
                    Paragraph::new()
                        .add_run(Run::new().add_text(&format!("{}. {}", i + 1, text)))
                );
                
                // Add blank lines for answers if not using separate sheet
                if !section.separate_sheet {
                    let line_count = lines.unwrap_or(10);
                    for _ in 0..line_count {
                        docx = docx.add_paragraph(
                            Paragraph::new()
                                .add_run(Run::new().add_text("_".repeat(50)))
                        );
                    }
                }
            }
            _ => {} // Skip non-text questions in long sections
        }
    }
    Ok(docx)
}

fn add_matching_v_questions(mut docx: Docx, section: &Section) -> Result<Docx> {
    let mut pairs: Vec<(String, String)> = Vec::new();
    
    for question in &section.questions {
        if let Question::Matching { left, right } = question {
            pairs.push((left.clone(), right.clone()));
        }
    }
    
    if pairs.is_empty() {
        return Ok(docx);
    }
    
    // Shuffle pairs and separate lefts and rights
    let mut rng = rand::thread_rng();
    pairs.shuffle(&mut rng);
    
    let lefts: Vec<String> = pairs.iter().map(|(l, _)| l.clone()).collect();
    let mut rights: Vec<String> = pairs.iter().map(|(_, r)| r.clone()).collect();
    rights.shuffle(&mut rng);
    
    // Create table with lefts and shuffled rights
    let mut table = Table::new(vec![]);
    
    for (i, (left, right)) in lefts.iter().zip(rights.iter()).enumerate() {
        let left_cell = TableCell::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_text(&format!("___ {}", left)))
        );
        let right_cell = TableCell::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_text(&format!("{}. {}", 
                char::from(b'A' + i as u8), right)))
        );
        
        table = table.add_row(TableRow::new(vec![left_cell, right_cell]));
    }
    
    docx = docx.add_table(table);
    Ok(docx)
}

fn add_matching_h_questions(mut docx: Docx, section: &Section) -> Result<Docx> {
    let mut terms: Vec<String> = Vec::new();
    let mut definitions: Vec<String> = Vec::new();
    
    for question in &section.questions {
        if let Question::Matching { left, right } = question {
            terms.push(left.clone());
            definitions.push(right.clone());
        }
    }
    
    if terms.is_empty() {
        return Ok(docx);
    }
    
    // Create term bank
    let rows = (terms.len() + 2) / 3; // 3 columns
    let mut term_table = Table::new(vec![]);
    
    for row in 0..rows {
        let mut cells = Vec::new();
        for col in 0..3 {
            let idx = row * 3 + col;
            if idx < terms.len() {
                let cell = TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text(&format!("{}. {}", 
                        char::from(b'A' + idx as u8), terms[idx])))
                );
                cells.push(cell);
            } else {
                cells.push(TableCell::new());
            }
        }
        term_table = term_table.add_row(TableRow::new(cells));
    }
    
    docx = docx.add_table(term_table);
    
    // Add matching definitions with blanks
    for (_i, definition) in definitions.iter().enumerate() {
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(&format!("_____ {}", definition)))
        );
    }
    
    Ok(docx)
}

fn add_blanks_questions(mut docx: Docx, section: &Section) -> Result<Docx> {
    for (i, question) in section.questions.iter().enumerate() {
        if let Question::Blank { text } = question {
            // Replace underscores with formatted blanks
            let parts: Vec<&str> = text.split('_').collect();
            let mut paragraph = Paragraph::new();
            
            paragraph = paragraph.add_run(Run::new().add_text(&format!("{}. ", i + 1)));
            
            for (j, part) in parts.iter().enumerate() {
                if j > 0 {
                    // Add blank space
                    paragraph = paragraph.add_run(Run::new().add_text("_______"));
                }
                if !part.is_empty() {
                    paragraph = paragraph.add_run(Run::new().add_text(*part));
                }
            }
            
            docx = docx.add_paragraph(paragraph);
        }
    }
    Ok(docx)
}

fn main() -> Result<()> {
    let args = Args::parse();
    
    let test_data = parse_template(&args.input)?;
    generate_docx(&test_data, &args.output, args.template.as_deref())?;
    
    println!("Generated DOCX file: {}", args.output);
    
    Ok(())
}
