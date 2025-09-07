use crate::types::{
    BlankQuestion, MatchingQuestion, OralQuestion, Question, Section, SectionType, TestData,
    TextQuestion,
};
use regex::Regex;
use std::fs;
use std::path::Path;

pub fn parse_template(file_path: &Path) -> Result<TestData, Box<dyn std::error::Error>> {
    let content = fs::read_to_string(file_path)?;
    let lines: Vec<&str> = content.lines().collect();

    let mut test = TestData {
        subject: String::new(),
        title: String::new(),
        sections: Vec::new(),
    };

    let mut sections = Vec::new();
    let mut current_section: Option<Section> = None;

    for line in lines {
        let line = line.trim_end(); // Only strip trailing whitespace
        if line.trim().is_empty() {
            continue;
        }

        if line.starts_with("# Test") {
            continue;
        } else if line.starts_with("Subject:") {
            test.subject = line.split(':').nth(1).unwrap_or("").trim().to_string();
        } else if line.starts_with("Title:") {
            test.title = line.split(':').nth(1).unwrap_or("").trim().to_string();
        } else if line.starts_with("## Section:") {
            let name = line.split(':').nth(1).unwrap_or("").trim().to_string();
            if let Some(section) = current_section.take() {
                sections.push(section);
            }
            current_section = Some(Section {
                name,
                questions: Vec::new(),
                section_type: None,
                separate_sheet: false,
                subtitle: None,
            });
        } else if line.starts_with("Type:") {
            if let Some(ref mut section) = current_section {
                let type_str = line.split(':').nth(1).unwrap_or("").trim();
                section.section_type = match type_str {
                    "short" => Some(SectionType::Short),
                    "long" => Some(SectionType::Long),
                    "matching_v" => Some(SectionType::MatchingV),
                    "matching_h" => Some(SectionType::MatchingH),
                    "blanks" => Some(SectionType::Blanks),
                    "oral" => Some(SectionType::Oral),
                    _ => None,
                };
            }
        } else if line.starts_with("Subtitle:") {
            if let Some(ref mut section) = current_section {
                section.subtitle = Some(line.split(':').nth(1).unwrap_or("").trim().to_string());
            }
        } else if line.starts_with("Oral:") {
            // Legacy support - convert "Oral: yes" to Type: oral
            if let Some(ref mut section) = current_section {
                let is_oral = line.split(':').nth(1).unwrap_or("").trim().to_lowercase() == "yes";
                if is_oral {
                    section.section_type = Some(SectionType::Oral);
                }
            }
        } else if line.starts_with("Separate Sheet:") {
            if let Some(ref mut section) = current_section {
                section.separate_sheet =
                    line.split(':').nth(1).unwrap_or("").trim().to_lowercase() == "yes";
            }
        } else if line.starts_with("    - ") || line.starts_with("\t- ") {
            // Handle sub-points for oral questions (indented with 4 spaces or tab)
            if let Some(ref mut section) = current_section {
                if matches!(section.section_type, Some(SectionType::Oral))
                    && !section.questions.is_empty()
                {
                    let sub_point = line
                        .trim()
                        .strip_prefix("- ")
                        .unwrap_or("")
                        .trim()
                        .to_string();
                    if let Some(Question::Oral(ref mut oral_q)) = section.questions.last_mut() {
                        oral_q.sub_points.push(sub_point);
                    }
                }
            }
        } else if line.starts_with("- ") {
            if let Some(ref mut section) = current_section {
                let q_text = line.strip_prefix("- ").unwrap_or("").trim();
                if let Some(question) = parse_question(q_text, section) {
                    section.questions.push(question);
                }
            }
        }
    }

    if let Some(section) = current_section {
        sections.push(section);
    }

    test.sections = sections;
    Ok(test)
}

fn parse_question(q: &str, current_section: &Section) -> Option<Question> {
    match current_section.section_type {
        Some(SectionType::MatchingV) | Some(SectionType::MatchingH) => {
            if q.contains(" -> ") {
                let parts: Vec<&str> = q.split(" -> ").collect();
                if parts.len() == 2 {
                    return Some(Question::Matching(MatchingQuestion {
                        left: parts[0].trim().to_string(),
                        right: parts[1].trim().to_string(),
                    }));
                }
            }
        }
        Some(SectionType::Blanks) => {
            return Some(Question::Blank(BlankQuestion {
                text: q.to_string(),
            }));
        }
        Some(SectionType::Oral) => {
            return Some(Question::Oral(OralQuestion {
                text: q.to_string(),
                sub_points: Vec::new(),
            }));
        }
        _ => {
            // For short and long questions
            let lines_regex = Regex::new(r"\((\d+)\s+lines?\)").unwrap();
            let lines_count = if let Some(captures) = lines_regex.captures(q) {
                captures
                    .get(1)
                    .and_then(|m| m.as_str().parse::<usize>().ok())
            } else {
                None
            };

            let text = lines_regex.replace(q, "").trim().to_string();

            return Some(Question::Text(TextQuestion {
                text,
                lines: lines_count,
            }));
        }
    }
    None
}
