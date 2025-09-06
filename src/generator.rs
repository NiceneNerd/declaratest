use crate::template;
use crate::types::{Question, Section, SectionType, TestData};
use docx_rs::*;
use rand::{seq::SliceRandom, thread_rng};
use regex::Regex;
use std::path::Path;

pub fn generate_docx(
    test_data: &TestData,
    output_path: &Path,
    template_path: Option<&Path>,
) -> Result<(), Box<dyn std::error::Error>> {
    // Parse template if provided, otherwise use defaults
    let template_info = if let Some(template) = template_path {
        template::parse_template(template)?
    } else {
        template::TemplateInfo::default()
    };

    // Create document and apply template styling
    let mut docx = template::apply_template_info(Docx::new(), &template_info);

    // Add header
    let header = Header::new().add_paragraph(create_header_paragraph());
    docx = docx.header(header);

    // Add subject and title
    docx = add_subject_and_title(docx, test_data);

    // Add sections
    for section in &test_data.sections {
        docx = add_section(docx, section, test_data)?;
    }

    // Save document
    let file = std::fs::File::create(output_path)?;
    docx.build().pack(file)?;

    Ok(())
}

fn create_header_paragraph() -> Paragraph {
    Paragraph::new()
        .add_run(Run::new().add_text("\u{2000}Name:\u{2000}"))
        .add_run(
            Run::new()
                .add_text("\u{2000}".repeat(18))
                .underline("single"),
        )
        .add_run(Run::new().add_text("\u{2000}Date:\u{2000}"))
        .add_run(
            Run::new()
                .add_text("\u{2000}".repeat(12))
                .underline("single"),
        )
        .align(AlignmentType::Right)
}

fn add_subject_and_title(mut docx: Docx, test_data: &TestData) -> Docx {
    // Add subject as subtitle using the Subtitle style
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text(format!("{} Test", test_data.subject)))
            .style("Subtitle"),
    );

    // Add title using the Title style
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text(&test_data.title))
            .style("Title"),
    );

    docx
}

fn add_section(
    mut docx: Docx,
    section: &Section,
    test_data: &TestData,
) -> Result<Docx, Box<dyn std::error::Error>> {
    // Add section heading using Heading 2 style
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text(&section.name))
            .style("Heading2"),
    );

    // Add special notes for long and oral sections
    if matches!(section.section_type, Some(SectionType::Long)) && section.separate_sheet {
        docx = docx.add_paragraph(
            Paragraph::new().add_run(
                Run::new()
                    .add_text("Use a separate sheet of paper")
                    .italic()
                    .size(20), // 10pt
            ),
        );
    }

    if matches!(section.section_type, Some(SectionType::Oral)) {
        docx = docx.add_paragraph(
            Paragraph::new().add_run(
                Run::new()
                    .add_text("To be completed orally")
                    .italic()
                    .size(20), // 10pt
            ),
        );
    }

    // Add subtitle if present
    if let Some(subtitle) = &section.subtitle {
        docx = docx.add_paragraph(Paragraph::new().add_run(
            Run::new().add_text(subtitle).italic().size(20), // 10pt
        ));
    }

    // Add questions based on type
    match section.section_type {
        Some(SectionType::Short) => {
            docx = add_short_questions(docx, section);
        }
        Some(SectionType::Long) => {
            docx = add_long_questions(docx, section);
        }
        Some(SectionType::MatchingV) => {
            docx = add_matching_v(docx, section)?;
        }
        Some(SectionType::MatchingH) => {
            docx = add_matching_h(docx, section)?;
        }
        Some(SectionType::Blanks) => {
            docx = add_blanks_questions(docx, section);
        }
        Some(SectionType::Oral) => {
            docx = add_oral_questions(docx, section, test_data)?;
        }
        None => {
            // Default to short questions
            docx = add_short_questions(docx, section);
        }
    }

    Ok(docx)
}

fn add_short_questions(mut docx: Docx, section: &Section) -> Docx {
    for (index, question) in section.questions.iter().enumerate() {
        if let Question::Text(text_q) = question {
            // Add numbered question with markdown support and List Number style
            let mut question_para = Paragraph::new().style("ListNumber");
            question_para = question_para.add_run(Run::new().add_text(format!("{}. ", index + 1)));
            question_para = add_markdown_runs_to_paragraph(question_para, &text_q.text);
            question_para = template::apply_line_spacing(question_para, 1.0);
            docx = docx.add_paragraph(question_para);

            // Add blank lines with better underline sizing
            let num_lines = text_q.lines.unwrap_or(1);
            for _ in 0..num_lines {
                let blank_para = template::apply_line_spacing(
                    Paragraph::new().add_run(
                        Run::new()
                            .add_text("\t___________________________")
                            .underline("single"),
                    ), // Better sized underline
                    1.5,
                );
                docx = docx.add_paragraph(blank_para);
            }
        }
    }
    docx
}

fn add_long_questions(mut docx: Docx, section: &Section) -> Docx {
    for (index, question) in section.questions.iter().enumerate() {
        if let Question::Text(text_q) = question {
            // Add numbered question with markdown support and List Number style
            let mut question_para = Paragraph::new().style("ListNumber");
            question_para = question_para.add_run(Run::new().add_text(format!("{}. ", index + 1)));
            question_para = add_markdown_runs_to_paragraph(question_para, &text_q.text);
            question_para = template::apply_line_spacing(question_para, 1.0);
            docx = docx.add_paragraph(question_para);

            // Add blank lines if not separate sheet with better sizing
            if !section.separate_sheet {
                let num_lines = text_q.lines.unwrap_or(10);
                for _ in 0..num_lines {
                    let blank_para = template::apply_line_spacing(
                        Paragraph::new().add_run(
                            Run::new()
                                .add_text("\t___________________________")
                                .underline("single"),
                        ), // Better sized underline
                        1.5,
                    );
                    docx = docx.add_paragraph(blank_para);
                }
            }
        }
    }
    docx
}

fn add_matching_v(mut docx: Docx, section: &Section) -> Result<Docx, Box<dyn std::error::Error>> {
    let mut pairs = Vec::new();
    for question in &section.questions {
        if let Question::Matching(matching_q) = question {
            pairs.push((matching_q.left.clone(), matching_q.right.clone()));
        }
    }

    if pairs.is_empty() {
        return Ok(docx);
    }

    let mut rng = thread_rng();
    pairs.shuffle(&mut rng);

    let (lefts, rights): (Vec<_>, Vec<_>) = pairs.into_iter().unzip();
    let mut rights = rights;
    rights.shuffle(&mut rng);

    // Create table
    let mut table_rows = Vec::new();

    for i in 0..lefts.len() {
        let left_cell = TableCell::new().add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("_____").underline("single")) // Better sized blank
                .add_run(Run::new().add_text(format!(" {}", lefts[i]))),
        );

        let right_text = if i < rights.len() {
            format!("{}. {}", (b'A' + i as u8) as char, rights[i])
        } else {
            String::new()
        };

        let right_cell = TableCell::new()
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&right_text)));

        table_rows.push(TableRow::new(vec![left_cell, right_cell]));
    }

    let table = Table::new(table_rows);
    docx = docx.add_table(table);

    Ok(docx)
}

fn add_matching_h(mut docx: Docx, section: &Section) -> Result<Docx, Box<dyn std::error::Error>> {
    let mut terms = Vec::new();
    let mut defs = Vec::new();

    for question in &section.questions {
        if let Question::Matching(matching_q) = question {
            terms.push(matching_q.left.clone());
            defs.push(matching_q.right.clone());
        }
    }

    let n_terms = terms.len();
    if n_terms == 0 {
        return Ok(docx);
    }

    // Calculate optimal table layout
    let mut best_rows = 1;
    let mut best_empty = None;
    for r in 1..=3 {
        let cols = n_terms.div_ceil(r);
        let empty = r * cols - n_terms;
        if best_empty.is_none()
            || empty < best_empty.unwrap()
            || (empty == best_empty.unwrap() && r < best_rows)
        {
            best_rows = r;
            best_empty = Some(empty);
        }
    }

    let rows = best_rows;
    let cols = n_terms.div_ceil(rows);

    // Create term bank table
    let mut term_table_rows = Vec::new();
    let mut idx = 0;
    for _r in 0..rows {
        let mut row_cells = Vec::new();
        for _c in 0..cols {
            let cell = if idx < terms.len() {
                let label = format!("{}. {}", (b'A' + idx as u8) as char, terms[idx]);
                TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text(&label).size(20)), // 10pt
                )
            } else {
                TableCell::new().add_paragraph(Paragraph::new())
            };
            row_cells.push(cell);
            idx += 1;
        }
        term_table_rows.push(TableRow::new(row_cells));
    }

    let term_table = Table::new(term_table_rows);
    docx = docx.add_table(term_table);

    // Create matching table with better sized blanks
    let n_defs = defs.len();
    if n_defs > 0 {
        let match_rows = n_defs.div_ceil(2);
        let mut match_table_rows = Vec::new();

        for i in 0..match_rows {
            let mut row_cells = Vec::new();

            // Left pair
            if i * 2 < n_defs {
                let blank_cell = TableCell::new().add_paragraph(
                    Paragraph::new()
                        .add_run(Run::new().add_text("_______").underline("single")) // Better sized blank
                        .add_run(Run::new().add_text(" ")),
                );
                let def_cell = TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&defs[i * 2])));
                row_cells.push(blank_cell);
                row_cells.push(def_cell);
            } else {
                row_cells.push(TableCell::new().add_paragraph(Paragraph::new()));
                row_cells.push(TableCell::new().add_paragraph(Paragraph::new()));
            }

            // Right pair
            if i * 2 + 1 < n_defs {
                let blank_cell = TableCell::new().add_paragraph(
                    Paragraph::new()
                        .add_run(Run::new().add_text("_______").underline("single")) // Better sized blank
                        .add_run(Run::new().add_text(" ")),
                );
                let def_cell = TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(&defs[i * 2 + 1])));
                row_cells.push(blank_cell);
                row_cells.push(def_cell);
            } else {
                row_cells.push(TableCell::new().add_paragraph(Paragraph::new()));
                row_cells.push(TableCell::new().add_paragraph(Paragraph::new()));
            }

            match_table_rows.push(TableRow::new(row_cells));
        }

        let match_table = Table::new(match_table_rows);
        docx = docx.add_table(match_table);
    }

    Ok(docx)
}

fn add_blanks_questions(mut docx: Docx, section: &Section) -> Docx {
    for (index, question) in section.questions.iter().enumerate() {
        if let Question::Blank(blank_q) = question {
            let mut para = Paragraph::new().style("ListNumber"); // Use List Number style

            // Add question number
            para = para.add_run(Run::new().add_text(format!("{}. ", index + 1)));

            // Parse underscores in the text - improved blank sizes
            let parts: Vec<&str> = blank_q.text.split('_').collect();
            for (i, part) in parts.iter().enumerate() {
                if i > 0 {
                    // Add properly sized underlined blank for each underscore
                    para = para.add_run(Run::new().add_text("__________").underline("single"));
                    // Fixed-width blank
                }

                if !part.is_empty() {
                    para = add_markdown_runs_to_paragraph(para, part);
                }
            }

            // Apply double spacing as per Python version
            para = template::apply_line_spacing(para, 2.0);
            docx = docx.add_paragraph(para);
        }
    }
    docx
}

fn add_oral_questions(
    mut docx: Docx,
    section: &Section,
    test_data: &TestData,
) -> Result<Docx, Box<dyn std::error::Error>> {
    // Add questions to main document with markdown support and List Number style
    for (index, question) in section.questions.iter().enumerate() {
        if let Question::Oral(oral_q) = question {
            let mut question_para = Paragraph::new().style("ListNumber");
            question_para = question_para.add_run(Run::new().add_text(format!("{}. ", index + 1)));
            question_para = add_markdown_runs_to_paragraph(question_para, &oral_q.text);
            question_para = template::apply_line_spacing(question_para, 1.0);
            docx = docx.add_paragraph(question_para);
        }
    }

    // Generate oral assessment sheet on a separate page
    docx = add_oral_assessment_sheet(docx, section, test_data)?;

    Ok(docx)
}

fn add_oral_assessment_sheet(
    mut docx: Docx,
    section: &Section,
    test_data: &TestData,
) -> Result<Docx, Box<dyn std::error::Error>> {
    // Add page break
    docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));

    // Add header for assessment sheet using Heading 1 style
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text(format!("{} - Oral Assessment Sheet", test_data.title)))
            .style("Heading1")
            .align(AlignmentType::Center),
    );

    // Create table
    let mut table_rows = Vec::new();

    for question in &section.questions {
        if let Question::Oral(oral_q) = question {
            // Main question row
            let question_cell =
                TableCell::new().add_paragraph(create_markdown_paragraph(&oral_q.text));

            let score_cell = if oral_q.sub_points.is_empty() {
                TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text("_______").underline("single")), // Better sized score blank
                )
            } else {
                TableCell::new().add_paragraph(Paragraph::new())
            };

            table_rows.push(TableRow::new(vec![question_cell, score_cell]));

            // Sub-point rows
            for sub_point in &oral_q.sub_points {
                let sub_question_cell = TableCell::new().add_paragraph(
                    Paragraph::new()
                        .add_run(Run::new().add_text("\t"))
                        .add_run(Run::new().add_text(sub_point)),
                );

                let sub_score_cell = TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text("_______").underline("single")), // Better sized score blank
                );

                table_rows.push(TableRow::new(vec![sub_question_cell, sub_score_cell]));
            }
        }
    }

    // Add notes row
    let notes_cell = TableCell::new().add_paragraph(
        Paragraph::new().add_run(Run::new().add_text("Notes").bold().size(18)), // 9pt
    );

    // Add blank paragraphs for notes
    let mut notes_cell_with_blanks = notes_cell;
    for _ in 0..5 {
        notes_cell_with_blanks = notes_cell_with_blanks.add_paragraph(Paragraph::new());
    }

    // Create two separate cells since docx-rs doesn't support cell merging yet
    table_rows.push(TableRow::new(vec![
        notes_cell_with_blanks,
        TableCell::new().add_paragraph(Paragraph::new()),
    ]));

    let table = Table::new(table_rows);
    docx = docx.add_table(table);

    Ok(docx)
}

fn create_markdown_paragraph(text: &str) -> Paragraph {
    let mut para = Paragraph::new();
    para = add_markdown_runs_to_paragraph(para, text);
    para
}

fn add_markdown_runs_to_paragraph(mut para: Paragraph, text: &str) -> Paragraph {
    // Simple markdown parser for **bold**, *italic*, _italic_
    let pattern = Regex::new(r"(\*\*[^*]+\*\*|\*[^*]+\*|_[^_]+_)").unwrap();
    let mut pos = 0;

    for mat in pattern.find_iter(text) {
        let start = mat.start();
        let end = mat.end();

        // Add text before the match
        if start > pos {
            para = para.add_run(Run::new().add_text(&text[pos..start]));
        }

        let md = &text[start..end];
        let content = md.trim_matches(['*', '_']);

        let mut run = Run::new().add_text(content);

        if md.starts_with("**") && md.ends_with("**") {
            run = run.bold();
        } else if (md.starts_with('*') && md.ends_with('*'))
            || (md.starts_with('_') && md.ends_with('_'))
        {
            run = run.italic();
        }

        para = para.add_run(run);
        pos = end;
    }

    // Add remaining text
    if pos < text.len() {
        para = para.add_run(Run::new().add_text(&text[pos..]));
    }

    para
}
