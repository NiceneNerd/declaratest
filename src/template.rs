use docx_rs::*;
use docx_rust::DocxFile;
use std::path::Path;

// Template information extracted from DOCX files
#[derive(Debug, Clone)]
pub struct TemplateInfo {
    pub page_width: Option<u32>,
    pub page_height: Option<u32>,
    pub margin_top: Option<i32>,
    pub margin_bottom: Option<i32>,
    pub margin_left: Option<i32>,
    pub margin_right: Option<i32>,
    pub header_margin: Option<i32>,
    pub footer_margin: Option<i32>,
    // Style information
    pub has_title_style: bool,
    pub has_subtitle_style: bool,
    pub has_heading_styles: bool,
    pub has_list_number_style: bool,
}

impl Default for TemplateInfo {
    fn default() -> Self {
        Self {
            // Default: 8.5 x 11 inches in twentieths of a point
            page_width: Some(12240),
            page_height: Some(15840),
            // Default: 1 inch margins
            margin_top: Some(1440),
            margin_bottom: Some(1440),
            margin_left: Some(1440),
            margin_right: Some(1440),
            header_margin: Some(720),
            footer_margin: Some(720),
            // Default styles
            has_title_style: false,
            has_subtitle_style: false,
            has_heading_styles: false,
            has_list_number_style: false,
        }
    }
}

// Parse template file using docx-rust and extract styling information
pub fn parse_template(template_path: &Path) -> Result<TemplateInfo, Box<dyn std::error::Error>> {
    println!("Parsing template file: {}", template_path.display());
    
    match DocxFile::from_file(template_path) {
        Ok(docx_file) => {
            println!("✓ Successfully parsed template file");
            
            // Try to parse the document to extract style information
            let mut template_info = TemplateInfo::default();
            
            // Try to parse and extract style information
            match docx_file.parse() {
                Ok(docx) => {
                    println!("✓ Successfully parsed document structure");
                    
                    // Check for available styles
                    for style in &docx.styles.styles {
                        let style_id = &style.style_id;
                        match style_id.as_ref() {
                            "Title" => template_info.has_title_style = true,
                            "Subtitle" => template_info.has_subtitle_style = true,
                            "Heading1" | "Heading 1" => template_info.has_heading_styles = true,
                            "Heading2" | "Heading 2" => template_info.has_heading_styles = true,
                            "ListNumber" | "List Number" => template_info.has_list_number_style = true,
                            _ => {}
                        }
                    }
                    
                    // Log found styles
                    let style_count = docx.styles.styles.len();
                    println!("✓ Found {} styles in template", style_count);
                    if template_info.has_title_style { println!("  - Title style found"); }
                    if template_info.has_subtitle_style { println!("  - Subtitle style found"); }
                    if template_info.has_heading_styles { println!("  - Heading styles found"); }
                    if template_info.has_list_number_style { println!("  - List Number style found"); }
                }
                Err(e) => {
                    println!("Warning: Could not parse document structure ({}), using defaults", e);
                    println!("✓ Template file processed - using basic template information");
                }
            }
            
            Ok(template_info)
        }
        Err(e) => {
            eprintln!("Warning: Could not parse template file: {}", e);
            eprintln!("Using default styling instead.");
            Ok(TemplateInfo::default())
        }
    }
}

// Apply template information to a docx-rs document
pub fn apply_template_info(docx: Docx, template_info: &TemplateInfo) -> Docx {
    let mut result = docx;
    
    // Apply page size if available
    if let (Some(width), Some(height)) = (template_info.page_width, template_info.page_height) {
        result = result.page_size(width, height);
    }
    
    // Apply margins if available
    let margin = PageMargin::new()
        .top(template_info.margin_top.unwrap_or(1440))
        .bottom(template_info.margin_bottom.unwrap_or(1440))
        .left(template_info.margin_left.unwrap_or(1440))
        .right(template_info.margin_right.unwrap_or(1440))
        .header(template_info.header_margin.unwrap_or(720))
        .footer(template_info.footer_margin.unwrap_or(720));
    
    result = result.page_margin(margin);
    
    // Add essential styles if not found in template
    result = add_essential_styles(result, template_info);
    
    result
}

// Add essential styles to the document if they're missing from the template
fn add_essential_styles(mut docx: Docx, template_info: &TemplateInfo) -> Docx {
    // Add Title style if not in template
    if !template_info.has_title_style {
        docx = docx.add_style(
            Style::new("Title", StyleType::Paragraph)
                .name("Title")
                .size(48)
                .bold()
                .color("2F5496")
                .align(AlignmentType::Center)
        );
    }
    
    // Add Subtitle style if not in template
    if !template_info.has_subtitle_style {
        docx = docx.add_style(
            Style::new("Subtitle", StyleType::Paragraph)
                .name("Subtitle")
                .size(28)
                .italic()
                .color("595959")
                .align(AlignmentType::Center)
        );
    }
    
    // Add Heading styles if not in template
    if !template_info.has_heading_styles {
        docx = docx.add_style(
            Style::new("Heading1", StyleType::Paragraph)
                .name("Heading 1")
                .size(32)
                .bold()
                .color("2F5496")
        );
        
        docx = docx.add_style(
            Style::new("Heading2", StyleType::Paragraph)
                .name("Heading 2")
                .size(26)
                .bold()
                .color("2F5496")
        );
    }
    
    // Add List Number style if not in template
    if !template_info.has_list_number_style {
        docx = docx.add_style(
            Style::new("ListNumber", StyleType::Paragraph)
                .name("List Number")
                .size(24)
        );
    }
    
    docx
}

// Line spacing utility (unchanged)
pub fn apply_line_spacing(para: Paragraph, spacing: f32) -> Paragraph {
    let spacing_val = (spacing * 240.0) as i32;
    para.line_spacing(LineSpacing::new().line(spacing_val))
}