use docx_rs::*;
use docx_rust::DocxFile;
use std::path::Path;
use std::collections::HashMap;

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
    // Parsed styles from template (style_id -> Style)
    pub parsed_styles: HashMap<String, docx_rs::Style>,
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
            // Empty styles map (defaults will be added later)
            parsed_styles: HashMap::new(),
        }
    }
}

// Convert docx_rust::Style to docx_rs::Style
fn convert_style(rust_style: &docx_rust::styles::Style) -> Option<docx_rs::Style> {
    // Get the style type
    let style_type = match rust_style.ty {
        Some(docx_rust::styles::StyleType::Paragraph) => StyleType::Paragraph,
        Some(docx_rust::styles::StyleType::Character) => StyleType::Character,
        Some(docx_rust::styles::StyleType::Table) => StyleType::Table,
        Some(docx_rust::styles::StyleType::Numbering) => StyleType::Numbering,
        _ => return None, // Skip unknown style types
    };
    
    // Create base style
    let style_id = rust_style.style_id.to_string();
    let mut rs_style = docx_rs::Style::new(&style_id, style_type);
    
    // Set name if available
    if let Some(ref name) = rust_style.name {
        rs_style = rs_style.name(name.value.to_string());
    }
    
    // Convert character (run) properties
    if let Some(ref char_props) = rust_style.character {
        // Size
        if let Some(ref sz) = char_props.size {
            let size = sz.value.abs() as usize; // Convert isize to usize
            rs_style = rs_style.size(size);
        }
        
        // Bold
        if char_props.bold.is_some() {
            rs_style = rs_style.bold();
        }
        
        // Italic
        if char_props.italics.is_some() {
            rs_style = rs_style.italic();
        }
        
        // Color
        if let Some(ref color) = char_props.color {
            if let Some(color_val) = color.value.as_ref() {
                rs_style = rs_style.color(color_val.to_string());
            }
        }
        
        // Font names - try to extract font information
        // Look for font-related properties in character properties
        // Note: docx_rust might store fonts differently, this is a best-effort extraction
        if let Some(ref fonts) = char_props.fonts {
            let mut run_fonts = RunFonts::new();
            
            // Try to extract ASCII font
            if let Some(ref ascii_font) = fonts.ascii_theme {
                run_fonts = run_fonts.ascii(ascii_font.to_string());
            } else if let Some(ref ascii_font) = fonts.ascii {
                run_fonts = run_fonts.ascii(ascii_font.to_string());
            }
            
            // Set the fonts on the style's run property
            rs_style.run_property = rs_style.run_property.fonts(run_fonts);
        }
    }
    
    // Convert paragraph properties
    if let Some(ref para_props) = rust_style.paragraph {
        // Alignment
        if let Some(ref jc) = para_props.justification {
            let alignment = match jc.value {
                docx_rust::formatting::JustificationVal::Center => AlignmentType::Center,
                docx_rust::formatting::JustificationVal::Right => AlignmentType::Right,
                docx_rust::formatting::JustificationVal::Left => AlignmentType::Left,
                docx_rust::formatting::JustificationVal::Both => AlignmentType::Justified,
                _ => AlignmentType::Left,
            };
            rs_style.paragraph_property = rs_style.paragraph_property.align(alignment);
        }
        
        // Paragraph spacing - extract before/after spacing and line spacing
        let mut line_spacing = LineSpacing::new();
        let mut has_spacing = false;
        
        // Try to extract line spacing
        if let Some(ref spacing) = para_props.spacing {
            // Extract before spacing (in twentieths of a point)
            if let Some(before) = spacing.before {
                line_spacing = line_spacing.before(before as u32);
                has_spacing = true;
            }
            
            // Extract after spacing (in twentieths of a point)
            if let Some(after) = spacing.after {
                line_spacing = line_spacing.after(after as u32);
                has_spacing = true;
            }
            
            // Extract line spacing rule and value
            if let Some(line_val) = spacing.line {
                line_spacing = line_spacing.line(line_val as i32);
                has_spacing = true;
            }
        }
        
        // Apply line spacing if we found any spacing information
        if has_spacing {
            rs_style.paragraph_property = rs_style.paragraph_property.line_spacing(line_spacing);
        }
    }
    
    Some(rs_style)
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
                    
                    // Convert and store styles
                    let mut converted_count = 0;
                    for style in &docx.styles.styles {
                        if let Some(converted_style) = convert_style(style) {
                            let style_id = style.style_id.to_string();
                            println!("✓ Converted style: {} ({})", style_id, 
                                   style.name.as_ref().map(|n| n.value.as_ref()).unwrap_or("unnamed"));
                            template_info.parsed_styles.insert(style_id, converted_style);
                            converted_count += 1;
                        }
                    }
                    
                    println!("✓ Successfully converted {} styles from template", converted_count);
                    
                    // Check for essential styles
                    let has_title = template_info.parsed_styles.contains_key("Title");
                    let has_subtitle = template_info.parsed_styles.contains_key("Subtitle");
                    let has_heading1 = template_info.parsed_styles.contains_key("Heading1") || 
                                     template_info.parsed_styles.contains_key("Heading 1");
                    let has_heading2 = template_info.parsed_styles.contains_key("Heading2") || 
                                     template_info.parsed_styles.contains_key("Heading 2");
                    let has_list_number = template_info.parsed_styles.contains_key("ListNumber") || 
                                        template_info.parsed_styles.contains_key("List Number");
                    
                    if has_title { println!("  - Title style found"); }
                    if has_subtitle { println!("  - Subtitle style found"); }
                    if has_heading1 || has_heading2 { println!("  - Heading styles found"); }
                    if has_list_number { println!("  - List Number style found"); }
                }
                Err(e) => {
                    println!("Warning: Could not parse document structure ({})", e);
                    println!("This is typically due to compatibility differences between docx_rust and docx_rs libraries.");
                    println!("Template file was loaded successfully - using defaults for missing styles.");
                    
                    // Even if parsing fails, we can still provide fallback behavior
                    // The template file was successfully loaded, so the --template argument is working
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
    
    // Apply parsed styles from template
    for (style_id, style) in &template_info.parsed_styles {
        println!("✓ Applying template style: {}", style_id);
        result = result.add_style(style.clone());
    }
    
    // Add essential styles if not found in template
    result = add_essential_styles(result, template_info);
    
    result
}

// Add essential styles to the document if they're missing from the template
fn add_essential_styles(mut docx: Docx, template_info: &TemplateInfo) -> Docx {
    // Check what styles are available from the template
    let has_title = template_info.parsed_styles.contains_key("Title");
    let has_subtitle = template_info.parsed_styles.contains_key("Subtitle");
    let has_heading1 = template_info.parsed_styles.contains_key("Heading1") || 
                      template_info.parsed_styles.contains_key("Heading 1");
    let has_heading2 = template_info.parsed_styles.contains_key("Heading2") || 
                      template_info.parsed_styles.contains_key("Heading 2");
    let has_list_number = template_info.parsed_styles.contains_key("ListNumber") || 
                         template_info.parsed_styles.contains_key("List Number");
    
    // Add Title style if not in template
    if !has_title {
        println!("✓ Adding default Title style");
        let mut title_style = Style::new("Title", StyleType::Paragraph)
            .name("Title")
            .size(48)
            .bold()
            .color("2F5496")
            .align(AlignmentType::Center);
        
        // Add proper spacing for title
        title_style.paragraph_property = title_style.paragraph_property.line_spacing(
            LineSpacing::new()
                .line(276)   // 1.15 line spacing (276/240 = 1.15)
                .before(240) // ~12pt before (one line worth)
                .after(60)   // ~3pt after (25% of line space)
        );
        
        docx = docx.add_style(title_style);
    }
    
    // Add Subtitle style if not in template
    if !has_subtitle {
        println!("✓ Adding default Subtitle style");
        let mut subtitle_style = Style::new("Subtitle", StyleType::Paragraph)
            .name("Subtitle")
            .size(28)
            .italic()
            .color("595959")
            .align(AlignmentType::Center);
        
        // Add proper spacing for subtitle  
        subtitle_style.paragraph_property = subtitle_style.paragraph_property.line_spacing(
            LineSpacing::new()
                .line(276)   // 1.15 line spacing
                .before(180) // ~9pt before  
                .after(60)   // ~3pt after
        );
        
        docx = docx.add_style(subtitle_style);
    }
    
    // Add Heading styles if not in template
    if !has_heading1 {
        println!("✓ Adding default Heading1 style");
        let mut heading1_style = Style::new("Heading1", StyleType::Paragraph)
            .name("Heading 1")
            .size(32)
            .bold()
            .color("2F5496");
        
        // Add proper spacing for heading1
        heading1_style.paragraph_property = heading1_style.paragraph_property.line_spacing(
            LineSpacing::new()
                .line(276)   // 1.15 line spacing
                .before(240) // ~12pt before (one line worth)
                .after(60)   // ~3pt after (25% of line space)
        );
        
        docx = docx.add_style(heading1_style);
    }
    
    if !has_heading2 {
        println!("✓ Adding default Heading2 style");
        let mut heading2_style = Style::new("Heading2", StyleType::Paragraph)
            .name("Heading 2")
            .size(26)
            .bold()
            .color("2F5496");
        
        // Add proper spacing for heading2
        heading2_style.paragraph_property = heading2_style.paragraph_property.line_spacing(
            LineSpacing::new()
                .line(276)   // 1.15 line spacing
                .before(200) // ~10pt before (slightly less than heading1)
                .after(40)   // ~2pt after (20% of line space)
        );
        
        docx = docx.add_style(heading2_style);
    }
    
    // Add List Number style if not in template
    if !has_list_number {
        println!("✓ Adding default ListNumber style");
        let mut list_style = Style::new("ListNumber", StyleType::Paragraph)
            .name("List Number")
            .size(24);
        
        // Add proper spacing for list number style
        list_style.paragraph_property = list_style.paragraph_property.line_spacing(
            LineSpacing::new()
                .line(276)   // 1.15 line spacing (default)
                .before(0)   // No space before questions
                .after(20)   // Small space after questions (~1pt)
        );
        
        docx = docx.add_style(list_style);
    }
    
    docx
}

// Line spacing utility
pub fn apply_line_spacing(para: Paragraph, spacing: f32) -> Paragraph {
    let spacing_val = (spacing * 240.0) as i32;
    para.line_spacing(LineSpacing::new().line(spacing_val))
}

// Apply paragraph spacing (before and after)
pub fn apply_paragraph_spacing(para: Paragraph, before_pt: f32, after_pt: f32) -> Paragraph {
    let before_twentieths = (before_pt * 20.0) as u32;
    let after_twentieths = (after_pt * 20.0) as u32;
    
    // Get current line spacing or use default
    let current_line_spacing = para.property.line_spacing.clone()
        .unwrap_or_else(|| LineSpacing::new().line(240)); // Default 1.0 line spacing
    
    para.line_spacing(
        current_line_spacing
            .before(before_twentieths)
            .after(after_twentieths)
    )
}

// Apply combined line and paragraph spacing
pub fn apply_combined_spacing(para: Paragraph, line_spacing: f32, before_pt: f32, after_pt: f32) -> Paragraph {
    let line_spacing_val = (line_spacing * 240.0) as i32;
    let before_twentieths = (before_pt * 20.0) as u32;
    let after_twentieths = (after_pt * 20.0) as u32;
    
    para.line_spacing(
        LineSpacing::new()
            .line(line_spacing_val)
            .before(before_twentieths)
            .after(after_twentieths)
    )
}