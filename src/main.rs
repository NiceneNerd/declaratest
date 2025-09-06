use clap::Parser;
use std::path::PathBuf;

mod types;
mod parser;
mod generator;
mod template;

#[derive(Parser)]
#[command(name = "declaratest")]
#[command(about = "Generate a DOCX test file from a template")]
struct Args {
    /// Path to the input template file (.md)
    input: PathBuf,
    
    /// Path to the output DOCX file
    #[arg(short, long, default_value = "test.docx")]
    output: PathBuf,
    
    /// Path to a DOCX template file (optional)
    #[arg(short, long)]
    template: Option<PathBuf>,
    
    /// Debug mode - explore template API
    #[arg(long)]
    debug_template: bool,
    
    /// Create a simple default template with built-in styles
    #[arg(long)]
    create_template: bool,
    
    /// Create a minimal template for testing
    #[arg(long)]
    create_minimal: bool,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();
    
    if args.debug_template {
        debug_template_api(args.template.as_deref());
        return Ok(());
    }
    
    if args.create_template {
        create_simple_template()?;
        return Ok(());
    }
    
    if args.create_minimal {
        create_minimal_template()?;
        return Ok(());
    }
    
    let test_data = parser::parse_template(&args.input)?;
    generator::generate_docx(&test_data, &args.output, args.template.as_deref())?;
    println!("Generated DOCX file: {}", args.output.display());
    
    Ok(())
}

fn debug_template_api(template_path: Option<&std::path::Path>) {
    use docx_rust::DocxFile;
    use docx_rs::*;
    
    println!("=== Debugging Template API ===");
    
    let template_file = template_path.unwrap_or(std::path::Path::new("default.docx"));
    
    match DocxFile::from_file(template_file) {
        Ok(docx_file) => {
            println!("✓ Successfully loaded DOCX file: {}", template_file.display());
            
            // Parse the file to get structured data
            match docx_file.parse() {
                Ok(docx) => {
                    println!("✓ Successfully parsed DOCX content");
                    
                    // Explore the parsed document
                    println!("\nDocument structure:");
                    println!("- Document body has {} elements", docx.document.body.content.len());
                    
                    // Check styles
                    println!("\nStyles found ({} total):", docx.styles.styles.len());
                    for style in &docx.styles.styles {
                        println!("- ID: {:?}, Name: {:?}, Type: {:?}", 
                                 style.style_id, style.name, style.ty);
                    }
                    
                    // Check core properties - let's see what fields are actually available
                    if let Some(_core) = &docx.core {
                        println!("\nCore properties found (fields available but private)");
                    }
                    
                    // Check settings if available 
                    if let Some(_settings) = &docx.settings {
                        println!("\nSettings found");
                    }
                    
                    // Check font table
                    if let Some(font_table) = &docx.font_table {
                        println!("\nFont table found with {} fonts", font_table.fonts.len());
                        for font in &font_table.fonts {
                            println!("- Font: {:?}", font.name);
                        }
                    }
                    
                }
                Err(e) => {
                    println!("✗ Failed to parse DOCX content: {}", e);
                }
            }
            
        }
        Err(e) => {
            println!("✗ Failed to load DOCX: {}", e);
        }
    }
    
    println!("\n=== Testing docx-rs style API ===");
    
    // Test creating styles with docx-rs
    let style = Style::new("TestStyle", StyleType::Paragraph)
        .name("Test Style");
    
    println!("✓ Created docx-rs style with ID: {:?}", style.style_id);
}

fn create_simple_template() -> Result<(), Box<dyn std::error::Error>> {
    use docx_rs::*;
    use std::fs::File;
    
    println!("Creating simple template with built-in styles...");
    
    // Create a simple DOCX with proper built-in styles
    let docx = Docx::new()
        // Add Title style 
        .add_style(
            Style::new("Title", StyleType::Paragraph)
                .name("Title")
                .size(48)
                .bold()
                .color("2F5496")
                .align(AlignmentType::Center)
        )
        // Add Subtitle style
        .add_style(
            Style::new("Subtitle", StyleType::Paragraph)
                .name("Subtitle")
                .size(28)
                .italic()
                .color("595959")
                .align(AlignmentType::Center)
        )
        // Add Heading 1 style
        .add_style(
            Style::new("Heading1", StyleType::Paragraph)
                .name("Heading 1")
                .size(32)
                .bold()
                .color("2F5496")
        )
        // Add Heading 2 style
        .add_style(
            Style::new("Heading2", StyleType::Paragraph)
                .name("Heading 2")
                .size(26)
                .bold()
                .color("2F5496")
        )
        // Add ListNumber style
        .add_style(
            Style::new("ListNumber", StyleType::Paragraph)
                .name("List Number")
                .size(24)
        )
        // Set page layout - 8.5 x 11 inches with 1 inch margins
        .page_size(12240, 15840) // twentieths of a point
        .page_margin(PageMargin::new()
            .top(1440)   // 1 inch in twentieths of a point
            .bottom(1440)
            .left(1440)
            .right(1440)
        );
    
    // Save to file
    let file = File::create("simple_default.docx")?;
    docx.build().pack(file)?;
    
    println!("✓ Created simple_default.docx with built-in styles");
    
    Ok(())
}

fn create_minimal_template() -> Result<(), Box<dyn std::error::Error>> {
    use docx_rs::*;
    use std::fs::File;
    
    println!("Creating minimal template for testing...");
    
    // Create a very simple DOCX with basic styles that should be parseable by docx_rust
    let docx = Docx::new()
        .add_style(
            Style::new("Title", StyleType::Paragraph)
                .name("Title")
                .size(48)
                .bold()
                .color("2F5496")
        )
        .add_style(
            Style::new("Subtitle", StyleType::Paragraph)
                .name("Subtitle")
                .size(28)
                .italic()
                .color("595959")
        )
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Sample Title"))
                .style("Title")
        )
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Sample Subtitle"))
                .style("Subtitle")
        );
    
    let file = File::create("minimal_template.docx")?;
    docx.build().pack(file)?;
    println!("✓ Created minimal_template.docx");
    
    Ok(())
}
