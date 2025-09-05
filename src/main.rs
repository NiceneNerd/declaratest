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
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();
    
    let test_data = parser::parse_template(&args.input)?;
    generator::generate_docx(&test_data, &args.output, args.template.as_deref())?;
    println!("Generated DOCX file: {}", args.output.display());
    
    Ok(())
}
