# Declaratest

**Declaratest** is a tool that generates DOCX test documents from structured markdown template files. It's designed to help educators and test creators quickly generate formatted test papers with various question types.

The project provides two implementations:
- **Python version**: Simple, lightweight implementation using `python-docx`
- **Rust version**: Feature-rich implementation with advanced template handling and styling

## Features

### Core Functionality
- Parse structured markdown files into formatted DOCX test documents
- Support for multiple question types: short answer, long answer, matching, fill-in-the-blanks, and oral questions
- Automatic formatting with professional styling
- Section-based organization with customizable layouts

### Question Types Supported

| Type | Description | Example |
|------|-------------|---------|
| `short` | Short answer questions | *What is the capital of France?* |
| `long` | Essay-style questions with optional line counts | *Describe the causes of World War I (5 lines)* |
| `matching_v` | Vertical matching (left column → right column) | *photosynthesis → Process by which plants make food* |
| `matching_h` | Horizontal matching layout | Same as vertical but laid out horizontally |
| `blanks` | Fill-in-the-blank questions | *The largest planet is ________.* |
| `oral` | Oral examination questions with sub-points | *Explain climate change* with bullet-point sub-topics |

### Rust-Specific Features
- **Template Creation**: Generate DOCX templates with built-in professional styles
- **Template Parsing**: Extract and reuse styling from existing DOCX files
- **Debug Mode**: Explore template API and styling information - shows detailed information about:
  - Document structure and element counts
  - All available styles (paragraph, character, table styles)
  - Font families available in the template
  - Template parsing success/failure details
- **Advanced Styling**: Sophisticated page layout, margins, and typography control

## Quick Start

Want to try it out immediately? Use the included sample file:

```bash
# Python version (if you have uv installed)
uv run declaratest.py sample.md

# Python version (traditional)
python3 declaratest.py sample.md  # requires: pip install python-docx

# Rust version
cargo run -- sample.md
```

This will generate `test.docx` with a complete example showcasing all question types.

## Installation

### Python Version

The Python version is a single script with embedded dependencies (PEP 723 format).

**Option 1: Using uv (recommended)**
```bash
# Install uv if not already installed
curl -LsSf https://astral.sh/uv/install.sh | sh

# Run directly - uv will handle dependencies automatically
uv run declaratest.py input.md
```

**Option 2: Traditional pip installation**
```bash
# Install dependencies manually
pip install python-docx

# Run the script
python3 declaratest.py input.md
```

### Rust Version
```bash
# Install Rust (if not already installed)
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh

# Build the project
cargo build --release

# The binary will be available at target/release/declaratest
```

## Usage

### Python Version
```bash
# Basic usage (with uv)
uv run declaratest.py input.md

# Basic usage (traditional)
python3 declaratest.py input.md

# Specify output file
python3 declaratest.py input.md -o my_test.docx

# Use a custom template
python3 declaratest.py input.md -t template.docx -o formatted_test.docx
```

**Note**: The Python version help text mentions `.txt` files, but it actually works with `.md` files as shown in the examples.

### Rust Version
```bash
# Basic usage
./declaratest input.md

# Specify output file
./declaratest input.md -o my_test.docx

# Use a custom template
./declaratest input.md -t template.docx -o formatted_test.docx

# Create a default template with built-in styles
./declaratest sample.md --create-template

# Create a minimal template for testing
./declaratest sample.md --create-minimal

# Debug template API (useful for development)
./declaratest sample.md --debug-template
```

## Input File Format

Declaratest uses a structured markdown format. Here's the basic structure:

```markdown
# Test
Subject: [Your Subject]
Title: [Test Title]

## Section: [Section Name]
Type: [question_type]
Separate Sheet: [yes/no]  # Optional
Subtitle: [section subtitle]  # Optional

- [Question 1]
- [Question 2]
- [Question 3]

## Section: [Another Section]
Type: [another_type]
- [More questions...]
```

### Example Input File

```markdown
# Test
Subject: Biology
Title: Cell Structure and Function Quiz

## Section: Quick Facts
Type: short
- What is the powerhouse of the cell?
- Name the process by which plants make food.
- How many chromosomes do humans have?

## Section: Essay Questions
Type: long
Separate Sheet: yes
- Explain the process of photosynthesis in detail (10 lines).
- Describe the differences between plant and animal cells (8 lines).

## Section: Match the Organelles
Type: matching_v
- Mitochondria -> Powerhouse of the cell
- Nucleus -> Control center
- Chloroplast -> Site of photosynthesis
- Ribosome -> Protein synthesis

## Section: Fill in the Blanks
Type: blanks
- The _______ is the basic unit of life.
- DNA is stored in the _______ of the cell.
- Plants use _______ to capture light energy.

## Section: Oral Assessment
Type: oral
- Explain the cell cycle
    - What happens during interphase?
    - Describe the stages of mitosis
    - What is the purpose of cell division?
- Discuss cellular respiration
    - Where does it occur?
    - What are the inputs and outputs?
```

## File Structure

### Python Implementation
- `declaratest.py` - Single-file implementation with all functionality

### Rust Implementation
- `src/main.rs` - CLI interface and main application logic
- `src/parser.rs` - Markdown template parsing
- `src/generator.rs` - DOCX document generation
- `src/template.rs` - Template handling and style management  
- `src/types.rs` - Data structures and type definitions
- `Cargo.toml` - Rust project configuration

## Dependencies

### Python Version
- `python-docx` - DOCX document creation and manipulation
- Python 3.11+ (recommended)

### Rust Version
- `docx-rs` - DOCX document creation
- `docx-rust` - DOCX template parsing (custom fork)
- `clap` - Command-line argument parsing
- `regex` - Text pattern matching
- `rand` - Random number generation for question shuffling
- `serde` - Serialization/deserialization

## Technical Overview

### How It Works

1. **Parsing**: The input markdown file is parsed to extract:
   - Test metadata (subject, title)
   - Section definitions with types and options
   - Questions organized by section type

2. **Processing**: Questions are processed based on their type:
   - Text questions can have line count specifications 
   - Matching questions are parsed into left/right pairs
   - Oral questions can have nested sub-points
   - Fill-in-the-blank questions are formatted for answer lines

3. **Generation**: The DOCX document is created with:
   - Professional header styling
   - Section-based organization
   - Appropriate formatting for each question type
   - Optional separate answer sheets for certain sections

### Architecture

**Python Implementation**: Single-file design with embedded dependencies using PEP 723 format
- `parse_template()` - Markdown parsing logic
- `generate_docx()` - Document generation with python-docx
- Inline type definitions using TypedDict

**Rust Implementation**: Modular design with separate concerns
- `parser.rs` - Template parsing with comprehensive error handling
- `generator.rs` - DOCX generation with docx-rs
- `template.rs` - Advanced template processing and style management
- `types.rs` - Shared data structures with serde serialization
- `main.rs` - CLI interface with clap

## Implementation Comparison

| Feature | Python | Rust |
|---------|--------|------|
| Basic document generation | ✅ | ✅ |
| Template file support | ✅ | ✅ |
| Template creation | ❌ | ✅ |
| Template debugging | ❌ | ✅ |
| Advanced styling | Basic | Advanced |
| Performance | Good | Excellent |
| Memory usage | Higher | Lower |
| Build time | None | ~1-2 minutes |
| Dependencies | Minimal | More complex |

## Contributing

Both implementations welcome contributions! The Python version is great for quick prototyping and simple additions, while the Rust version is better for performance improvements and advanced features.

### Development Setup
```bash
# For Python development
pip install python-docx

# For Rust development  
cargo build
cargo run -- sample.md

# Test the implementations
cargo run -- sample.md -o rust_output.docx
python3 declaratest.py sample.md -o python_output.docx
```

## Examples

See `sample.md` for a comprehensive example that demonstrates all supported question types and formatting options.

The generated DOCX files include:
- Professional header with subject and title
- Properly formatted sections with appropriate spacing
- Numbered questions with consistent styling
- Support for separate answer sheets
- Oral assessment sheets for oral question types