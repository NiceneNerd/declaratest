# Declaratest

A command-line tool for generating DOCX test documents from Markdown-like templates. This tool supports various question types including short answer, long answer, matching questions, and fill-in-the-blank questions.

## Features

- **Multiple Question Types**: Short answer, long answer, matching (vertical/horizontal), and fill-in-the-blank questions
- **Markdown Support**: Basic markdown formatting with **bold** and *italic* text
- **Flexible Layout**: Automatic page layout with proper margins and headers
- **Randomization**: Matching questions are automatically shuffled for test security
- **Cross-platform**: Available in both Python and Rust implementations

## Installation

### Rust Version (Recommended)

1. Make sure you have [Rust installed](https://rustup.rs/)
2. Clone this repository
3. Build the project:
   ```bash
   cargo build --release
   ```
4. Run the tool:
   ```bash
   cargo run -- input.md -o output.docx
   ```

### Python Version

1. Install Python dependencies:
   ```bash
   pip install python-docx
   ```
2. Run the Python script:
   ```bash
   python declaratest.py input.md -o output.docx
   ```

## Usage

```bash
declaratest [OPTIONS] <INPUT>

Arguments:
  <INPUT>  Path to the input template file (.txt or .md)

Options:
  -o, --output <OUTPUT>      Path to the output DOCX file [default: test.docx]
  -t, --template <TEMPLATE>  Path to a DOCX template file (optional)
  -h, --help                 Print help
```

## Template Format

The template file uses a simple Markdown-like format:

```markdown
# Test Title
Subject: Mathematics
Title: Algebra Fundamentals

## Section: Short Answer Questions
Type: short
- What is the value of x in the equation 2x + 5 = 15?
- Simplify the expression 3(x + 4) - 2x.

## Section: Essay Questions  
Type: long
Separate Sheet: yes
- Explain the difference between linear and quadratic equations. (5 lines)
- Describe how to solve a system of equations using substitution.

## Section: Matching
Type: matching_v
- Linear equation -> An equation with variables raised to the first power
- Quadratic equation -> An equation containing a variable raised to the second power
- Coefficient -> A numerical factor in a term

## Section: Fill in the Blanks
Type: blanks
- A _______ equation has variables raised to the first power only.
- The _______ of a quadratic equation is the highest point on a parabola opening downward.
```

### Section Types

- **short**: Short answer questions with customizable answer lines
- **long**: Essay-style questions with optional separate answer sheets
- **matching_v**: Vertical matching questions (left column matched to right column)
- **matching_h**: Horizontal matching with term bank above definitions
- **blanks**: Fill-in-the-blank questions with underlined spaces

### Question Options

- **Line count**: Add `(N lines)` to specify answer space for text questions
- **Separate sheet**: Use `Separate Sheet: yes` to indicate external answer sheets
- **Markdown formatting**: Use `**bold**` and `*italic*` or `_italic_` text formatting

## Examples

See `sample.md` for a complete example template that demonstrates all question types and formatting options.

## Technical Details

### Rust Implementation

- Built with `docx-rs` for DOCX generation
- Uses `clap` for command-line interface
- Implements regex-based markdown parsing
- Randomizes matching questions for test security

### Python Implementation

- Uses `python-docx` for document generation
- Supports advanced page layout and formatting
- Includes sophisticated table generation for matching questions

## Contributing

Contributions are welcome! Please feel free to submit issues and enhancement requests.

## License

This project is open source and available under the MIT License.