# Declaratest Windows GUI - Technical Specifications

## Executive Summary

This document provides comprehensive specifications for creating a Windows-specific GUI version of Declaratest with live Markdown editing and DOCX preview capabilities. The application will enable educators to create formatted test documents with real-time visual feedback.

## Project Overview

### Current Implementation Analysis

Declaratest currently exists in two command-line implementations:

1. **Python Implementation (`declaratest.py`)**
   - Uses `python-docx` library for DOCX generation
   - Comprehensive feature set including:
     - Multiple question types (short, long, matching, blanks, oral)
     - Markdown formatting support in questions (**bold**, *italic*)
     - Template-based DOCX output
     - Complex table layouts for matching questions
     - Oral assessment sheet generation
   - ~717 lines of well-structured Python code

2. **Rust Implementation (`src/`)**
   - Uses `docx-rs` for DOCX creation
   - Uses `docx-rust` for template parsing
   - Modular architecture with separate files:
     - `types.rs`: Data structures
     - `parser.rs`: Markdown template parser
     - `generator.rs`: DOCX generator
     - `template.rs`: Template handling
   - Similar feature parity with Python version

### Template Format

Both implementations parse a simple Markdown-like format:

```markdown
# Test
Subject: General Knowledge
Title: The Wonders of Everyday Life

## Section: Quick Facts
Type: short
- What is the most common element in the Earth's atmosphere?
- Name a fruit that is both red and green when ripe.

## Section: Creative Writing
Type: long
Separate Sheet: yes
- Write a short story about a lost sock.

## Section: Match the Term
Type: matching_v
- photosynthesis -> Process by which plants make food
- alloy -> Mixture of metals

## Section: Fill in the Blanks
Type: blanks
- The largest planet in our solar system is ___________.

## Section: Oral Questions
Type: oral
- What are the main causes of climate change?
    - What happens during evaporation?
    - How do clouds form?
```

Supported section types:
- `short`: Short answer questions with blank lines
- `long`: Essay questions (optionally on separate sheet)
- `matching_v`: Vertical matching (left-right table)
- `matching_h`: Horizontal matching (term bank + fill-in)
- `blanks`: Fill-in-the-blank questions
- `oral`: Oral examination questions with sub-points

## .NET GUI Application Specifications

### 1. Technology Stack Recommendations

#### Core Framework
- **.NET 8.0** (or latest LTS version)
  - Modern, performant, and well-supported
  - Excellent cross-platform tooling (though targeting Windows)
  - Rich ecosystem of libraries

#### UI Framework

**WPF (Windows Presentation Foundation)**

WPF is the chosen UI framework for this application, providing:

- **Native Windows Experience:** 
  - Native Windows look and feel
  - Full integration with Windows desktop environment
  - Proper support for Windows theming and accessibility features

- **Mature & Stable:**
  - Proven technology with 15+ years of production use
  - Extensive documentation and large community
  - Vast ecosystem of third-party controls and libraries

- **Rich Development Features:**
  - XAML-based declarative UI design
  - Powerful data binding capabilities with INotifyPropertyChanged
  - Built-in support for complex layouts including split views
  - Native support for Office document embedding via ActiveX

- **Performance:**
  - Excellent performance for document editing scenarios
  - Hardware-accelerated rendering
  - Efficient memory management for large documents

- **Office Integration:**
  - Seamless COM Interop support for Office applications
  - Native hosting of Word document controls
  - Direct access to Word's rendering engine

WPF is the perfect fit for this Windows-specific application, especially given the requirement for Office Interop integration.

#### Markdown Editor Component

**AvalonEdit** is the chosen text editor component:

- Open-source WPF text editor with rich features
- Comprehensive syntax highlighting support
- Line numbering and code folding
- Built-in undo/redo functionality
- Search and replace capabilities
- Excellent performance with large documents
- Native WPF integration
- Extensible architecture for custom features

#### DOCX Generation Library

**Open XML SDK (Microsoft)** is the chosen library for DOCX generation:

- **NuGet Package:** `DocumentFormat.OpenXml`
- Official Microsoft library for Office Open XML
- Complete OpenXML specification support
- Well-maintained and actively developed
- Extensive documentation with examples
- Full programmatic control over DOCX structure
- Used successfully in the existing Python and Rust implementations (via compatible libraries)
- Perfect compatibility with the template-based approach

#### DOCX Preview System

**Microsoft Office Interop** is the chosen approach for live DOCX preview:

**Implementation Approach:**
- Use `Microsoft.Office.Interop.Word` to display DOCX documents
- Embed Word control directly in the WPF application using WindowsFormsHost
- Load generated DOCX files directly into the embedded Word instance
- Leverage Word's native rendering engine for pixel-perfect preview

**Key Advantages:**

1. **Exact Rendering:**
   - 100% accurate preview - what you see is exactly what you get
   - No approximation or conversion artifacts
   - All Word features render correctly (tables, formatting, styles)
   - Complex layouts (matching sections, oral assessment sheets) display perfectly

2. **No Conversion Layer:**
   - Direct DOCX display without HTML/PDF conversion
   - Eliminates potential conversion bugs
   - Faster preview updates (no conversion overhead)
   - Simplified architecture

3. **Feature Completeness:**
   - All DOCX features supported automatically
   - Handles complex tables and nested structures
   - Proper style rendering from templates
   - Native support for headers, footers, page breaks

4. **Professional Output:**
   - Preview exactly matches printed/distributed tests
   - Teachers see precisely how tests will appear to students
   - Reduces preview-to-final discrepancies

**Implementation Details:**

```csharp
// Preview control in XAML
<WindowsFormsHost>
    <wf:Integration xmlns:wf="clr-namespace:System.Windows.Forms;assembly=System.Windows.Forms">
        <!-- Word document control hosted here -->
    </wf:Integration>
</WindowsFormsHost>
```

**Technical Considerations:**

1. **Prerequisites:**
   - Requires Microsoft Word installed on the user's machine
   - Typically not an issue for educational institutions
   - Word is standard software in most schools and universities

2. **COM Interop:**
   - Well-documented COM Interop approach
   - Mature technology with extensive examples
   - Proper object lifetime management required
   - Release COM objects explicitly to prevent memory leaks

3. **Performance:**
   - Initial Word instance creation: ~1-2 seconds
   - Document loading: < 500ms for typical tests
   - Reuse Word instance between updates for better performance
   - Background thread loading for UI responsiveness

4. **Licensing:**
   - Uses existing Word license on user's machine
   - No additional licensing costs
   - Standard Microsoft Office EULA applies

**Preview Update Strategy:**

1. User types in Markdown editor
2. Changes debounced (500ms delay)
3. Markdown parsed to TestData model
4. DOCX generated to temp file using Open XML SDK
5. Temp DOCX loaded into Word control
6. Word renders preview instantly
7. Temp file cleaned up after preview update

**Error Handling:**

- Graceful fallback if Word not installed (show message)
- Detection of Word installation during startup
- Clear error messages for COM Interop issues
- Option to export to DOCX without preview if Word unavailable

This approach provides the highest fidelity preview possible, ensuring educators see exactly how their tests will appear when printed or distributed.

### 2. Application Architecture

#### MVVM Pattern (Model-View-ViewModel)

```
Declaratest.GUI/
├── Models/
│   ├── TestData.cs
│   ├── Section.cs
│   ├── Question.cs (base class)
│   ├── TextQuestion.cs
│   ├── MatchingQuestion.cs
│   ├── BlankQuestion.cs
│   └── OralQuestion.cs
├── ViewModels/
│   ├── MainViewModel.cs
│   ├── EditorViewModel.cs
│   ├── PreviewViewModel.cs
│   └── SettingsViewModel.cs
├── Views/
│   ├── MainWindow.xaml
│   ├── EditorView.xaml
│   ├── PreviewView.xaml
│   └── SettingsView.xaml
├── Services/
│   ├── IMarkdownParser.cs
│   ├── MarkdownParser.cs
│   ├── IDocxGenerator.cs
│   ├── DocxGenerator.cs
│   ├── IPreviewService.cs
│   ├── WordInteropPreviewService.cs
│   └── IFileService.cs
├── Converters/
│   ├── MarkdownToHtmlConverter.cs
│   └── ValidationErrorConverter.cs
├── Helpers/
│   ├── MarkdownSyntaxHighlighting.cs
│   └── RelayCommand.cs
└── Resources/
    ├── Styles/
    │   └── AppStyles.xaml
    └── Templates/
        └── DefaultTemplate.docx
```

#### Core Classes

**Models/TestData.cs:**
```csharp
public class TestData
{
    public string Subject { get; set; }
    public string Title { get; set; }
    public List<Section> Sections { get; set; }
}

public class Section
{
    public string Name { get; set; }
    public SectionType? Type { get; set; }
    public List<Question> Questions { get; set; }
    public bool SeparateSheet { get; set; }
    public string Subtitle { get; set; }
}

public enum SectionType
{
    Short,
    Long,
    MatchingV,
    MatchingH,
    Blanks,
    Oral
}

public abstract class Question
{
    public string Text { get; set; }
}

public class TextQuestion : Question
{
    public int? Lines { get; set; }
}

public class MatchingQuestion : Question
{
    public string Left { get; set; }
    public string Right { get; set; }
}

public class BlankQuestion : Question { }

public class OralQuestion : Question
{
    public List<string> SubPoints { get; set; }
}
```

### 3. User Interface Design

#### Main Window Layout

```
┌─────────────────────────────────────────────────────────────┐
│ [File] [Edit] [View] [Help]                    [Settings] ☰  │
├─────────────────────────────────────────────────────────────┤
│ ┌─────────────────────┬─┬─────────────────────────────────┐ │
│ │                     │ │                                 │ │
│ │  MARKDOWN EDITOR    │░│     DOCX PREVIEW                │ │
│ │                     │░│                                 │ │
│ │  # Test             │░│  ┌────────────────────────────┐ │ │
│ │  Subject: Science   │░│  │  Name: _____ Date: _____   │ │ │
│ │  Title: Quiz        │░│  │                             │ │ │
│ │                     │░│  │      Science Test           │ │ │
│ │  ## Section: Short  │░│  │         Quiz                │ │ │
│ │  Type: short        │░│  │                             │ │ │
│ │  - Question 1       │░│  │  Short Answers              │ │ │
│ │  - Question 2       │░│  │  1. Question 1              │ │ │
│ │                     │░│  │     _______________         │ │ │
│ │                     │░│  │  2. Question 2              │ │ │
│ │  [line numbers]     │░│  │     _______________         │ │ │
│ │  [syntax highlight] │░│  │                             │ │ │
│ │                     │░│  └────────────────────────────┘ │ │
│ │                     │░│                                 │ │
│ └─────────────────────┴─┴─────────────────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│ Status: Ready | Line 1, Col 1 | [Generate DOCX] [Export...] │
└─────────────────────────────────────────────────────────────┘
```

#### Key UI Features

1. **Split-Pane Layout**
   - Resizable vertical splitter
   - Left pane: Markdown editor
   - Right pane: Live preview
   - Optional horizontal split for validation/errors

2. **Markdown Editor (Left Pane)**
   - Line numbers
   - Syntax highlighting for:
     - Headers (`#`, `##`)
     - Metadata fields (`Subject:`, `Title:`, `Type:`)
     - List items (`-`)
     - Markdown formatting (`**bold**`, `*italic*`)
     - Matching syntax (`->`)
     - Blanks (`___`)
   - Auto-completion for section types
   - Bracket matching
   - Code folding for sections
   - Search and replace
   - Undo/redo with history

3. **Live Preview (Right Pane)**
   - Real-time HTML rendering
   - Approximates DOCX output
   - Synchronized scrolling option
   - Zoom controls (50%, 75%, 100%, 125%, 150%)
   - "Export to DOCX" quick action

4. **Menu Bar**
   - **File:**
     - New (Ctrl+N)
     - Open (Ctrl+O)
     - Save (Ctrl+S)
     - Save As (Ctrl+Shift+S)
     - Recent Files →
     - Export to DOCX (Ctrl+E)
     - Export to PDF
     - Exit
   - **Edit:**
     - Undo (Ctrl+Z)
     - Redo (Ctrl+Y)
     - Cut (Ctrl+X)
     - Copy (Ctrl+C)
     - Paste (Ctrl+V)
     - Find (Ctrl+F)
     - Replace (Ctrl+H)
     - Insert Section Template →
   - **View:**
     - Toggle Preview (F5)
     - Toggle Line Numbers
     - Toggle Syntax Highlighting
     - Zoom In (Ctrl++)
     - Zoom Out (Ctrl+-)
     - Reset Zoom (Ctrl+0)
     - Synchronized Scrolling
   - **Help:**
     - Documentation
     - Syntax Reference
     - Sample Template
     - About

5. **Toolbar (Quick Actions)**
   - Save button
   - Undo/Redo buttons
   - Bold/Italic formatting helpers
   - Section templates dropdown
   - Generate DOCX button (prominent)
   - Settings button

6. **Status Bar**
   - Current line and column
   - Character/word count
   - Validation status
   - Preview status (Loading/Ready/Error)
   - Export button

### 4. Key Features and Functionality

#### 4.1 Auto-Save and Session Management
- Auto-save every 2 minutes to temp location
- Restore unsaved work on crash
- Session state preservation
- Recent files list (up to 10)

#### 4.2 Real-Time Validation
- Parse markdown as user types (with debounce)
- Highlight syntax errors inline
- Show validation messages:
  - Missing required fields (Subject, Title)
  - Invalid section types
  - Malformed matching questions
  - Empty sections
- Visual indicators in editor (squiggly underlines)

#### 4.3 Live Preview Updates
- Debounced preview refresh (500ms after last keystroke)
- Incremental updates for performance
- Loading indicator during generation
- Error display if preview fails

#### 4.4 Template Management
- Built-in default template
- Load custom DOCX templates
- Template selector in settings
- Template properties:
  - Page size
  - Margins
  - Fonts
  - Styles (Title, Subtitle, Heading 1/2)

#### 4.5 Export Options
- **Export to DOCX:**
  - File dialog with preview
  - Multiple format options
  - Template selection
  - Custom filename with smart defaults
- **Export to PDF:**
  - Convert DOCX → PDF
  - Requires PDF library (e.g., PdfSharp, iTextSharp)

#### 4.6 Section Templates
- Quick-insert templates via menu or shortcut
- Pre-configured section structures:
  - Short Answer Section
  - Essay Section
  - Matching Section
  - Fill-in-Blank Section
  - Oral Examination Section
- Customizable templates

#### 4.7 Editor Enhancements
- **Code Snippets:**
  - Type `short` + Tab → Insert short answer section
  - Type `long` + Tab → Insert essay section
  - Type `match` + Tab → Insert matching section
- **Smart Formatting:**
  - Auto-indent sub-points under oral questions
  - Auto-completion for section types
  - Bracket matching for markdown
- **Find and Replace:**
  - Regular expression support
  - Match case option
  - Whole word option

### 5. Implementation Roadmap

#### Phase 1: Core Infrastructure (Weeks 1-2)
- [x] Set up .NET WPF project
- [x] Implement MVVM architecture
- [x] Create basic UI layout with split panes
- [x] Integrate AvalonEdit for markdown editing
- [x] Basic syntax highlighting

#### Phase 2: Markdown Parsing (Week 3)
- [x] Port parser logic from Python/Rust
- [x] Implement TestData models
- [x] Create MarkdownParser service
- [x] Add real-time validation
- [x] Unit tests for parser

#### Phase 3: DOCX Generation (Weeks 4-5)
- [x] Integrate Open XML SDK
- [x] Port DOCX generator from Python/Rust
- [x] Implement all section types:
  - Short answer
  - Long answer
  - Matching (vertical & horizontal)
  - Fill-in-blanks
  - Oral questions
- [x] Template support
- [x] Unit tests for generator

#### Phase 4: Preview System (Week 6)
- [x] Implement Office Interop service
- [x] Integrate Word control via WindowsFormsHost
- [x] Word installation detection
- [x] Document lifecycle management (open/close/cleanup)
- [x] Debounced preview updates
- [x] Temp file management

#### Phase 5: Editor Features (Week 7)
- [x] Advanced syntax highlighting
- [x] Auto-completion
- [x] Code snippets
- [x] Find/Replace
- [x] Section templates

#### Phase 6: File Management (Week 8)
- [x] New/Open/Save/Save As
- [x] Auto-save functionality
- [x] Recent files
- [x] Session restoration
- [x] Template management

#### Phase 7: Polish and Testing (Weeks 9-10)
- [x] UI/UX refinements
- [x] Performance optimization
- [x] Comprehensive testing
- [x] Documentation
- [x] Sample templates
- [x] User guide

#### Phase 8: Advanced Features (Optional)
- [ ] PDF export
- [ ] Print preview
- [ ] Multiple document tabs
- [ ] Collaboration features
- [ ] Cloud storage integration
- [ ] Spell check integration

### 6. Technical Implementation Details

#### 6.1 Markdown Parser Implementation

Based on the Python/Rust implementations, the parser should:

1. **Read line-by-line with state management:**
   ```csharp
   public class MarkdownParser : IMarkdownParser
   {
       public TestData Parse(string markdown)
       {
           var lines = markdown.Split('\n');
           var testData = new TestData { Sections = new List<Section>() };
           Section currentSection = null;
           
           foreach (var line in lines)
           {
               var trimmed = line.TrimEnd();
               if (string.IsNullOrWhiteSpace(trimmed)) continue;
               
               if (trimmed.StartsWith("Subject:"))
                   testData.Subject = ExtractValue(trimmed);
               else if (trimmed.StartsWith("Title:"))
                   testData.Title = ExtractValue(trimmed);
               else if (trimmed.StartsWith("## Section:"))
               {
                   if (currentSection != null)
                       testData.Sections.Add(currentSection);
                   currentSection = new Section
                   {
                       Name = ExtractValue(trimmed),
                       Questions = new List<Question>()
                   };
               }
               // ... handle other cases
           }
           
           if (currentSection != null)
               testData.Sections.Add(currentSection);
               
           return testData;
       }
   }
   ```

2. **Handle all metadata fields:**
   - `Subject:`, `Title:` (test-level)
   - `Type:`, `Subtitle:`, `Separate Sheet:`, `Oral:` (section-level)

3. **Parse questions based on section type:**
   - Matching: Split on `->` arrow
   - Blanks: Preserve underscores
   - Oral: Track indented sub-points
   - Text: Extract optional line count `(N lines)`

4. **Support inline markdown:**
   - `**bold**` → Bold formatting
   - `*italic*` or `_italic_` → Italic formatting

#### 6.2 DOCX Generator Implementation

Key components based on existing implementations:

1. **Document Structure:**
   ```csharp
   public class DocxGenerator : IDocxGenerator
   {
       public void Generate(TestData data, string outputPath, string templatePath)
       {
           using (WordprocessingDocument doc = 
               WordprocessingDocument.Create(outputPath, 
                   WordprocessingDocumentType.Document))
           {
               // Create main document part
               MainDocumentPart mainPart = doc.AddMainDocumentPart();
               mainPart.Document = new Document();
               Body body = mainPart.Document.AppendChild(new Body());
               
               // Add header
               AddHeader(doc, mainPart);
               
               // Add subject and title
               AddSubjectAndTitle(body, data);
               
               // Add sections
               foreach (var section in data.Sections)
               {
                   AddSection(body, section, data);
               }
               
               mainPart.Document.Save();
           }
       }
   }
   ```

2. **Header Generation:**
   - Right-aligned
   - Name and Date fields with underlines
   - Use em-spaces for consistent spacing

3. **Section Rendering by Type:**
   - **Short:** Numbered questions with underlined blank lines
   - **Long:** Numbered questions with multiple blank lines or "separate sheet" note
   - **Matching (vertical):** Two-column table with shuffled items
   - **Matching (horizontal):** Term bank grid + definition table
   - **Blanks:** Questions with inline underlined blanks
   - **Oral:** Questions list + separate assessment sheet at end

4. **Styling:**
   - Use built-in or template styles
   - Heading 2 for sections
   - List Number for questions
   - Custom table styles for matching

5. **Advanced Features:**
   - Markdown to Word formatting (bold/italic)
   - Line spacing control
   - Page breaks for oral sheets
   - Cell borders and shading

#### 6.3 Office Interop Preview Implementation

**Word Interop Service:**

```csharp
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Runtime.InteropServices;

public class WordInteropPreviewService : IPreviewService, IDisposable
{
    private Application _wordApp;
    private Document _currentDocument;
    private readonly object _missing = System.Reflection.Missing.Value;
    private bool _isInitialized;

    public async Task InitializeAsync()
    {
        if (_isInitialized) return;
        
        try
        {
            // Create Word application instance (reused for all previews)
            _wordApp = new Application
            {
                Visible = false, // Run in background
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
            _isInitialized = true;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                "Microsoft Word is not installed or not properly registered.", ex);
        }
    }

    public async Task<bool> LoadDocumentAsync(string docxPath, object wordControlSite)
    {
        if (!_isInitialized)
            await InitializeAsync();

        try
        {
            // Close previous document if exists
            CloseCurrentDocument();

            // Open the DOCX file
            _currentDocument = _wordApp.Documents.Open(
                FileName: docxPath,
                ReadOnly: true,
                Visible: true,
                ref _missing, ref _missing, ref _missing,
                ref _missing, ref _missing, ref _missing,
                ref _missing, ref _missing, ref _missing,
                ref _missing, ref _missing, ref _missing,
                ref _missing);

            // Activate the document in the control
            _wordApp.Visible = true;
            _currentDocument.Activate();

            return true;
        }
        catch (Exception ex)
        {
            // Log error
            return false;
        }
    }

    public void CloseCurrentDocument()
    {
        if (_currentDocument != null)
        {
            try
            {
                _currentDocument.Close(SaveChanges: false);
                Marshal.ReleaseComObject(_currentDocument);
            }
            catch { }
            finally
            {
                _currentDocument = null;
            }
        }
    }

    public void Dispose()
    {
        CloseCurrentDocument();
        
        if (_wordApp != null)
        {
            try
            {
                _wordApp.Quit(SaveChanges: false);
                Marshal.ReleaseComObject(_wordApp);
            }
            catch { }
            finally
            {
                _wordApp = null;
            }
        }
        
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
```

**Preview Update Flow:**

```csharp
// In PreviewViewModel
private readonly WordInteropPreviewService _previewService;
private readonly IDocxGenerator _docxGenerator;
private readonly IMarkdownParser _parser;
private string _tempDocxPath;

private async Task UpdatePreview()
{
    try
    {
        // Parse markdown to data model
        var testData = _parser.Parse(MarkdownContent);
        
        // Generate DOCX to temp location
        _tempDocxPath = Path.Combine(Path.GetTempPath(), 
            $"declaratest_preview_{Guid.NewGuid()}.docx");
        
        _docxGenerator.Generate(testData, _tempDocxPath, TemplatePath);
        
        // Load into Word control
        await _previewService.LoadDocumentAsync(_tempDocxPath, WordControlHost);
        
        PreviewStatus = "Ready";
    }
    catch (Exception ex)
    {
        PreviewStatus = $"Error: {ex.Message}";
        // Log exception
    }
}

protected override void OnClosing()
{
    // Cleanup temp files
    if (!string.IsNullOrEmpty(_tempDocxPath) && File.Exists(_tempDocxPath))
    {
        try
        {
            File.Delete(_tempDocxPath);
        }
        catch { }
    }
    
    _previewService?.Dispose();
}
```

**XAML Integration:**

```xaml
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="5"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    
    <!-- Left: Markdown Editor -->
    <avalonedit:TextEditor Grid.Column="0" 
                          Name="MarkdownEditor"
                          FontFamily="Consolas"
                          FontSize="12"
                          ShowLineNumbers="True"/>
    
    <GridSplitter Grid.Column="1" 
                  HorizontalAlignment="Stretch"/>
    
    <!-- Right: Word Preview -->
    <Grid Grid.Column="2">
        <WindowsFormsHost Name="WordControlHost">
            <!-- Word control embedded here via code-behind -->
        </WindowsFormsHost>
        
        <!-- Status overlay -->
        <TextBlock Text="{Binding PreviewStatus}"
                   HorizontalAlignment="Center"
                   VerticalAlignment="Bottom"
                   Margin="10"
                   Background="#80FFFFFF"
                   Padding="10,5"/>
    </Grid>
</Grid>
```

**Installation Detection:**

```csharp
public static class WordDetector
{
    public static bool IsWordInstalled()
    {
        try
        {
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null) return false;
            
            // Try to create instance
            var wordApp = Activator.CreateInstance(wordType);
            if (wordApp != null)
            {
                Marshal.ReleaseComObject(wordApp);
                return true;
            }
        }
        catch
        {
            return false;
        }
        return false;
    }
    
    public static string GetWordVersion()
    {
        try
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            string version = wordApp.Version;
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            return version;
        }
        catch
        {
            return "Unknown";
        }
    }
}
```

#### 6.4 Syntax Highlighting

Using AvalonEdit's syntax highlighting system:

```csharp
public class MarkdownSyntaxHighlighting
{
    public static IHighlightingDefinition GetHighlighting()
    {
        var highlighting = new HighlightingDefinition();
        
        // Headers
        highlighting.MainRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"^##?\s+.*$"),
            Color = new HighlightingColor
            {
                Foreground = new SimpleHighlightingBrush(Colors.Blue),
                FontWeight = FontWeights.Bold
            }
        });
        
        // Metadata fields
        highlighting.MainRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"^(Subject|Title|Type|Subtitle|Separate Sheet|Oral):"),
            Color = new HighlightingColor
            {
                Foreground = new SimpleHighlightingBrush(Colors.Purple),
                FontWeight = FontWeights.Bold
            }
        });
        
        // List items
        highlighting.MainRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"^[\s]*-\s"),
            Color = new HighlightingColor
            {
                Foreground = new SimpleHighlightingBrush(Colors.Green)
            }
        });
        
        // Bold markdown
        highlighting.MainRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"\*\*[^*]+\*\*"),
            Color = new HighlightingColor
            {
                FontWeight = FontWeights.Bold
            }
        });
        
        // Italic markdown
        highlighting.MainRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"\*[^*]+\*|_[^_]+_"),
            Color = new HighlightingColor
            {
                FontStyle = FontStyles.Italic
            }
        });
        
        return highlighting;
    }
}
```

### 7. Testing Strategy

#### 7.1 Unit Tests
- **Parser Tests:**
  - Valid template parsing
  - Invalid syntax handling
  - Edge cases (empty sections, malformed matching)
  - All section types
  - Markdown formatting

- **Generator Tests:**
  - DOCX structure validation
  - Style application
  - Table generation
  - Template usage
  - Output file integrity

#### 7.2 Integration Tests
- End-to-end markdown to DOCX
- Template application
- File I/O operations
- Preview generation

#### 7.3 UI Tests
- Editor functionality
- Preview updates
- Menu actions
- Keyboard shortcuts
- File dialogs

#### 7.4 Performance Tests
- Large document handling (100+ questions)
- Preview generation time
- Memory usage
- Startup time

### 8. Deployment and Distribution

#### 8.1 Packaging Options

**Option A: ClickOnce Deployment**
- Simple installation
- Auto-updates
- Limited customization

**Option B: MSI Installer**
- Professional installer
- Custom branding
- Registry integration
- More control

**Option C: MSIX Package**
- Modern Windows packaging
- Microsoft Store distribution option
- Automatic updates
- Sandboxed execution

**RECOMMENDATION: MSIX for modern Windows, MSI for enterprise**

#### 8.2 Prerequisites
- .NET 8.0 Runtime
- Microsoft Word 2016 or later (required for preview functionality)
- Windows 10 version 1809 or later

#### 8.3 Installation
- One-click installer
- Optional: Desktop shortcut
- Optional: Start menu entry
- File association for `.md` files (optional)

### 9. Performance Considerations

#### 9.1 Memory Management
- Dispose COM objects properly (Office Interop)
- Release Word document references explicitly
- Call GC.Collect() after disposing Word instances
- Clear preview cache periodically
- Lazy-load large documents
- Use weak event handlers to prevent leaks
- Reuse Word Application instance across previews

#### 9.2 Responsiveness
- Async/await for all I/O operations
- Debounced preview updates (500ms)
- Background thread for DOCX generation
- Progress indicators for long operations
- Cancellation tokens for cancellable operations

#### 9.3 Optimization
- Incremental parsing (only re-parse changed sections)
- Cache parsed results
- Optimize regex patterns
- Minimize UI updates during batch operations

### 10. Error Handling and Logging

#### 10.1 User-Facing Errors
- Graceful error messages
- Specific error descriptions
- Suggested fixes
- Error reporting option

#### 10.2 Logging
- Use Serilog or NLog
- Log levels: Debug, Info, Warning, Error, Fatal
- Log to file in user's AppData
- Rotation policy (keep last 7 days)

#### 10.3 Crash Recovery
- Save editor state before operations
- Restore last session on startup after crash
- Offer to recover unsaved work

### 11. Security Considerations

#### 11.1 File Handling
- Validate file extensions
- Sanitize file paths
- Check file sizes
- Prevent path traversal attacks

#### 11.2 Content Security
- Sanitize markdown input before processing
- Validate DOCX structures
- Prevent XXE attacks in XML parsing
- Proper COM object disposal to prevent resource leaks
- Validate temp file paths to prevent directory traversal

#### 11.3 User Data
- Store user preferences securely
- Encrypt sensitive settings
- No telemetry without consent

### 12. Accessibility

#### 12.1 Keyboard Navigation
- Full keyboard accessibility
- Tab order optimization
- Keyboard shortcuts for all major actions
- Focus indicators

#### 12.2 Screen Reader Support
- ARIA labels
- Descriptive element names
- Status announcements
- Alternative text for icons

#### 12.3 Visual Accessibility
- High contrast mode support
- Configurable font sizes
- Color-blind friendly syntax highlighting
- Zoom functionality

### 13. Internationalization

While initial version is English-only, design for future localization:

- Separate resource files for strings
- Use resource keys instead of hardcoded strings
- Support for RTL languages in UI
- Locale-aware date/number formatting

### 14. Documentation Requirements

#### 14.1 User Documentation
- Quick start guide
- Full user manual
- Syntax reference (quick reference card)
- Video tutorials (optional)
- FAQ

#### 14.2 Developer Documentation
- API documentation (XML comments)
- Architecture overview
- Extension guide
- Build instructions

#### 14.3 Sample Content
- Sample template files
- Template gallery
- Best practices guide

### 15. Future Enhancement Ideas

#### 15.1 Advanced Features
- Multiple document tabs
- Split preview (before/after editing)
- Version control integration
- Collaborative editing
- Cloud storage integration (OneDrive, Dropbox)

#### 15.2 Content Features
- Question bank/library
- Randomization options
- Answer key generation
- Rubric creation
- Grade calculation

#### 15.3 Export Features
- HTML export
- PDF export
- Google Docs format
- Canvas LMS integration
- Moodle export

#### 15.4 Template Features
- Visual template editor
- Template marketplace
- Custom style designer
- Theme support

### 16. Dependencies and NuGet Packages

Recommended packages:

```xml
<ItemGroup>
  <!-- Core Framework -->
  <PackageReference Include="Microsoft.NET.Sdk.WindowsDesktop" Version="8.0.0" />
  
  <!-- DOCX Generation -->
  <PackageReference Include="DocumentFormat.OpenXml" Version="3.0.0" />
  
  <!-- Markdown Editor -->
  <PackageReference Include="AvalonEdit" Version="6.3.0" />
  
  <!-- Office Interop for Preview -->
  <PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.4797.1004" />
  
  <!-- Windows Forms Integration (for hosting Word control) -->
  <PackageReference Include="System.Windows.Forms" Version="8.0.0" />
  
  <!-- MVVM Helpers -->
  <PackageReference Include="CommunityToolkit.Mvvm" Version="8.2.2" />
  
  <!-- Logging -->
  <PackageReference Include="Serilog" Version="3.1.1" />
  <PackageReference Include="Serilog.Sinks.File" Version="5.0.0" />
  
  <!-- JSON Configuration -->
  <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
  
  <!-- PDF Export (Optional) -->
  <PackageReference Include="PdfSharp" Version="6.0.0" />
  
  <!-- Testing -->
  <PackageReference Include="xunit" Version="2.6.2" />
  <PackageReference Include="Moq" Version="4.20.69" />
</ItemGroup>
```

**Additional COM References:**

The application also requires COM Interop types for Office integration. These are typically automatically generated when using `Microsoft.Office.Interop.Word`, but ensure the following is in your project file:

```xml
<PropertyGroup>
  <EnableComHosting>true</EnableComHosting>
  <EmbedInteropTypes>false</EmbedInteropTypes>
</PropertyGroup>
```

### 17. Minimum System Requirements

- **Operating System:** Windows 10 (version 1809) or later
- **Processor:** 1 GHz or faster
- **RAM:** 4 GB minimum, 8 GB recommended (for Word preview)
- **Disk Space:** 500 MB for application + Word installation
- **Display:** 1280x720 minimum resolution
- **.NET Runtime:** .NET 8.0 or later
- **Additional Software:** Microsoft Word 2016 or later (required for preview functionality)

### 18. Getting Started for Developers

#### 18.1 Development Environment Setup

1. **Install Prerequisites:**
   - Visual Studio 2022 (Community or higher)
   - .NET 8.0 SDK
   - Windows 10 SDK

2. **Clone Repository:**
   ```bash
   git clone https://github.com/NiceneNerd/declaratest.git
   cd declaratest
   ```

3. **Create Solution:**
   ```bash
   dotnet new sln -n Declaratest
   dotnet new wpf -n Declaratest.GUI
   dotnet sln add Declaratest.GUI
   ```

4. **Add Dependencies:**
   - Install NuGet packages listed above
   - Reference AvalonEdit
   - Reference Open XML SDK

5. **Initial Architecture:**
   - Set up MVVM folders
   - Create base classes
   - Implement dependency injection

#### 18.2 First Milestone: Basic Editor

Goal: Split-pane editor with markdown on left, Word preview on right

1. Create MainWindow with Grid layout
2. Add AvalonEdit to left pane
3. Add WindowsFormsHost for Word control to right pane
4. Implement basic syntax highlighting
5. Test with sample template

#### 18.3 Second Milestone: Parser

Goal: Parse markdown template to TestData object

1. Port parser logic from Python/Rust
2. Create unit tests
3. Implement validation
4. Connect to editor for real-time parsing

#### 18.4 Third Milestone: Generator

Goal: Generate DOCX from TestData

1. Port generator logic from Python/Rust
2. Implement all section types
3. Create unit tests
4. Test output compatibility with original

#### 18.5 Fourth Milestone: Preview

Goal: Live Office Interop preview

1. Implement WordInteropPreviewService
2. Handle Word COM object lifecycle
3. Implement document loading and display
4. Add debounced updates
5. Implement error handling and Word detection

### 19. Code Style and Best Practices

#### 19.1 C# Conventions
- Follow Microsoft C# Coding Conventions
- Use async/await for I/O operations
- Implement INotifyPropertyChanged for data binding
- Use dependency injection
- Prefer LINQ for collections
- Use nullable reference types

#### 19.2 XAML Conventions
- Use data binding over code-behind
- Separate concerns (no business logic in views)
- Use resources for reusable styles
- Name controls meaningfully
- Use attached properties appropriately

#### 19.3 Git Conventions
- Meaningful commit messages
- Feature branches
- Pull request reviews
- Keep commits atomic

### 20. Licensing and Legal

Consider licensing options:

- **MIT License:** Permissive, widely adopted
- **GPL v3:** Copyleft, ensures open source
- **Apache 2.0:** Patent grant, permissive
- **Proprietary:** Commercial licensing

Ensure compliance with dependencies:
- Open XML SDK: MIT License
- AvalonEdit: MIT License
- Microsoft Office Interop: Included with Microsoft Office license
  - Requires end users to have valid Microsoft Office license
  - No additional licensing costs for application distribution
  - Standard Microsoft Office EULA applies

### 21. Success Metrics

How to measure success:

- **User Satisfaction:**
  - Preview latency < 1 second (initial Word load)
  - Preview update < 500ms (subsequent updates)
  - DOCX generation < 2 seconds for typical test
  - Startup time < 3 seconds (plus Word initialization)
  
- **Reliability:**
  - Crash rate < 0.1%
  - Data loss rate: 0%
  - Parser success rate > 99%

- **Usability:**
  - New user can create test in < 5 minutes
  - Feature discovery rate > 80%
  - Keyboard navigation coverage: 100%

## Conclusion

This specification provides a comprehensive blueprint for creating a Windows GUI version of Declaratest. The chosen technology stack (**WPF + Open XML SDK + Office Interop**) provides the optimal solution for this application:

- **Exact Fidelity:** Office Interop ensures pixel-perfect preview matching final output
- **Maturity:** WPF is a proven, well-established Windows UI technology
- **Performance:** Native Windows performance with hardware acceleration
- **Maintainability:** Large community, extensive documentation, and abundant resources
- **Flexibility:** Complete control over DOCX generation and preview
- **Professional Quality:** Leverages Microsoft Word's rendering engine for production-grade output

**Key Advantages of This Approach:**

1. **Perfect Preview Accuracy:** What teachers see is exactly what students will receive
2. **Zero Conversion Artifacts:** No HTML/PDF intermediary means no formatting surprises
3. **Enterprise Ready:** Suitable for educational institutions already using Microsoft Office
4. **Reduced Development Risk:** Well-documented COM Interop patterns and extensive examples
5. **Future Proof:** Compatible with current and future Word versions

The modular MVVM architecture allows for incremental development and future enhancements. By following this specification, developers can create a production-quality application that significantly improves the user experience over the command-line versions while maintaining 100% compatibility with the existing template format.

**Target Audience Alignment:**

This solution is perfectly aligned with the educational market where:
- Microsoft Office is nearly ubiquitous
- Accurate document rendering is critical
- Professional appearance matters for assessment materials
- IT support for Office applications is readily available

### Next Steps

1. **Review and Approval:** Review this specification with stakeholders
2. **Prototype:** Create a proof-of-concept for critical features
3. **Development:** Follow the phased implementation roadmap
4. **Testing:** Comprehensive testing at each phase
5. **Beta Release:** Limited release for user feedback
6. **Production Release:** Full public release with documentation

---

**Document Version:** 2.0  
**Date:** February 8, 2026  
**Author:** Declaratest Development Team  
**Status:** Updated - Committed to WPF and Office Interop
**Changes from v1.0:**
- Committed to WPF as the UI framework (removed alternatives)
- Committed to Office Interop for DOCX preview (replaced WebView2 approach)
- Updated code examples to reflect Office Interop implementation
- Updated NuGet packages and prerequisites
- Updated performance and memory management considerations
