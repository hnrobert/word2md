# DOCX to Markdown Converter

A Python-based DOCX to Markdown converter that supports converting Microsoft Word documents to Markdown format.

## Features

- ✅ Support for heading conversion (H1-H6)
- ✅ Support for paragraph text
- ✅ Support for bold, italic, underline formatting
- ✅ Support for ordered and unordered lists
- ✅ Support for table conversion
- ✅ Support for image extraction and conversion
- ✅ Automatic folder structure creation
- ✅ Automatic blank line and format cleanup
- ✅ Command line interface
- ✅ Batch conversion support
- ✅ Smart title handling with proper heading level adjustment
- ✅ Intelligent formatting merge (e.g., adjacent underline tags)
- ✅ Font-size based heading detection (when no heading styles are present)

## Installation

### Install from source

```bash
git clone https://github.com/HNRobert/docx2md.git
cd docx2md
pip install -e .
```

### Install dependencies only

```bash
pip install -r requirements.txt
```

Or install directly:

```bash
pip install python-docx
```

## Usage

### Command Line Tool

After installation, you can use the `docx2md` command:

```bash
# Convert single file
docx2md document.docx

# Specify output file
docx2md document.docx -o output.md

# Show verbose output
docx2md document.docx -v

# Batch conversion
docx2md *.docx -o output_directory/
```

### Python Script

You can also run the converter directly:

```bash
# Convert single file to auto-generated folder structure
python main.py document.docx

# Specify output file
python main.py document.docx -o output.md

# Show verbose output
python main.py document.docx -o output.md -v
```

### Advanced Usage

```bash
# Batch conversion to output directory
python main.py *.docx -o output_directory/

# Output to stdout
python main.py document.docx
```

## Project Structure

The project is now organized as a modular package:

```text
docx2md/
├── main.py                    # Main entry point
├── docx_converter/            # Main package
│   ├── __init__.py           # Package initialization
│   ├── cli.py                # Command line interface
│   ├── converter.py          # Main converter class
│   ├── document_processor.py # Document processing logic
│   ├── paragraph_processor.py # Paragraph processing
│   ├── formatting.py         # Text formatting (bold, italic, etc.)
│   ├── list_processor.py     # List handling
│   ├── table_processor.py    # Table conversion
│   ├── image_processor.py    # Image processing in paragraphs
│   ├── image_extractor.py    # Image extraction from DOCX
│   └── utils.py              # Utility functions
├── assets/
│   └── sample.docx           # Sample test file
├── requirements.txt          # Dependencies
└── README.md                # Documentation
```

## Supported Formats

### Text Formatting

- **Bold** → `**Bold**`
- _Italic_ → `*Italic*`
- <u>Underline</u> → `<u>Underline</u>`

### Heading Detection

The converter supports multiple methods for detecting headings:

1. **Style-based detection**: Converts Word heading styles (Heading 1-6, Title) to Markdown headings
2. **Font-size based detection**: When no heading styles are present, automatically detects headings based on font size hierarchy
   - Analyzes all paragraphs with uniform font sizes
   - Determines the baseline font size (most common size, usually normal text)
   - Assigns heading levels to larger font sizes in descending order
   - Example: If baseline is 12pt, then 18pt → # (H1), 16pt → ## (H2), 14pt → ### (H3)

### Headings

- Word heading styles → Markdown headings (# ## ### etc.)
- Smart title handling: When a "Title" style is present, all other headings are automatically adjusted down one level

### Lists

- Unordered lists (•, -, \* etc.) → `- Item`
- Ordered lists (1., 2., etc.) → `1. Item`

### Tables

- Word tables → Markdown table format

### Images

- Automatic extraction of images from DOCX
- Save to `assets/` directory under document name folder
- Create proper image references in Markdown: `![Image](./assets/image_001.png)`

## Output Structure

After conversion, the following structure is created:

```text
document_name/
├── document_name.md
└── assets/
    ├── image_001.jpg
    ├── image_002.png
    └── ...
```

## Example

### Input (DOCX)

A document with the following structure:

- Title style: "TEST DOC"
- Heading 1: "Title 1"
- Heading 2: "Title 2"
- Heading 3: "Title 3"
- Various text formatting including bold, italic, and underlined text

### Output (Markdown)

```markdown
# TEST DOC

## Title 1

### Title 2

#### Title 3

This is a paragraph with **bold text**, _italic text_, and <u>underlined text</u>.

- Unordered list item 1
- Unordered list item 2

1. Ordered list item 1
2. Ordered list item 2

![Image](./assets/image_001.jpg)
```

## Development

### Current Project Structure

```text
docx2md/
├── main.py                    # Main entry point
├── docx_converter/            # Main package
│   ├── __init__.py           # Package initialization
│   ├── cli.py                # Command line interface
│   ├── converter.py          # Main converter class
│   ├── document_processor.py # Document processing logic
│   ├── paragraph_processor.py # Paragraph processing
│   ├── formatting.py         # Text formatting (bold, italic, etc.)
│   ├── list_processor.py     # List handling
│   ├── table_processor.py    # Table conversion
│   ├── image_processor.py    # Image processing in paragraphs
│   ├── image_extractor.py    # Image extraction from DOCX
│   └── utils.py              # Utility functions
├── assets/
│   └── sample.docx           # Sample test file
├── requirements.txt          # Dependencies
└── README.md                # Documentation
```

### Architecture Benefits

- **Modular Design**: Each component has a single responsibility
- **Easy Testing**: Individual modules can be tested independently
- **Maintainable**: Clear separation of concerns
- **Extensible**: Easy to add new features or modify existing ones

### Key Modules

- **`DocxToMarkdownConverter`**: Main orchestrator class
- **`DocumentProcessor`**: Handles document-level processing and title detection
- **`ParagraphProcessor`**: Manages paragraph conversion and formatting
- **`ImageExtractor`**: Extracts and maps images from DOCX files
- **`ListProcessor`**: Handles ordered and unordered list conversion
- **`TableProcessor`**: Converts Word tables to Markdown format
- **`TextFormatter`**: Handles text formatting (bold, italic, underline)

### Extending Functionality

The modular structure makes it easy to extend functionality:

#### Adding New Text Formatting

Edit `docx_converter/formatting.py` to add support for new text styles.

#### Supporting New List Types

Modify `docx_converter/list_processor.py` to handle different list formats.

#### Enhancing Image Processing

Update `docx_converter/image_processor.py` and `docx_converter/image_extractor.py` for advanced image handling.

#### Custom Document Elements

Add new processors in the `docx_converter/` directory and integrate them via `document_processor.py`.

### Development Workflow

1. Install dependencies: `pip install -r requirements.txt`
2. Run tests: `python main.py assets/sample.docx`
3. Add new features in appropriate modules
4. Test with various DOCX files
5. Update documentation

## Notes

1. The converter primarily supports basic document formats; complex formatting may require manual adjustment
2. Images are automatically extracted and saved to the assets folder
3. Complex table layouts may need manual optimization
4. Some Word-specific formats have no equivalent in Markdown and will be simplified

## License

MIT License

## Contributing

Issues and Pull Requests are welcome to improve this converter.
