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

## Installation

```bash
pip install -r requirements.txt
```

Or install directly:

```bash
pip install python-docx
```

## Usage

### Basic Usage

```bash
# Convert single file to auto-generated folder structure
python src/main.py document.docx

# Specify output file
python src/main.py document.docx -o output.md

# Show verbose output
python src/main.py document.docx -o output.md -v
```

### Advanced Usage

```bash
# Batch conversion to output directory
python src/main.py *.docx -o output_directory/

# Output to stdout
python src/main.py document.docx
```

## Project Structure

The project provides a comprehensive converter with advanced features:

- **src/main.py** - Full-featured converter with image extraction, smart title handling, and batch processing

## Supported Formats

### Text Formatting

- **Bold** → `**Bold**`
- _Italic_ → `*Italic*`
- <u>Underline</u> → `<u>Underline</u>`

### Headings

- Word heading styles → Markdown headings (# ## ### etc.)
- Smart title handling: When a "Title" style is present, all other headings are automatically adjusted down one level

### Lists

- Unordered lists (•, -, * etc.) → `- Item`
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

This is a paragraph with **bold text**, *italic text*, and <u>underlined text</u>.

- Unordered list item 1
- Unordered list item 2

1. Ordered list item 1
2. Ordered list item 2

![Image](./assets/image_001.jpg)
```

## Development

### Current Project Structure

```text
docx-markdown-converter/
├── src/
│   └── main.py            # Full-featured converter
├── assets/
│   └── sample.docx        # Sample test file
├── requirements.txt       # Dependencies
├── .venv/                 # Virtual environment
└── README.md             # Documentation
```

### Key Features Implementation

The converter includes several advanced features:

- **Smart Title Handling**: Automatically detects "Title" style and adjusts all heading levels accordingly
- **Format Merging**: Intelligently merges adjacent formatting tags (e.g., `<u>Under</u><u>lined</u>` → `<u>Underlined</u>`)
- **Image Extraction**: Extracts images from DOCX files and creates proper folder structure
- **List Detection**: Handles various list formats and styles
- **Table Conversion**: Converts Word tables to Markdown format

### Extending Functionality

To add more features, modify the converter class methods:

- `_convert_document()` - Main document processing
- `_convert_paragraph()` - Paragraph processing
- `_convert_paragraph_formatting()` - Text formatting
- `_convert_table()` - Table processing
- `_extract_images()` - Image extraction

## Notes

1. The converter primarily supports basic document formats; complex formatting may require manual adjustment
2. Images are automatically extracted and saved to the assets folder
3. Complex table layouts may need manual optimization
4. Some Word-specific formats have no equivalent in Markdown and will be simplified

## License

MIT License

## Contributing

Issues and Pull Requests are welcome to improve this converter.
