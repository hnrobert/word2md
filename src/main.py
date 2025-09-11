#!/usr/bin/env python3
"""
DOCX to Markdown Converter

A Python tool for converting Microsoft Word documents (.docx) to Markdown format.
Supports conversion of text, headings, lists, tables, links and other basic formats.
"""

import argparse
import logging
import os
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Optional

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Inches
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DocxToMarkdownConverter:
    """DOCX to Markdown converter class"""

    def __init__(self):
        self.output_lines = []
        self.current_list_level = 0
        self.list_counters = {}  # Track ordered list counters
        self.in_list = False  # Track if currently in a list
        self.list_type = None  # Track current list type
        self.image_counter = 0  # Image counter
        self.assets_dir = None  # Asset folder path
        self.output_folder = None  # Output folder path
        self.image_map = {}  # Image ID to filename mapping
        self.heading_offset = 0  # Heading level offset

    def convert_file(self, input_path: str, output_path: Optional[str] = None) -> str:
        """
        Convert DOCX file to Markdown format

        Args:
            input_path: Input DOCX file path
            output_path: Output Markdown file path (optional)

        Returns:
            Markdown content string
        """
        try:
            # Check if input file exists
            if not os.path.exists(input_path):
                raise FileNotFoundError(f"Input file does not exist: {input_path}")

            # Determine output path and folder structure
            input_stem = Path(input_path).stem
            if output_path:
                if os.path.isdir(output_path) or output_path.endswith('/'):
                    # Output to directory
                    self.output_folder = os.path.join(output_path, input_stem)
                    final_output_path = os.path.join(
                        self.output_folder, f"{input_stem}.md")
                else:
                    # Output to specified file
                    final_output_path = output_path
                    self.output_folder = os.path.dirname(output_path)
                    if not self.output_folder:
                        self.output_folder = input_stem
            else:
                # Default: create folder named after file
                self.output_folder = input_stem
                final_output_path = os.path.join(
                    self.output_folder, f"{input_stem}.md")

            # Create output folder and assets folder
            os.makedirs(self.output_folder, exist_ok=True)
            self.assets_dir = os.path.join(self.output_folder, "assets")
            os.makedirs(self.assets_dir, exist_ok=True)

            # Load DOCX document
            logger.info(f"Loading document: {input_path}")
            doc = Document(input_path)

            # Reset output
            self.output_lines = []
            self.current_list_level = 0
            self.list_counters = {}
            self.in_list = False
            self.list_type = None
            self.image_counter = 0
            self.heading_offset = 0

            # Convert document content
            self._convert_document(doc, input_path)

            # Generate Markdown content
            markdown_content = '\n'.join(self.output_lines)

            # Clean up extra blank lines - merge multiple consecutive blank lines into single blank line
            import re
            markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)
            # Remove blank lines at beginning and end
            markdown_content = markdown_content.strip()

            # Add extra blank line at the end
            markdown_content += '\n'

            # Write to file
            self._write_output(markdown_content, final_output_path)
            logger.info(f"Conversion completed, output file: {final_output_path}")

            return markdown_content

        except Exception as e:
            logger.error(f"Error occurred during conversion: {str(e)}")
            raise

    def _convert_document(self, doc, input_path):
        """Convert main document content"""
        # First extract all images
        self._extract_images(input_path)

        # First check if there are Title style paragraphs, if so use as main title
        title_found = False
        first_heading_found = False

        # First pass: check if Title style exists
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                if 'title' in style_name and paragraph.text.strip():
                    title_found = True
                    break

        # Set heading offset: if Title style exists, all headings are adjusted down one level
        self.heading_offset = 1 if title_found else 0

        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                # Check Title style
                if 'title' in style_name and paragraph.text.strip():
                    self.output_lines.append(f"# {paragraph.text.strip()}")
                    self.output_lines.append('')
                    continue

                # If no Title, first Heading 1 becomes main title
                if not title_found and not first_heading_found and 'heading 1' in style_name and paragraph.text.strip():
                    self.output_lines.append(f"# {paragraph.text.strip()}")
                    self.output_lines.append('')
                    first_heading_found = True
                    continue

                self._convert_paragraph(paragraph)
            elif element.tag.endswith('tbl'):  # Table
                table = Table(element, doc)
                self._convert_table(table)

    def _convert_paragraph(self, paragraph: Paragraph):
        """Convert paragraph"""
        # Get paragraph text
        text = paragraph.text.strip()

        # First check if paragraph contains images (regardless of text content)
        images_text = self._process_paragraph_images(paragraph)

        # If paragraph is mainly images (no text or very little text)
        if images_text and (not text or len(text) < 3):
            self.output_lines.append(images_text)
            self.output_lines.append('')
            return

        # Skip empty paragraphs but keep one blank line for separation
        if not text and not images_text:
            # If previous line is not empty, add blank line
            if self.output_lines and self.output_lines[-1] != '':
                self.output_lines.append('')
            return

        # Check paragraph style
        style_name = paragraph.style.name.lower(
        ) if paragraph.style and paragraph.style.name else ''

        # Skip Title style, already handled in _convert_document
        if 'title' in style_name:
            return

        # Check if it's a list item
        is_list = self._is_list_paragraph(paragraph)

        # If previously in list but current is not list item, list ends
        if self.in_list and not is_list:
            self.output_lines.append('')  # Add blank line to separate list and subsequent content
            self.in_list = False
            self.list_type = None

        # Handle headings (adjust level based on Title style presence)
        if 'heading' in style_name:
            level = self._extract_heading_level(style_name)

            # If Title style exists, all headings are adjusted down one level
            if hasattr(self, 'heading_offset'):
                level += self.heading_offset
            else:
                # Backward compatibility: if it's Heading 1 and no Title, convert to Heading 2
                if level == 1:
                    level = 2

            # Ensure not exceeding 6 heading levels
            level = min(level, 6)

            self.output_lines.append(f"{'#' * level} {text}")
            self.output_lines.append('')
            return

        # Handle lists
        if is_list:
            self._convert_list_item(paragraph)
            return

        # Handle regular paragraphs
        # If paragraph contains images, insert images first
        if images_text:
            self.output_lines.append(images_text)
            self.output_lines.append('')

        # Handle text content
        if text:  # Only process when paragraph has text
            markdown_text = self._convert_paragraph_formatting(paragraph)
            self.output_lines.append(markdown_text)
            self.output_lines.append('')

    def _extract_heading_level(self, style_name: str) -> int:
        """Extract heading level from style name"""
        import re
        match = re.search(r'heading\s*(\d+)', style_name)
        if match:
            return min(int(match.group(1)), 6)  # Markdown supports maximum 6 heading levels
        return 1

    def _is_list_paragraph(self, paragraph: Paragraph) -> bool:
        """Check if paragraph is a list item"""
        # Check paragraph numbering format
        if paragraph._element.pPr is not None:
            numPr = paragraph._element.pPr.find(
                './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
            if numPr is not None:
                return True

        # Check if paragraph style is a list style
        style_name = paragraph.style.name.lower(
        ) if paragraph.style and paragraph.style.name else ''
        if 'list' in style_name or 'bullet' in style_name:
            return True

        # Check if text starts with list markers
        text = paragraph.text.strip()
        list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
        if any(text.startswith(marker + ' ') for marker in list_markers):
            return True

        # Check if it's a numbered list
        import re
        if re.match(r'^\d+[\.）]\s+', text):
            return True

        return False

    def _convert_list_item(self, paragraph: Paragraph):
        """Convert list item"""
        text = paragraph.text.strip()

        # Check paragraph style
        style_name = paragraph.style.name.lower(
        ) if paragraph.style and paragraph.style.name else ''

        # Determine list type
        is_ordered = False

        # Check if it's an ordered list (starts with number)
        import re
        if re.match(r'^\d+[\.）]\s+', text):
            is_ordered = True
            text = re.sub(r'^\d+[\.）]\s+', '', text)
        elif ('list' in style_name or 'bullet' in style_name) and not any(text.startswith(marker + ' ') for marker in ['•', '◦', '▪', '▫', '‣', '-', '*', '+']):
            # If it's a list style but no obvious markers, check if it's ordered list style
            if 'number' in style_name or 'ordered' in style_name:
                is_ordered = True
        else:
            # Remove common unordered list markers
            list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
            for marker in list_markers:
                if text.startswith(marker + ' '):
                    text = text[len(marker):].strip()
                    break

        # Detect list type change or start new list
        current_list_type = 'ordered' if is_ordered else 'unordered'

        if not self.in_list or self.list_type != current_list_type:
            # Start new list or list type changed
            self.in_list = True
            self.list_type = current_list_type
            if is_ordered:
                self.list_counters[self.current_list_level] = 1

        # Generate list marker
        if is_ordered:
            # Use counter to generate correct sequence number
            counter = self.list_counters.get(self.current_list_level, 1)
            list_marker = f"{counter}."
            self.list_counters[self.current_list_level] = counter + 1
        else:
            list_marker = '-'

        # Add appropriate indentation
        indent = '  ' * self.current_list_level
        formatted_text = self._convert_paragraph_formatting(paragraph, text)
        self.output_lines.append(f"{indent}{list_marker} {formatted_text}")

    def _convert_paragraph_formatting(self, paragraph: Paragraph, custom_text: Optional[str] = None) -> str:
        """Convert paragraph formatting (bold, italic, links, etc.)"""
        if custom_text:
            # If custom text is provided, use simplified processing
            return custom_text

        result = []
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            # Apply formatting
            if run.bold:
                text = f"**{text}**"
            if run.italic:
                text = f"*{text}*"
            if run.underline:
                text = f"<u>{text}</u>"

            result.append(text)

        # Merge adjacent same HTML tags
        final_result = ''.join(result)

        # Merge adjacent underline tags
        import re
        final_result = re.sub(r'</u><u>', '', final_result)

        return final_result

    def _convert_table(self, table: Table):
        """Convert table"""
        self.output_lines.append('')  # Blank line before table

        # Convert table rows
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' ')
                     for cell in row.cells]

            # Table row
            self.output_lines.append('| ' + ' | '.join(cells) + ' |')

            # Add header separator (after first row)
            if i == 0:
                separator = '|' + ''.join([' --- |' for _ in cells])
                self.output_lines.append(separator)

        self.output_lines.append('')  # Blank line after table

    def _write_output(self, content: str, output_path: str):
        """Write output file"""
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def _extract_images(self, docx_path: str):
        """Extract images from DOCX file and establish mapping relationship"""
        if not self.assets_dir:
            return

        try:
            # Reset image counter and mapping
            self.image_counter = 0
            self.image_map = {}

            # DOCX file is actually a ZIP file
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                # Read relationship file to get image relationship mapping
                try:
                    rels_content = docx_zip.read(
                        'word/_rels/document.xml.rels').decode('utf-8')
                    import xml.etree.ElementTree as ET
                    rels_root = ET.fromstring(rels_content)

                    # Establish relationship ID to image file mapping
                    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        rel_type = rel.get('Type', '')
                        if 'image' in rel_type.lower():
                            rel_id = rel.get('Id')
                            target = rel.get('Target')
                            if target and target.startswith('media/'):
                                full_path = f"word/{target}"
                                if full_path in [f.filename for f in docx_zip.filelist]:
                                    # Extract image
                                    self.image_counter += 1
                                    file_ext = os.path.splitext(target)[
                                        1].lower()
                                    new_filename = f"image_{self.image_counter:03d}{file_ext}"
                                    output_path = os.path.join(
                                        self.assets_dir, new_filename)

                                    with docx_zip.open(full_path) as source:
                                        with open(output_path, 'wb') as target_file:
                                            shutil.copyfileobj(
                                                source, target_file)

                                    # Establish mapping relationship
                                    self.image_map[rel_id] = new_filename
                                    logger.info(
                                        f"Extracted image: {new_filename} (ID: {rel_id})")

                except Exception as e:
                    logger.warning(f"Unable to parse image relationships, using fallback method: {e}")
                    # Fallback method: directly extract all images from media folder
                    for file_info in docx_zip.filelist:
                        if file_info.filename.startswith('word/media/'):
                            file_ext = os.path.splitext(
                                file_info.filename)[1].lower()
                            if file_ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg']:
                                self.image_counter += 1
                                new_filename = f"image_{self.image_counter:03d}{file_ext}"
                                output_path = os.path.join(
                                    self.assets_dir, new_filename)

                                with docx_zip.open(file_info.filename) as source:
                                    with open(output_path, 'wb') as target:
                                        shutil.copyfileobj(source, target)

                                logger.info(f"Extracted image: {new_filename}")

        except Exception as e:
            logger.warning(f"Error extracting images: {str(e)}")

    def _process_paragraph_images(self, paragraph: Paragraph) -> str:
        """Process images in paragraph"""
        images_found = []

        # Check if paragraph contains image elements
        para_element = paragraph._element

        # Method 1: Find w:drawing elements (new image format)
        drawings = para_element.xpath('.//w:drawing')
        logger.debug(f"Found {len(drawings)} drawing elements in paragraph")

        for drawing in drawings:
            # Find image relationship ID - using simplified method
            blip_elements = []
            # Iterate through all elements in drawing to find blip elements
            for elem in drawing.iter():
                if elem.tag.endswith('}blip'):
                    blip_elements.append(elem)

            for blip in blip_elements:
                rel_id = blip.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                logger.debug(f"Found image relationship ID: {rel_id}")

                if rel_id and rel_id in self.image_map:
                    # Found corresponding image file
                    image_filename = self.image_map[rel_id]
                    image_ref = f"![Image](./assets/{image_filename})"
                    images_found.append(image_ref)
                    logger.info(f"Inserted image link: {image_filename}")
                elif rel_id:
                    # Relationship ID exists but not in mapping, use generic method
                    if self.image_counter > 0:
                        # Use extracted images
                        image_filename = f"image_{len(images_found) + 1:03d}.png"
                        image_ref = f"![Image](./assets/{image_filename})"
                        images_found.append(image_ref)
                        logger.info(f"Using generic image link: {image_filename}")

            # If no blip elements found but have drawing, indicates there are images
            if not blip_elements and self.image_counter > 0:
                image_filename = f"image_{len(images_found) + 1:03d}.png"
                image_ref = f"![Image](./assets/{image_filename})"
                images_found.append(image_ref)
                logger.info(f"Using fallback image link: {image_filename}")

        # Method 2: Find w:pict elements (old image format)
        picts = para_element.xpath('.//w:pict')
        logger.debug(f"Found {len(picts)} pict elements in paragraph")

        for pict in picts:
            # Create reference for old image format
            if self.image_counter > 0:
                image_filename = f"image_{len(images_found) + 1:03d}.png"
                image_ref = f"![Image](./assets/{image_filename})"
                images_found.append(image_ref)
                logger.info(f"Inserted old image link: {image_filename}")

        # Method 3: Check images in runs
        for run in paragraph.runs:
            run_drawings = run._element.xpath('.//w:drawing')
            run_picts = run._element.xpath('.//w:pict')

            if run_drawings or run_picts:
                logger.debug(
                    f"Found image elements in run: drawings={len(run_drawings)}, picts={len(run_picts)}")

                # If no images found earlier but there are image elements here
                if not images_found and self.image_counter > 0:
                    image_filename = f"image_001.png"
                    image_ref = f"![Image](./assets/{image_filename})"
                    images_found.append(image_ref)
                    logger.info(f"Inserted image link in run: {image_filename}")

        if images_found:
            logger.info(f"Total {len(images_found)} images found in paragraph")

        return '\n'.join(images_found) if images_found else ""


def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description='Convert DOCX files to Markdown format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s input.docx                    # Output to stdout
  %(prog)s input.docx -o output.md       # Output to file
  %(prog)s *.docx -o output_dir/         # Batch conversion
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input DOCX file paths (supports wildcards)'
    )

    parser.add_argument(
        '-o', '--output',
        help='Output file or directory path'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Show verbose output'
    )

    args = parser.parse_args()

    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    converter = DocxToMarkdownConverter()

    try:
        for input_file in args.input_files:
            # Handle wildcards
            from glob import glob
            matching_files = glob(input_file)

            if not matching_files:
                logger.warning(f"No matching files found: {input_file}")
                continue

            for file_path in matching_files:
                if not file_path.lower().endswith('.docx'):
                    logger.warning(f"Skipping non-DOCX file: {file_path}")
                    continue

                # Determine output path
                output_path = None
                if args.output:
                    if os.path.isdir(args.output) or args.output.endswith('/'):
                        # Output to directory
                        base_name = Path(file_path).stem
                        output_path = os.path.join(
                            args.output, f"{base_name}.md")
                    else:
                        # Output to specified file
                        output_path = args.output

                # Execute conversion
                markdown_content = converter.convert_file(
                    file_path, output_path)

                # If no output file specified, print to stdout
                if not output_path:
                    print(f"\n=== {file_path} ===\n")
                    print(markdown_content)

    except KeyboardInterrupt:
        logger.info("Conversion interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Program execution failed: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
