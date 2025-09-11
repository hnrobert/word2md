"""
Document processing module for handling main document conversion.
"""

from typing import Any, List, Dict

from .paragraph_processor import ParagraphProcessor
from .table_processor import TableProcessor
from .utils import find_font_size_based_headings, clean_markdown_content

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class DocumentProcessor:
    """Handles main document processing and coordination"""

    def __init__(self, image_extractor, output_lines: List[str]):
        self.output_lines = output_lines
        self.paragraph_processor = ParagraphProcessor(
            image_extractor, output_lines)
        self.table_processor = TableProcessor(output_lines)
        self.font_size_headings: Dict[float, int] = {}

    def convert_document(self, doc: Any) -> None:
        """Convert main document content"""
        # First check if there are Title style paragraphs, if so use as main title
        title_found = self._check_for_title_style(doc)

        # Check if there are any heading styles in the document
        heading_styles_found = self._check_for_heading_styles(doc)

        # If no heading styles found, analyze font sizes to create heading hierarchy
        if not heading_styles_found:
            self.font_size_headings = find_font_size_based_headings(doc)

        # Set heading offset: if Title style exists, all headings are adjusted down one level
        heading_offset = 1 if title_found else 0
        self.paragraph_processor.set_heading_offset(heading_offset)
        self.paragraph_processor.set_font_size_headings(
            self.font_size_headings)

        # Process all document elements
        first_heading_found = False
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

                self.paragraph_processor.convert_paragraph(paragraph)

            elif element.tag.endswith('tbl'):  # Table
                table = Table(element, doc)
                self.table_processor.convert_table(table)

        # Post-process to fix heading levels and punctuation
        self._fix_heading_levels()

    def _fix_heading_levels(self) -> None:
        """Fix heading level jumps and remove punctuation from headings"""
        import re

        lines = self.output_lines[:]
        self.output_lines.clear()

        last_heading_level = 0

        for line in lines:
            # Check if this is a heading line
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)

            if heading_match:
                current_hashes = heading_match.group(1)
                heading_text = heading_match.group(2)
                current_level = len(current_hashes)

                # Fix heading level jumps (MD001)
                if last_heading_level > 0:  # Not the first heading
                    max_allowed_level = last_heading_level + 1
                    if current_level > max_allowed_level:
                        # Reduce level to avoid jumping
                        current_level = max_allowed_level
                        current_hashes = '#' * current_level

                # Clean heading text - remove trailing punctuation
                clean_heading_text = self._clean_heading_text(heading_text)

                # Update the line with fixed level and clean text
                fixed_line = f"{current_hashes} {clean_heading_text}"
                self.output_lines.append(fixed_line)

                last_heading_level = current_level
            else:
                # Not a heading, keep as is
                self.output_lines.append(line)

    def _clean_heading_text(self, text: str) -> str:
        """Remove trailing punctuation from heading text"""
        import re

        # Remove trailing punctuation like 。！？：；，
        text = re.sub(r'[。！？：；，]+$', '', text.strip())

        # Also remove trailing colons and periods in English
        text = re.sub(r'[:\.]+$', '', text.strip())

        return text.strip()

    def _check_for_title_style(self, doc: Any) -> bool:
        """Check if document contains Title style paragraphs"""
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                if 'title' in style_name and paragraph.text.strip():
                    return True
        return False

    def _check_for_heading_styles(self, doc: Any) -> bool:
        """Check if document contains any Heading style paragraphs"""
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                if 'heading' in style_name and paragraph.text.strip():
                    return True
        return False
