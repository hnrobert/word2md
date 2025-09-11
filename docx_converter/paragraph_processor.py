"""
Paragraph processing module for DOCX to Markdown conversion.
"""

import logging
from typing import Dict, List

from .formatting import TextFormatter
from .image_processor import ImageProcessor
from .list_processor import ListProcessor
from .utils import extract_heading_level, get_paragraph_font_size

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)

logger = logging.getLogger(__name__)


class ParagraphProcessor:
    """Handles paragraph processing and conversion"""

    def __init__(self, image_extractor, output_lines: List[str]):
        self.output_lines = output_lines
        self.text_formatter = TextFormatter()
        self.image_processor = ImageProcessor(image_extractor)
        self.list_processor = ListProcessor(output_lines, self.text_formatter)
        self.heading_offset = 0
        self.font_size_headings: Dict[float, int] = {}

    def set_heading_offset(self, offset: int):
        """Set heading level offset"""
        self.heading_offset = offset

    def set_font_size_headings(self, font_size_headings: Dict[float, int]):
        """Set font size to heading level mapping"""
        self.font_size_headings = font_size_headings

    def convert_paragraph(self, paragraph: Paragraph) -> None:
        """Convert paragraph to Markdown"""
        # Get paragraph text
        text = paragraph.text.strip()

        # First check if paragraph contains images (regardless of text content)
        images_text = self.image_processor.process_paragraph_images(paragraph)

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

        # Skip Title style, already handled in document processor
        if 'title' in style_name:
            return

        # Check if it's a list item (but exclude chapter/section numbers)
        is_list = self.list_processor.is_list_paragraph(
            paragraph) and not self._is_section_number(text)

        # If previously in list but current is not list item, list ends
        if self.list_processor.in_list and not is_list:
            # Add blank line to separate list and subsequent content
            self.output_lines.append('')
            self.list_processor.end_list()

        # Handle headings (adjust level based on Title style presence)
        if 'heading' in style_name:
            self._convert_heading(paragraph, text, style_name)
            return

        # Check if paragraph should be treated as heading based on font size
        if self.font_size_headings and self._is_font_size_heading(paragraph):
            self._convert_font_size_heading(paragraph, text)
            return

        # Check if paragraph should be treated as heading based on formatting (bold text)
        if self._is_formatted_heading(paragraph, text):
            self._convert_formatted_heading(paragraph, text)
            return

        # Check if it's a section number that should be treated as heading
        if self._is_section_number(text):
            self._convert_section_number_heading(paragraph, text)
            return

        # Handle lists
        if is_list:
            self.list_processor.convert_list_item(paragraph)
            return

        # Handle regular paragraphs
        # If paragraph contains images, insert images first
        if images_text:
            self.output_lines.append(images_text)
            self.output_lines.append('')

        # Handle text content
        if text:  # Only process when paragraph has text
            markdown_text = self.text_formatter.convert_paragraph_formatting(
                paragraph)
            self.output_lines.append(markdown_text)
            self.output_lines.append('')

    def _convert_heading(self, paragraph: Paragraph, text: str, style_name: str) -> None:
        """Convert heading paragraph"""
        level = extract_heading_level(style_name)

        # If Title style exists, all headings are adjusted down one level
        level += self.heading_offset

        # Ensure not exceeding 6 heading levels
        level = min(level, 6)

        self.output_lines.append(f"{'#' * level} {text}")
        self.output_lines.append('')

    def _is_font_size_heading(self, paragraph: Paragraph) -> bool:
        """Check if paragraph should be treated as heading based on font size"""
        font_size = get_paragraph_font_size(paragraph)
        if font_size is None:
            return False

        # Check if this font size is mapped to a heading level (non-zero)
        return self.font_size_headings.get(font_size, 0) > 0

    def _convert_font_size_heading(self, paragraph: Paragraph, text: str) -> None:
        """Convert paragraph to heading based on font size"""
        font_size = get_paragraph_font_size(paragraph)
        if font_size is None:
            return

        # Get heading level from font size mapping
        level = self.font_size_headings.get(font_size, 1)

        # Apply heading offset if needed
        level += self.heading_offset

        # Ensure not exceeding 6 heading levels
        level = min(level, 6)

        self.output_lines.append(f"{'#' * level} {text}")
        self.output_lines.append('')

    def _is_section_number(self, text: str) -> bool:
        """Check if text is a section/chapter number rather than a list item"""
        # These typically have more substantial content after the number
        import re

        # Look for numbered items with substantial content that seem like section headers
        if re.match(r'^\d+\.\s+[^，。！？；：]*[入门|介绍|概述|基础|原理|设计|分析|方法|系统|结构|材料|工艺]', text):
            return True

        # Only identify as section titles if they contain meaningful section keywords
        # This prevents ordinary list items like "1. Object 1" from being treated as headings
        section_keywords = ['入门', '介绍', '概述', '基础', '原理', '设计', '分析', '方法', '系统', '结构', '材料', '工艺',
                            '课程', '培训', '学习', '知识', '技能', '理论', '实践', '应用']
        if re.match(r'^\d+\.\s+', text):
            # Check if the text contains section-related keywords
            if any(keyword in text for keyword in section_keywords):
                return True

        return False

    def _is_formatted_heading(self, paragraph: Paragraph, text: str) -> bool:
        """Check if paragraph should be treated as heading based on formatting"""
        # Skip empty text
        if not text.strip():
            return False

        # Skip if it's already identified as a list
        if self.list_processor.is_list_paragraph(paragraph):
            return False

        # Check if entire paragraph is bold (indicating it might be a heading)
        has_bold_text = False
        total_text_length = 0
        bold_text_length = 0

        for run in paragraph.runs:
            if run.text.strip():
                total_text_length += len(run.text.strip())
                if run.bold:
                    bold_text_length += len(run.text.strip())
                    has_bold_text = True

        # If most of the text is bold, consider it a heading
        if has_bold_text and total_text_length > 0:
            bold_ratio = bold_text_length / total_text_length
            if bold_ratio >= 0.8:  # At least 80% of text is bold
                # Check for heading-like patterns
                return self._looks_like_heading(text)

        return False

    def _looks_like_heading(self, text: str) -> bool:
        """Check if text looks like a heading"""
        import re

        # Patterns that look like headings (but only for bold/formatted text)
        heading_patterns = [
            r'^[一二三四五六七八九十]+、\s*',  # Chinese numbers like "一、"
            r'^[第]\d+[章节部分]\s*',  # Like "第1章"
            r'^[课程|培训|内容|说明|工具|资源|考核]',  # Common heading words at start
        ]

        for pattern in heading_patterns:
            if re.match(pattern, text):
                return True

        # Check for keyword-starting text (strong indicators of headings)
        heading_starters = ['最终考核：', '软件工具', '在线资源', '具体内容', '培训课程', '核心知识']
        if any(text.startswith(starter) for starter in heading_starters):
            return True

        # Also check for short, descriptive text (likely headings)
        if len(text.strip()) <= 100 and not text.endswith('。'):
            # Check for keyword patterns indicating headings
            heading_keywords = ['入门', '基础', '课程', '培训',
                                '工具', '软件', '资源', '概述', '介绍', '说明', '内容', '考核']
            if any(keyword in text for keyword in heading_keywords):
                return True

        return False

    def _convert_formatted_heading(self, paragraph: Paragraph, text: str) -> None:
        """Convert formatted paragraph to heading"""
        import re

        # Determine heading level based on text pattern
        # Default level for formatted headings (bold text without specific patterns)
        level = 3

        # Chinese section numbers (一、二、三、) - main sections
        if re.match(r'^[一二三四五六七八九十]+、', text):
            level = 2
        # For other formatted headings, use consistent level based on formatting only
        # All bold text without specific numbering patterns gets the same level

        # Apply heading offset if needed
        level += self.heading_offset

        # Ensure not exceeding 6 heading levels
        level = min(level, 6)

        # Remove bold formatting since we're converting to markdown heading
        clean_text = text.strip()

        self.output_lines.append(f"{'#' * level} {clean_text}")
        self.output_lines.append('')

    def _convert_section_number_heading(self, paragraph: Paragraph, text: str) -> None:
        """Convert section number paragraph to heading"""
        import re

        # Determine heading level based on section number pattern
        level = 3  # Default level for numbered sections like "1. 基础力学入门"

        # Apply heading offset if needed
        level += self.heading_offset

        # Ensure not exceeding 6 heading levels
        level = min(level, 6)

        # Clean text
        clean_text = text.strip()

        self.output_lines.append(f"{'#' * level} {clean_text}")
        self.output_lines.append('')
