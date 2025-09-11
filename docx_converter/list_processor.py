"""
List processing module for handling ordered and unordered lists.
"""

from typing import Dict, List, Optional

from .utils import is_list_marker_text, is_numbered_list_text, remove_list_markers

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class ListProcessor:
    """Handles list processing and conversion"""

    def __init__(self, output_lines: List[str], text_formatter):
        self.output_lines = output_lines
        self.text_formatter = text_formatter
        self.current_list_level = 0
        self.list_counters: Dict[int, int] = {}
        self.in_list = False
        self.list_type: Optional[str] = None

    def is_list_paragraph(self, paragraph: Paragraph) -> bool:
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
        if is_list_marker_text(text):
            return True

        # Check if it's a numbered list
        if is_numbered_list_text(text):
            return True

        return False

    def _get_list_level(self, paragraph: Paragraph) -> int:
        """Determine the list level (indentation depth) of a paragraph"""
        try:
            # Check for numbering format in paragraph properties
            if paragraph._element.pPr is not None:
                numPr = paragraph._element.pPr.find(
                    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
                if numPr is not None:
                    # Try to get the list level from numbering properties
                    ilvl = numPr.find(
                        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
                    if ilvl is not None:
                        level = ilvl.get(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        if level is not None:
                            return int(level)

                # Check indentation from paragraph properties
                ind = paragraph._element.pPr.find(
                    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
                if ind is not None:
                    left_val = ind.get(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left')
                    hanging_val = ind.get(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging')

                    if left_val:
                        # Convert twips to approximate indentation level
                        # 1 inch = 1440 twips, typical indent is about 0.5 inch per level
                        indent_twips = int(left_val)
                        # 720 twips ≈ 0.5 inch
                        level = max(0, indent_twips // 720)
                        return min(level, 5)  # Cap at reasonable level

            # Fallback: analyze text for visual markers
            text = paragraph.text.strip()
            if text.startswith('o\t') or text.startswith('o '):
                return 1  # Sub-item
            elif text.startswith('▪') or text.startswith('◦'):
                return 1
            elif text.startswith('\t'):
                return text.count('\t')

        except Exception:
            pass

        return 0  # Default to top level

    def convert_list_item(self, paragraph: Paragraph) -> None:
        """Convert list item"""
        text = paragraph.text.strip()

        # Check paragraph style
        style_name = paragraph.style.name.lower(
        ) if paragraph.style and paragraph.style.name else ''

        # Detect list level from indentation or numbering format
        list_level = self._get_list_level(paragraph)

        # Determine list type
        is_ordered = self._determine_list_type(text, style_name)

        # Remove list markers from text
        cleaned_text = remove_list_markers(text)

        # Detect list type change or start new list
        current_list_type = 'ordered' if is_ordered else 'unordered'

        if not self.in_list or self.list_type != current_list_type:
            # Start new list or list type changed
            self.in_list = True
            self.list_type = current_list_type
            self.current_list_level = list_level
            if is_ordered:
                self.list_counters[list_level] = 1
        else:
            # Update current level
            self.current_list_level = list_level

        # Generate list marker
        if is_ordered:
            # Use counter to generate correct sequence number
            counter = self.list_counters.get(list_level, 1)
            list_marker = f"{counter}."
            self.list_counters[list_level] = counter + 1
        else:
            list_marker = '-'

        # Add appropriate indentation based on list level
        indent = '  ' * list_level
        formatted_text = self.text_formatter.convert_paragraph_formatting(
            paragraph, cleaned_text)
        self.output_lines.append(f"{indent}{list_marker} {formatted_text}")

    def end_list(self) -> None:
        """End current list"""
        self.in_list = False
        self.list_type = None

    def _determine_list_type(self, text: str, style_name: str) -> bool:
        """Determine if list is ordered or unordered"""
        # Check if it's an ordered list (starts with number)
        if is_numbered_list_text(text):
            return True
        elif ('list' in style_name or 'bullet' in style_name) and not is_list_marker_text(text):
            # If it's a list style but no obvious markers, check if it's ordered list style
            if 'number' in style_name or 'ordered' in style_name:
                return True

        return False
