"""
Text formatting module for converting Word formatting to Markdown.
"""

from typing import Optional

from .utils import merge_adjacent_tags

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class TextFormatter:
    """Handles text formatting conversion from Word to Markdown"""

    def convert_paragraph_formatting(self, paragraph: Paragraph, custom_text: Optional[str] = None) -> str:
        """
        Convert paragraph formatting (bold, italic, links, etc.)

        Args:
            paragraph: Word paragraph object
            custom_text: Custom text to use instead of paragraph runs

        Returns:
            Formatted Markdown text
        """
        if custom_text:
            # If custom text is provided, use simplified processing
            return custom_text

        # First check if the entire paragraph is a hyperlink
        hyperlink_result = self._process_paragraph_hyperlinks(paragraph)
        if hyperlink_result:
            return hyperlink_result

        result = []
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            # Check if run contains hyperlink
            hyperlink = self._get_hyperlink(run, paragraph)

            # Apply formatting
            if run.bold:
                text = f"**{text}**"
            if run.italic:
                text = f"*{text}*"
            if run.underline:
                text = f"<u>{text}</u>"

            # Apply hyperlink formatting
            if hyperlink:
                text = f"[{text}]({hyperlink})"

            result.append(text)        # Merge adjacent same HTML tags
        final_result = ''.join(result)

        # Merge adjacent tags of same type
        final_result = merge_adjacent_tags(final_result)

        return final_result

    def _get_hyperlink(self, run, paragraph) -> Optional[str]:
        """Extract hyperlink URL from run"""
        try:
            # Check if the run is part of a hyperlink
            element = run._element

            # Check the run's parent paragraph for hyperlink elements
            para_element = paragraph._element
            for child in para_element.iter():
                if child.tag.endswith('hyperlink'):
                    # Check if this hyperlink contains our run
                    for run_elem in child.iter():
                        if run_elem == element or element in list(run_elem.iter()):
                            # Get the relationship ID
                            rel_id = child.get(
                                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                            if rel_id:
                                try:
                                    part = paragraph.part
                                    rel = part.rels[rel_id]
                                    return rel.target_ref
                                except (KeyError, AttributeError):
                                    pass
                            break

            # Also check if the run element itself is within a hyperlink by traversing up
            current = element
            while current is not None:
                if current.tag.endswith('hyperlink'):
                    rel_id = current.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rel_id:
                        try:
                            part = paragraph.part
                            rel = part.rels[rel_id]
                            return rel.target_ref
                        except (KeyError, AttributeError):
                            pass
                    break
                current = current.getparent()

        except Exception:
            # If any error occurs during hyperlink extraction, return None
            pass

        return None

    def _process_paragraph_hyperlinks(self, paragraph) -> Optional[str]:
        """Process paragraph-level hyperlinks that contain the entire paragraph text"""
        try:
            para_element = paragraph._element

            # Look for hyperlink elements in the paragraph
            for child in para_element.iter():
                if child.tag.endswith('hyperlink'):
                    # Get the relationship ID
                    rel_id = child.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rel_id:
                        try:
                            part = paragraph.part
                            rel = part.rels[rel_id]
                            url = rel.target_ref

                            # Extract all text from the hyperlink element
                            hyperlink_text = ""
                            for text_elem in child.iter():
                                if text_elem.tag.endswith('t') and text_elem.text:
                                    hyperlink_text += text_elem.text

                            if hyperlink_text.strip():
                                return f"[{hyperlink_text.strip()}]({url})"
                        except (KeyError, AttributeError):
                            pass

        except Exception:
            pass

        return None
