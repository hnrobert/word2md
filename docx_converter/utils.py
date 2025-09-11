"""
Utility functions for DOCX to Markdown conversion.
"""

import re
from typing import List, Dict, Set, Tuple, Optional

try:
    from docx.text.paragraph import Paragraph
    from docx.shared import Pt
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


def clean_markdown_content(output_lines: List[str]) -> str:
    """
    Clean and format Markdown content

    Args:
        output_lines: List of output lines

    Returns:
        Cleaned Markdown content string
    """
    # Generate Markdown content
    markdown_content = '\n'.join(output_lines)

    # Clean up extra blank lines - merge multiple consecutive blank lines into single blank line
    markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)

    # Remove blank lines at beginning and end
    markdown_content = markdown_content.strip()

    # Add extra blank line at the end
    markdown_content += '\n'

    return markdown_content


def extract_heading_level(style_name: str) -> int:
    """Extract heading level from style name"""
    match = re.search(r'heading\s*(\d+)', style_name)
    if match:
        # Markdown supports maximum 6 heading levels
        return min(int(match.group(1)), 6)
    return 1


def merge_adjacent_tags(text: str) -> str:
    """Merge adjacent HTML tags of the same type"""
    # Merge adjacent underline tags
    text = re.sub(r'</u><u>', '', text)
    return text


def is_list_marker_text(text: str) -> bool:
    """Check if text starts with list markers"""
    list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
    return any(text.startswith(marker + ' ') for marker in list_markers)


def is_numbered_list_text(text: str) -> bool:
    """Check if text is a numbered list"""
    return bool(re.match(r'^\d+[\.）]\s+', text))


def remove_list_markers(text: str) -> str:
    """Remove list markers from text"""
    # Remove numbered list markers
    text = re.sub(r'^\d+[\.）]\s+', '', text)

    # Remove unordered list markers
    list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
    for marker in list_markers:
        if text.startswith(marker + ' '):
            text = text[len(marker):].strip()
    return text


def get_paragraph_font_size(paragraph: Paragraph) -> Optional[float]:
    """
    Get the font size of a paragraph in points.
    Returns the most common font size in the paragraph, or None if not found.
    """
    font_sizes = []

    for run in paragraph.runs:
        if run.font.size:
            # Convert from Twips to Points (1 point = 20 twips)
            size_in_points = run.font.size.pt
            font_sizes.append(size_in_points)

    if not font_sizes:
        return None

    # Return the most common font size
    from collections import Counter
    return Counter(font_sizes).most_common(1)[0][0]


def is_paragraph_uniform_font_size(paragraph: Paragraph) -> bool:
    """
    Check if all text in the paragraph has the same font size.
    """
    font_sizes = set()

    for run in paragraph.runs:
        if run.font.size and run.text.strip():  # Only consider runs with text
            size_in_points = run.font.size.pt
            font_sizes.add(size_in_points)

    # Paragraph is uniform if it has at most one font size
    return len(font_sizes) <= 1


def analyze_font_size_hierarchy(paragraphs_with_sizes: List[Tuple[Paragraph, float]]) -> Dict[float, int]:
    """
    Analyze font sizes and assign heading levels based on size hierarchy.

    Args:
        paragraphs_with_sizes: List of (paragraph, font_size) tuples

    Returns:
        Dictionary mapping font_size to heading_level (1-6, or 0 for normal text)
    """
    if not paragraphs_with_sizes:
        return {}

    # Get unique font sizes, sorted in descending order (largest first)
    unique_sizes = sorted(
        set(size for _, size in paragraphs_with_sizes), reverse=True)

    # If only one size, it's probably normal text
    if len(unique_sizes) == 1:
        return {unique_sizes[0]: 0}

    # Determine the baseline size (most common size, likely normal text)
    from collections import Counter
    size_counts = Counter(size for _, size in paragraphs_with_sizes)
    baseline_size = size_counts.most_common(1)[0][0]

    # Assign heading levels to sizes larger than baseline
    size_to_level = {}
    heading_level = 1

    for size in unique_sizes:
        if size > baseline_size and heading_level <= 6:
            size_to_level[size] = heading_level
            heading_level += 1
        else:
            size_to_level[size] = 0  # Normal text

    return size_to_level


def find_font_size_based_headings(doc) -> Dict[float, int]:
    """
    Analyze the entire document to find potential headings based on font size.
    Returns a mapping of font_size -> heading_level.
    """
    candidates = []

    # First pass: collect all paragraphs with uniform font sizes
    for element in doc.element.body:
        if element.tag.endswith('p'):  # Paragraph
            paragraph = Paragraph(element, doc)
            text = paragraph.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Skip paragraphs with existing heading styles
            style_name = paragraph.style.name.lower(
            ) if paragraph.style and paragraph.style.name else ''
            if 'heading' in style_name or 'title' in style_name:
                continue

            # Check if paragraph has uniform font size
            if is_paragraph_uniform_font_size(paragraph):
                font_size = get_paragraph_font_size(paragraph)
                if font_size:
                    candidates.append((paragraph, font_size))

    # Analyze and assign heading levels
    return analyze_font_size_hierarchy(candidates)
