"""
DOCX to Markdown Converter Package

A Python package for converting Microsoft Word documents (.docx) to Markdown format.
Also supports legacy .doc files by converting them to .docx (requires LibreOffice).
Supports conversion of text, headings, lists, tables, links and other basic formats.
"""

from .cli import main
from .converter import DocxToMarkdownConverter

__version__ = "1.0.3"
__author__ = "hnrobert"

__all__ = ['DocxToMarkdownConverter', 'main']
