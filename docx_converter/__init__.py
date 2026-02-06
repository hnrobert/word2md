"""
DOCX to Markdown Converter Package

A Python package for converting Microsoft Word documents (.docx) to Markdown format.
Also supports legacy .doc files by converting them to .docx (requires LibreOffice).
Supports conversion of text, headings, lists, tables, links and other basic formats.
"""

from .converter import DocxToMarkdownConverter
from .cli import main

__version__ = "1.0.0"
__author__ = "HNRobert"

__all__ = ['DocxToMarkdownConverter', 'main']
