#!/usr/bin/env python3
"""
Core converter module for DOCX to Markdown conversion.
"""

import logging
import os
import shutil
import zipfile
from pathlib import Path
from typing import Optional

from .document_processor import DocumentProcessor
from .image_extractor import ImageExtractor
from .utils import clean_markdown_content

try:
    from docx import Document
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)

logger = logging.getLogger(__name__)


class DocxToMarkdownConverter:
    """DOCX to Markdown converter class"""

    def __init__(self):
        self.output_lines = []
        self.output_folder = None
        self.assets_dir = None
        self.document_processor = None
        self.image_extractor = None

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
                raise FileNotFoundError(
                    f"Input file does not exist: {input_path}")

            # Setup output structure
            self._setup_output_structure(input_path, output_path)

            # Load DOCX document
            logger.info(f"Loading document: {input_path}")
            doc = Document(input_path)

            # Reset output
            self.output_lines = []

            # Initialize processors
            if self.assets_dir:
                self.image_extractor = ImageExtractor(self.assets_dir)
            else:
                # Fallback if assets_dir is None
                self.image_extractor = ImageExtractor("")

            self.document_processor = DocumentProcessor(
                self.image_extractor,
                self.output_lines
            )

            # Extract images first
            if self.image_extractor and self.assets_dir:
                self.image_extractor.extract_images(input_path)

            # Convert document content
            self.document_processor.convert_document(doc)

            # Generate and clean Markdown content
            markdown_content = clean_markdown_content(self.output_lines)

            # Write to file
            final_output_path = self._get_final_output_path(
                input_path, output_path)
            self._write_output(markdown_content, final_output_path)

            # Clean up empty assets directory
            self._cleanup_empty_assets_dir()

            logger.info(
                f"Conversion completed, output file: {final_output_path}")

            return markdown_content

        except Exception as e:
            logger.error(f"Error occurred during conversion: {str(e)}")
            raise

    def _setup_output_structure(self, input_path: str, output_path: Optional[str]):
        """Setup output folder structure"""
        input_stem = Path(input_path).stem

        if output_path:
            if os.path.isdir(output_path) or output_path.endswith('/'):
                self.output_folder = os.path.join(output_path, input_stem)
            else:
                self.output_folder = os.path.dirname(output_path)
                if not self.output_folder:
                    self.output_folder = input_stem
        else:
            self.output_folder = input_stem

        # Create output folder and assets folder
        os.makedirs(self.output_folder, exist_ok=True)
        self.assets_dir = os.path.join(self.output_folder, "assets")
        os.makedirs(self.assets_dir, exist_ok=True)

    def _get_final_output_path(self, input_path: str, output_path: Optional[str]) -> str:
        """Get the final output file path"""
        input_stem = Path(input_path).stem

        if output_path:
            if os.path.isdir(output_path) or output_path.endswith('/'):
                if self.output_folder:
                    return os.path.join(self.output_folder, f"{input_stem}.md")
            else:
                return output_path

        if self.output_folder:
            return os.path.join(self.output_folder, f"{input_stem}.md")

        return f"{input_stem}.md"

    def _write_output(self, content: str, output_path: str):
        """Write output file"""
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def _cleanup_empty_assets_dir(self):
        """Remove assets directory if it's empty"""
        if self.assets_dir and os.path.exists(self.assets_dir):
            try:
                # Check if assets directory is empty
                if not os.listdir(self.assets_dir):
                    os.rmdir(self.assets_dir)
                    logger.debug(
                        f"Removed empty assets directory: {self.assets_dir}")
                else:
                    logger.debug(
                        f"Assets directory not empty, keeping: {self.assets_dir}")
            except OSError as e:
                logger.debug(f"Could not remove assets directory: {e}")
