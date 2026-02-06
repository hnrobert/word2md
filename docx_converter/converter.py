#!/usr/bin/env python3
"""
Core converter module for DOCX to Markdown conversion.
"""

import logging
import os
import shutil
import subprocess
import tempfile
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
        temp_dir: Optional[str] = None
        temp_docx_path: Optional[str] = None

        try:
            # Check if input file exists
            if not os.path.exists(input_path):
                raise FileNotFoundError(
                    f"Input file does not exist: {input_path}")

            # Setup output structure
            self._setup_output_structure(input_path, output_path)

            # Convert legacy .doc to a temporary .docx (python-docx can't open .doc)
            effective_input_path = input_path
            if input_path.lower().endswith('.doc'):
                temp_dir = tempfile.mkdtemp(prefix='docx2md_', suffix='_docx')
                temp_docx_path = self._convert_doc_to_docx(
                    input_path, temp_dir)
                effective_input_path = temp_docx_path

            # Load DOCX document
            logger.info(f"Loading document: {effective_input_path}")
            doc = Document(effective_input_path)

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
                self.image_extractor.extract_images(effective_input_path)

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
        finally:
            # Clean up temporary conversion artifacts
            if temp_docx_path:
                try:
                    os.remove(temp_docx_path)
                except OSError:
                    pass
            if temp_dir:
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except OSError:
                    pass

    def _convert_doc_to_docx(self, input_doc_path: str, out_dir: str) -> str:
        """Convert a legacy .doc file to .docx using LibreOffice/soffice.

        Returns the converted .docx path.
        """
        soffice_path = self._find_soffice_executable()

        # LibreOffice writes the output docx into out_dir, keeping the base name.
        cmd = [
            soffice_path,
            '--headless',
            '--nologo',
            '--nofirststartwizard',
            '--convert-to',
            'docx',
            '--outdir',
            out_dir,
            input_doc_path,
        ]

        logger.info(
            f"Converting .doc to .docx via LibreOffice: {input_doc_path}")
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE)
        except FileNotFoundError as e:
            raise RuntimeError(
                "LibreOffice (soffice) not found. Install LibreOffice to convert .doc files. "
                "On macOS: brew install --cask libreoffice"
            ) from e
        except subprocess.CalledProcessError as e:
            stderr = (e.stderr or b'').decode('utf-8', errors='replace')
            raise RuntimeError(
                f"Failed to convert .doc to .docx using LibreOffice. Details: {stderr.strip()}"
            ) from e

        expected = os.path.join(out_dir, f"{Path(input_doc_path).stem}.docx")
        if os.path.exists(expected):
            return expected

        # Fallback: find any produced .docx
        candidates = [p for p in Path(out_dir).glob('*.docx') if p.is_file()]
        if len(candidates) == 1:
            return str(candidates[0])
        if candidates:
            # If multiple, pick the newest
            newest = max(candidates, key=lambda p: p.stat().st_mtime)
            return str(newest)

        raise RuntimeError(
            "LibreOffice reported success but no .docx was produced.")

    def _find_soffice_executable(self) -> str:
        """Locate LibreOffice command for headless conversion across OSes.

        Order of checks:
        1. Environment variable `DOCX2MD_SOFFICE_PATH`
        2. Common executable names on PATH: `soffice`, `libreoffice`
        3. Known installation paths per-platform (macOS, Windows, common Linux locations)
        4. Flatpak/exported paths
        Raises a RuntimeError with helpful instructions if not found.
        """
        import platform

        # 1) Allow explicit override via environment variable
        env_path = os.environ.get(
            'DOCX2MD_SOFFICE_PATH') or os.environ.get('SOFFICE_PATH')
        if env_path:
            if os.path.exists(env_path) and os.access(env_path, os.X_OK):
                return env_path
            else:
                raise RuntimeError(
                    f"DOCX2MD_SOFFICE_PATH is set to '{env_path}' but file is not executable or doesn't exist.")

        # 2) Common executable names on PATH
        for name in ('soffice', 'libreoffice'):
            found = shutil.which(name)
            if found:
                return found

        system = platform.system()

        # 3) Platform-specific likely locations
        candidates = []

        if system == 'Darwin':  # macOS
            candidates += [
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/soffice.bin',
                '/usr/local/bin/soffice',
                '/opt/homebrew/bin/soffice',
                '/opt/local/bin/soffice',
            ]
        elif system == 'Windows':  # Windows
            program_files = os.environ.get('ProgramFiles', r'C:\Program Files')
            program_files_x86 = os.environ.get(
                'ProgramFiles(x86)', r'C:\Program Files (x86)')
            candidates += [
                os.path.join(program_files, 'LibreOffice',
                             'program', 'soffice.exe'),
                os.path.join(program_files_x86, 'LibreOffice',
                             'program', 'soffice.exe'),
            ]
        else:  # Linux / other Unix
            candidates += [
                '/usr/bin/libreoffice',
                '/usr/bin/soffice',
                '/usr/local/bin/libreoffice',
                '/usr/local/bin/soffice',
                '/snap/bin/libreoffice',
            ]

        # 4) Check flatpak / exported locations
        candidates += [
            '/var/lib/flatpak/exports/bin/libreoffice',
            '/var/lib/flatpak/exports/bin/soffice',
        ]

        # Verify candidate paths
        for p in candidates:
            if p and os.path.exists(p) and os.access(p, os.X_OK):
                return p

        # Nothing found — provide a helpful error message
        hint_lines = [
            'LibreOffice (soffice) not found on this system.',
            'To enable .doc support the converter needs LibreOffice for .doc → .docx conversion.',
            'Options:',
            "  * Install LibreOffice and ensure `soffice` is on PATH (macOS: `brew install --cask libreoffice`).",
            "  * Set the environment variable `DOCX2MD_SOFFICE_PATH` to the soffice executable path.",
        ]
        raise RuntimeError('\n'.join(hint_lines))

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
