"""
Command line interface for DOCX to Markdown converter.
"""

import argparse
import logging
import os
import sys
from glob import glob
from pathlib import Path

from .converter import DocxToMarkdownConverter

logger = logging.getLogger(__name__)


def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description='Convert DOCX files to Markdown format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s input.docx                    # Output to stdout
  %(prog)s input.docx -o output.md       # Output to file
    %(prog)s input.doc                     # Legacy .doc (requires LibreOffice)
  %(prog)s *.docx -o output_dir/         # Batch conversion
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input Word file paths (.docx or .doc; supports wildcards)'
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
    else:
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

    converter = DocxToMarkdownConverter()

    try:
        for input_file in args.input_files:
            # Handle wildcards
            matching_files = glob(input_file)

            if not matching_files:
                logger.warning(f"No matching files found: {input_file}")
                continue

            for file_path in matching_files:
                if not file_path.lower().endswith(('.docx', '.doc')):
                    logger.warning(f"Skipping non-Word file: {file_path}")
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
