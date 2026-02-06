#!/usr/bin/env python3
"""Bump project version in source files.

This script updates version strings in:
 - docx_converter/__init__.py (__version__)
 - setup.py (version=)

Usage:
    python scripts/bump_version.py 1.2.3

It prints the list of modified files and exits with:
 - 0 when changes were made and files updated
 - 2 when no changes required (already at requested version)
 - 1 on validation or other errors
"""
import re
import sys
from pathlib import Path

VERSION_RE = re.compile(r"^\d+\.\d+\.\d+(?:[\-+].*)?$")


def validate_version(v: str) -> bool:
    return bool(VERSION_RE.match(v.strip()))


def replace_in_file(path: Path, pattern: re.Pattern, repl: str) -> bool:
    text = path.read_text(encoding='utf-8')
    new_text, n = pattern.subn(repl, text)
    if n:
        path.write_text(new_text, encoding='utf-8')
        return True
    return False


def main():
    if len(sys.argv) < 2:
        print("Usage: bump_version.py <new-version>")
        return 1

    new_version = sys.argv[1].strip()
    if not validate_version(new_version):
        print(
            f"Invalid version: {new_version}. Expect semantic version like 1.2.3")
        return 1

    repo_root = Path(__file__).resolve().parents[1]

    files_changed = []

    # docx_converter/__init__.py
    init_py = repo_root / 'docx_converter' / '__init__.py'
    if init_py.exists():
        pattern = re.compile(r"__version__\s*=\s*['\"][^'\"]+['\"]")
        repl = f"__version__ = \"{new_version}\""
        if replace_in_file(init_py, pattern, repl):
            files_changed.append(str(init_py))

    # setup.py
    setup_py = repo_root / 'setup.py'
    if setup_py.exists():
        pattern = re.compile(r"version\s*=\s*['\"][^'\"]+['\"]")
        repl = f"version=\"{new_version}\""
        if replace_in_file(setup_py, pattern, repl):
            files_changed.append(str(setup_py))

    if not files_changed:
        print(f"No files needed updating; already at version {new_version}.")
        return 2

    print("Updated files:")
    for f in files_changed:
        print(" - ", f)

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
