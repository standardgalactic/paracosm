#!/usr/bin/env python3
"""
A robust script to extract plain text from MHTML, DOCX, Markdown (.md), and HTML files.
It can process a single file or all supported files recursively in the current directory and subdirectories.
- For MHTML: parses as a MIME message, extracts text from HTML (or plain text) parts, removes HTML tags using BeautifulSoup.
- For DOCX: extracts plain text from paragraphs using python-docx.
- For Markdown: reads the file, strips Markdown formatting to plain text using markdown and BeautifulSoup.
- For HTML: extracts plain text using BeautifulSoup.
Writes extracted text to .txt files.

Usage:
    python extract_text.py [file]  # Process a specific file (.mhtml, .docx, .md, .html)
    python extract_text.py         # Process all supported files in directory & subdirectories

Dependencies:
    pip install beautifulsoup4 python-docx markdown

"""

import os
import sys
import argparse
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup

def extract_text_from_mhtml(file_path):
    """Extract plain text from an MHTML file by parsing its MIME structure."""
    try:
        with open(file_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
    except Exception as e:
        print(f"Failed to parse MHTML file {file_path!r}: {e}")
        return ""

    text_parts = []
    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/html':
                    html_content = part.get_content()
                    soup = BeautifulSoup(html_content, 'html.parser')
                    text_parts.append(soup.get_text(separator="\n", strip=True))
                elif content_type == 'text/plain':
                    text_parts.append(part.get_content())
        else:
            content_type = msg.get_content_type()
            if content_type == 'text/html':
                soup = BeautifulSoup(msg.get_content(), 'html.parser')
                text_parts.append(soup.get_text(separator="\n", strip=True))
            else:
                text_parts.append(msg.get_content())
        return "\n\n".join(text_parts)
    except Exception as e:
        print(f"Failed to extract from MHTML file {file_path!r}: {e}")
        return ""

def extract_text_from_docx(file_path):
    """Extract all text from a DOCX file."""
    try:
        from docx import Document
    except ImportError:
        print("Missing dependency: python-docx. Install with 'pip install python-docx'")
        sys.exit(1)
    try:
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        return text
    except Exception as e:
        print(f"Failed to read DOCX file {file_path!r}: {e}")
        return ""

def extract_text_from_md(file_path):
    """Extract plain text from a Markdown file by stripping formatting."""
    try:
        import markdown
    except ImportError:
        print("Missing dependency: markdown. Install with 'pip install markdown'")
        sys.exit(1)
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            md_content = f.read()
        html = markdown.markdown(md_content)
        soup = BeautifulSoup(html, 'html.parser')
        return soup.get_text(separator="\n", strip=True)
    except Exception as e:
        print(f"Failed to process Markdown file {file_path!r}: {e}")
        return ""

def extract_text_from_html(file_path):
    """Extract plain text from an HTML file using BeautifulSoup."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            html_content = f.read()
        soup = BeautifulSoup(html_content, 'html.parser')
        return soup.get_text(separator="\n", strip=True)
    except Exception as e:
        print(f"Failed to process HTML file {file_path!r}: {e}")
        return ""

def process_file(file_path):
    """Process a single file and write the extracted text to a .txt file, only if it doesn't yet exist."""
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        return
    ext = os.path.splitext(file_path)[1].lower()
    output_file = os.path.splitext(file_path)[0] + ".txt"
    if os.path.exists(output_file):
        print(f"Skipping {file_path} (output {output_file} already exists).")
        return

    if ext == '.mhtml':
        print(f'Processing MHTML: {file_path} ...')
        extracted_text = extract_text_from_mhtml(file_path)
    elif ext == '.docx':
        print(f'Processing DOCX: {file_path} ...')
        extracted_text = extract_text_from_docx(file_path)
    elif ext == '.md':
        print(f'Processing Markdown: {file_path} ...')
        extracted_text = extract_text_from_md(file_path)
    elif ext == '.html':
        print(f'Processing HTML: {file_path} ...')
        extracted_text = extract_text_from_html(file_path)
    else:
        print(f"Error: '{file_path}' is not a supported file type (.mhtml, .docx, .md, .html).")
        return

    try:
        with open(output_file, "w", encoding="utf-8", errors="replace") as out:
            out.write(extracted_text)
        print(f"Extracted text written to {output_file}")
    except Exception as e:
        print(f"Failed to write output file {output_file!r}: {e}")

def find_supported_files(root_dir, supported_exts):
    """Recursively find supported files in root_dir and its subdirectories."""
    supported_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith(supported_exts):
                supported_files.append(os.path.join(dirpath, filename))
    return supported_files

def main():
    parser = argparse.ArgumentParser(
        description="Extract plain text from MHTML, DOCX, Markdown, or HTML files (recursively)."
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Path to a specific file to process (.mhtml, .docx, .md, .html)"
    )
    args = parser.parse_args()

    supported_exts = ('.mhtml', '.docx', '.md', '.html')

    if args.file:
        process_file(args.file)
        return

    # Recursively find all supported files
    files = find_supported_files('.', supported_exts)
    if not files:
        print("No supported files (.mhtml, .docx, .md, .html) found in the current directory or subdirectories.")
        return

    print(f"Found {len(files)} supported files in the current directory and subdirectories:")
    for f in files:
        print(f" - {f}")
    response = input("Do you want to process all of these files? (yes/no): ").strip().lower()

    if response in ('yes', 'y'):
        for file_path in files:
            process_file(file_path)
    else:
        print("No files were processed.")

if __name__ == "__main__":
    main()